"""Automates repetitive "Production/Shipping Entry" data entry with a small PySide6 UI.

Pick the screen positions of the Piece Mark box, Quantity box, and Add button
once (saved to entry_config.json), paste/load a list of piece marks, then hit
Start: the tool clicks the Piece Mark box, types each value, clicks the
Quantity box, types the quantity, clicks Add, and repeats for the next line.

PyAutoGUI's fail-safe is left on: drag the mouse to a screen corner to abort.
A 3-second countdown before each run gives you time to click back onto the
target window.
"""

from __future__ import annotations

import ctypes
import json
import os
import re
import sys
import threading
import time
from collections import Counter
from pathlib import Path
from typing import Optional

try:
    from pynput import keyboard as pynput_keyboard
except Exception:  # pragma: no cover - optional dependency fallback
    pynput_keyboard = None

# Must happen before importing PySide6/pyautogui so Qt and pyautogui agree on
# physical-pixel coordinates (otherwise HiDPI scaling makes clicks drift).
if os.name == "nt":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # per-monitor DPI aware
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "0")
    os.environ.setdefault("QT_SCALE_FACTOR", "1")

import pyautogui
from PySide6.QtCore import QEventLoop, QPoint, QRect, QRectF, QThread, QTimer, Qt, Signal
from PySide6.QtGui import QColor, QFont, QGuiApplication, QPainter, QPen
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "entry_config.json"


def _virtual_desktop_rect() -> QRect:
    rect = QRect()
    for screen in QGuiApplication.screens():
        rect = rect.united(screen.geometry())
    return rect


def _load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {
        "piece_mark_pos": None,
        "quantity_pos": None,
        "add_button_pos": None,
        "delay_ms": 400,
        "default_quantity": 1,
    }


def _save_config(cfg: dict) -> None:
    CONFIG_PATH.write_text(json.dumps(cfg, indent=2), encoding="utf-8")


_BRACKETED_TEXT_RE = re.compile(r"\[[^\]]*\]|\{[^}]*\}|\([^)]*\)")


def _parse_piece_marks(text: str) -> list[str]:
    """Return piece marks while ignoring bracketed extra information."""
    marks = []
    for raw_line in text.splitlines():
        line = _BRACKETED_TEXT_RE.sub("", raw_line).strip()
        if line:
            marks.append(line)
    return marks


def _build_entry_plan(piece_marks, default_quantity):
    """Return unique piece marks with quantity equal to the number of times that
    mark appears in the list, falling back to the default quantity when it only
    appears once."""
    counts = Counter(piece_marks)
    ordered_unique = []
    seen = set()
    for mark in piece_marks:
        if mark not in seen:
            ordered_unique.append(mark)
            seen.add(mark)
    return [(mark, counts[mark] if counts[mark] > 1 else default_quantity) for mark in ordered_unique]


class SelectionOverlay(QWidget):
    """Full-desktop transparent overlay used to pick a single point by click."""

    result_ready = Signal(object)

    def __init__(self, instructions: str):
        super().__init__()
        self.instructions = instructions
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setCursor(Qt.CrossCursor)
        self.setGeometry(_virtual_desktop_rect())

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.result_ready.emit(None)
            self.close()
        else:
            super().keyPressEvent(event)

    def mousePressEvent(self, event):
        if event.button() != Qt.LeftButton:
            return
        self.result_ready.emit(event.globalPosition().toPoint())
        self.close()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(QRectF(self.rect()), QColor(0, 0, 0, 90))
        painter.setPen(QPen(Qt.white))
        painter.setFont(QFont("Segoe UI", 12, QFont.Bold))
        painter.drawText(self.rect().adjusted(0, 24, 0, 0), Qt.AlignHCenter | Qt.AlignTop, self.instructions)

    @staticmethod
    def pick(instructions: str) -> Optional[QPoint]:
        overlay = SelectionOverlay(instructions)
        loop = QEventLoop()
        captured: dict = {}

        def on_result(value):
            captured["value"] = value
            loop.quit()

        overlay.result_ready.connect(on_result)
        overlay.show()
        overlay.activateWindow()
        loop.exec()
        return captured.get("value")


class EntryWorker(QThread):
    log = Signal(str)
    progress = Signal(int, int)
    current_row = Signal(int)
    paused = Signal()
    resumed = Signal()
    finished_ok = Signal()
    stopped = Signal()
    failed = Signal(str)

    def __init__(self, entries, piece_mark_pos, quantity_pos, add_pos, delay_ms, countdown_s, start_index=0):
        super().__init__()
        self.entries = entries
        self.piece_mark_pos = piece_mark_pos
        self.quantity_pos = quantity_pos
        self.add_pos = add_pos
        self.delay_ms = delay_ms
        self.countdown_s = countdown_s
        self.start_index = max(0, min(start_index, len(entries) - 1 if entries else 0))
        self._stop_requested = False
        self._pause_requested = False
        self._pause_event = threading.Event()
        self._pause_event.set()

    def is_paused(self) -> bool:
        return self._pause_requested

    def pause(self):
        if self._pause_requested:
            return
        self._pause_requested = True
        self._pause_event.clear()

    def resume(self):
        if not self._pause_requested:
            return
        self._pause_requested = False
        self._pause_event.set()
        self.resumed.emit()

    def stop(self):
        self._stop_requested = True

    def run(self):
        try:
            for remaining in range(self.countdown_s, 0, -1):
                if self._stop_requested:
                    self.log.emit("Stopped before starting.")
                    self.stopped.emit()
                    return
                while self._pause_requested:
                    self.paused.emit()
                    self._pause_event.wait(0.1)
                    if self._stop_requested:
                        self.log.emit("Stopped before starting.")
                        self.stopped.emit()
                        return
                self.log.emit(f"Starting in {remaining}...")
                time.sleep(1)

            delay = self.delay_ms / 1000.0
            total = len(self.entries)
            for index in range(self.start_index, total):
                if self._stop_requested:
                    self.log.emit("Stopped.")
                    self.stopped.emit()
                    return

                while self._pause_requested:
                    self.paused.emit()
                    self._pause_event.wait(0.1)
                    if self._stop_requested:
                        self.log.emit("Stopped.")
                        self.stopped.emit()
                        return

                self.current_row.emit(index)
                mark, qty = self.entries[index]

                pyautogui.click(self.piece_mark_pos[0], self.piece_mark_pos[1])
                time.sleep(delay)
                pyautogui.hotkey("ctrl", "a")
                pyautogui.typewrite(mark, interval=0.01)
                time.sleep(delay)

                pyautogui.click(self.quantity_pos[0], self.quantity_pos[1])
                time.sleep(delay)
                pyautogui.hotkey("ctrl", "a")
                pyautogui.typewrite(str(qty), interval=0.01)
                time.sleep(delay)

                pyautogui.click(self.add_pos[0], self.add_pos[1])
                time.sleep(delay)

                self.log.emit(f"[{index + 1}/{total}] Added {mark} (qty {qty})")
                self.progress.emit(index + 1, total)

            self.finished_ok.emit()
        except pyautogui.FailSafeException:
            self.failed.emit("Aborted: mouse moved to a screen corner (fail-safe).")
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Piece Mark Entry Automator")
        self.cfg = _load_config()
        self.worker: Optional[EntryWorker] = None
        self.entries = []
        self._global_escape_listener = None

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        pos_group = QGroupBox("Screen positions")
        pos_form = QFormLayout(pos_group)

        self.piece_mark_label = QLabel()
        self.quantity_label = QLabel()
        self.add_button_label = QLabel()

        pos_form.addRow(self._pos_row("Piece Mark box:", self.piece_mark_label, self._pick_piece_mark))
        pos_form.addRow(self._pos_row("Quantity box:", self.quantity_label, self._pick_quantity))
        pos_form.addRow(self._pos_row("Add button:", self.add_button_label, self._pick_add_button))
        layout.addWidget(pos_group)

        settings_group = QGroupBox("Settings")
        settings_form = QFormLayout(settings_group)

        self.quantity_spin = QSpinBox()
        self.quantity_spin.setRange(1, 999999)
        self.quantity_spin.setValue(int(self.cfg.get("default_quantity", 1)))
        settings_form.addRow("Default quantity:", self.quantity_spin)

        self.delay_spin = QSpinBox()
        self.delay_spin.setRange(50, 5000)
        self.delay_spin.setSingleStep(50)
        self.delay_spin.setSuffix(" ms")
        self.delay_spin.setValue(int(self.cfg.get("delay_ms", 400)))
        settings_form.addRow("Delay between actions:", self.delay_spin)

        self.countdown_spin = QSpinBox()
        self.countdown_spin.setRange(0, 15)
        self.countdown_spin.setSuffix(" s")
        self.countdown_spin.setValue(3)
        settings_form.addRow("Start countdown:", self.countdown_spin)

        layout.addWidget(settings_group)

        list_group = QGroupBox("Piece marks (one per line)")
        list_layout = QVBoxLayout(list_group)
        self.marks_edit = QPlainTextEdit()
        self.marks_edit.setPlaceholderText("55311021\n55311022\n55312000\n...")
        self.marks_edit.textChanged.connect(self._refresh_queue_table)
        list_layout.addWidget(self.marks_edit)
        layout.addWidget(list_group)

        self.start_row_spin = QSpinBox()
        self.start_row_spin.setRange(1, 1)
        self.start_row_spin.setValue(1)
        self.start_row_spin.setEnabled(True)
        start_row_widget = QWidget()
        start_row_layout = QHBoxLayout(start_row_widget)
        start_row_layout.setContentsMargins(0, 0, 0, 0)
        start_row_layout.addWidget(QLabel("Start from row:"))
        start_row_layout.addWidget(self.start_row_spin)
        layout.addWidget(start_row_widget)

        self.entries_table = QTableWidget(0, 4)
        self.entries_table.setHorizontalHeaderLabels(["Line", "Piece Mark", "Qty", "Status"])
        self.entries_table.verticalHeader().setVisible(False)
        self.entries_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.entries_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.entries_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.entries_table.setAlternatingRowColors(True)
        self.entries_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.entries_table.setMinimumHeight(220)
        layout.addWidget(self.entries_table)

        button_row = QHBoxLayout()
        self.start_button = QPushButton("Start (F5)")
        self.start_button.clicked.connect(self._start)
        self.pause_button = QPushButton("Pause")
        self.pause_button.clicked.connect(self._pause)
        self.pause_button.setEnabled(False)
        self.resume_button = QPushButton("Resume")
        self.resume_button.clicked.connect(self._resume)
        self.resume_button.setEnabled(False)
        self.stop_button = QPushButton("Stop")
        self.stop_button.clicked.connect(self._stop)
        self.stop_button.setEnabled(False)
        button_row.addWidget(self.start_button)
        button_row.addWidget(self.pause_button)
        button_row.addWidget(self.resume_button)
        button_row.addWidget(self.stop_button)
        layout.addLayout(button_row)

        self.status_label = QLabel("Ready.")
        layout.addWidget(self.status_label)

        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setMaximumBlockCount(500)
        layout.addWidget(self.log_edit)

        self._refresh_pos_labels()
        self._refresh_queue_table()
        self._install_global_escape_listener()
        self.resize(700, 760)

    def _pos_row(self, caption: str, value_label: QLabel, pick_slot) -> QWidget:
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.addWidget(value_label, 1)
        pick_button = QPushButton("Pick...")
        pick_button.clicked.connect(pick_slot)
        row_layout.addWidget(pick_button)
        wrapper = QWidget()
        wrapper_layout = QHBoxLayout(wrapper)
        wrapper_layout.setContentsMargins(0, 0, 0, 0)
        wrapper_layout.addWidget(QLabel(caption))
        wrapper_layout.addWidget(row, 1)
        return wrapper

    def _refresh_pos_labels(self):
        self.piece_mark_label.setText(self._fmt_pos(self.cfg.get("piece_mark_pos")))
        self.quantity_label.setText(self._fmt_pos(self.cfg.get("quantity_pos")))
        self.add_button_label.setText(self._fmt_pos(self.cfg.get("add_button_pos")))

    def _refresh_queue_table(self):
        marks = _parse_piece_marks(self.marks_edit.toPlainText())
        self.entries = _build_entry_plan(marks, self.quantity_spin.value())
        self.entries_table.setRowCount(len(self.entries))
        for row, (mark, qty) in enumerate(self.entries):
            self.entries_table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
            self.entries_table.setItem(row, 1, QTableWidgetItem(mark))
            self.entries_table.setItem(row, 2, QTableWidgetItem(str(qty)))
            self.entries_table.setItem(row, 3, QTableWidgetItem("Queued"))
        if self.entries:
            max_row = len(self.entries)
            self.start_row_spin.setRange(1, max_row)
            self.start_row_spin.setValue(min(self.start_row_spin.value(), max_row))
        else:
            self.start_row_spin.setRange(1, 1)
            self.start_row_spin.setValue(1)
        self.entries_table.resizeColumnsToContents()
        self.entries_table.selectRow(max(0, min(self.start_row_spin.value() - 1, max(0, self.entries_table.rowCount() - 1))))

    @staticmethod
    def _fmt_pos(pos) -> str:
        return f"({pos[0]}, {pos[1]})" if pos else "not set"

    def _pick_piece_mark(self):
        self._pick_into("piece_mark_pos", "Click the Piece Mark text box")

    def _pick_quantity(self):
        self._pick_into("quantity_pos", "Click the Quantity text box")

    def _pick_add_button(self):
        self._pick_into("add_button_pos", "Click the Add button")

    def _pick_into(self, key: str, instructions: str):
        self.hide()
        QApplication.processEvents()
        time.sleep(0.2)  # let the main window actually disappear before overlay grabs the screen
        point = SelectionOverlay.pick(instructions)
        self.show()
        if point is not None:
            self.cfg[key] = [point.x(), point.y()]
            _save_config(self.cfg)
            self._refresh_pos_labels()

    def _install_global_escape_listener(self):
        if pynput_keyboard is None:
            return
        self._global_escape_listener = pynput_keyboard.Listener(on_press=self._handle_global_escape)
        self._global_escape_listener.start()

    def _handle_global_escape(self, key):
        if key == pynput_keyboard.Key.esc and self.worker and self.worker.isRunning():
            if self.worker.is_paused():
                QTimer.singleShot(0, self._resume)
            else:
                QTimer.singleShot(0, self._pause)

    def _start(self):
        marks = [line.strip() for line in self.marks_edit.toPlainText().splitlines() if line.strip()]
        if not marks:
            QMessageBox.warning(self, "No piece marks", "Paste at least one piece mark, one per line.")
            return
        for key, caption in (
            ("piece_mark_pos", "Piece Mark box"),
            ("quantity_pos", "Quantity box"),
            ("add_button_pos", "Add button"),
        ):
            if not self.cfg.get(key):
                QMessageBox.warning(self, "Missing position", f"Pick the {caption} position first.")
                return

        self.cfg["default_quantity"] = self.quantity_spin.value()
        self.cfg["delay_ms"] = self.delay_spin.value()
        _save_config(self.cfg)

        self._refresh_queue_table()
        start_index = max(0, min(self.start_row_spin.value() - 1, len(self.entries) - 1)) if self.entries else 0
        self.worker = EntryWorker(
            entries=self.entries,
            piece_mark_pos=tuple(self.cfg["piece_mark_pos"]),
            quantity_pos=tuple(self.cfg["quantity_pos"]),
            add_pos=tuple(self.cfg["add_button_pos"]),
            delay_ms=self.delay_spin.value(),
            countdown_s=self.countdown_spin.value(),
            start_index=start_index,
        )
        self.worker.log.connect(self.log_edit.appendPlainText)
        self.worker.progress.connect(self._on_progress)
        self.worker.current_row.connect(self._on_current_row)
        self.worker.paused.connect(self._on_paused)
        self.worker.resumed.connect(self._on_resumed)
        self.worker.finished_ok.connect(self._on_finished)
        self.worker.stopped.connect(self._on_stopped)
        self.worker.failed.connect(self._on_failed)

        self.start_button.setEnabled(False)
        self.pause_button.setEnabled(True)
        self.resume_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.marks_edit.setEnabled(False)
        self.start_row_spin.setEnabled(False)
        self.quantity_spin.setEnabled(False)
        self.status_label.setText(f"Starting from row {start_index + 1}...")
        self._set_queue_statuses(start_index, "Running")
        self.worker.start()

    def _set_queue_statuses(self, current_index: Optional[int] = None, current_label: str = "Queued"):
        if not self.entries:
            return
        for row in range(self.entries_table.rowCount()):
            if current_index is not None and row == current_index:
                self.entries_table.setItem(row, 3, QTableWidgetItem(current_label))
            elif current_index is not None and row < current_index:
                self.entries_table.setItem(row, 3, QTableWidgetItem("Done"))
            else:
                self.entries_table.setItem(row, 3, QTableWidgetItem("Queued"))
        if current_index is not None and 0 <= current_index < self.entries_table.rowCount():
            self.entries_table.selectRow(current_index)

    def _pause(self):
        if self.worker and self.worker.isRunning() and not self.worker.is_paused():
            self.worker.pause()
            self.pause_button.setEnabled(False)
            self.resume_button.setEnabled(True)
            self.status_label.setText("Paused — press Esc or Resume.")

    def _resume(self):
        if self.worker and self.worker.isRunning() and self.worker.is_paused():
            self.worker.resume()
            self.pause_button.setEnabled(True)
            self.resume_button.setEnabled(False)
            self.status_label.setText("Resumed.")

    def _stop(self):
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.pause_button.setEnabled(False)
            self.resume_button.setEnabled(False)
            self.stop_button.setEnabled(False)
        else:
            self.start_button.setEnabled(True)
            self.pause_button.setEnabled(False)
            self.resume_button.setEnabled(False)
            self.stop_button.setEnabled(False)

    def _on_current_row(self, row_index: int):
        self._set_queue_statuses(row_index, "Running")

    def _on_paused(self):
        self.pause_button.setEnabled(False)
        self.resume_button.setEnabled(True)
        self.status_label.setText("Paused — press Esc or Resume.")

    def _on_resumed(self):
        self.pause_button.setEnabled(True)
        self.resume_button.setEnabled(False)
        self.status_label.setText("Resumed.")

    def _on_progress(self, row_number: int, total: int):
        self.status_label.setText(f"Row {row_number} of {total}")
        self._set_queue_statuses(max(0, row_number - 1), "Done")

    def _on_finished(self):
        self.status_label.setText("Done.")
        self.start_button.setEnabled(True)
        self.pause_button.setEnabled(False)
        self.resume_button.setEnabled(False)
        self.stop_button.setEnabled(False)
        self.marks_edit.setEnabled(True)
        self.start_row_spin.setEnabled(True)
        self.quantity_spin.setEnabled(True)
        self._set_queue_statuses(None, "Done")
        if self.entries_table.rowCount():
            self.entries_table.selectRow(self.entries_table.rowCount() - 1)

    def _on_stopped(self):
        self.status_label.setText("Stopped.")
        self.start_button.setEnabled(True)
        self.pause_button.setEnabled(False)
        self.resume_button.setEnabled(False)
        self.stop_button.setEnabled(False)
        self.marks_edit.setEnabled(True)
        self.start_row_spin.setEnabled(True)
        self.quantity_spin.setEnabled(True)

    def _on_failed(self, message: str):
        self.status_label.setText("Stopped (error).")
        self.log_edit.appendPlainText(f"ERROR: {message}")
        self.start_button.setEnabled(True)
        self.pause_button.setEnabled(False)
        self.resume_button.setEnabled(False)
        self.stop_button.setEnabled(False)
        self.marks_edit.setEnabled(True)
        self.start_row_spin.setEnabled(True)
        self.quantity_spin.setEnabled(True)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_F5 and self.start_button.isEnabled():
            self._start()
            return
        if event.key() == Qt.Key_Escape:
            if self.worker and self.worker.isRunning():
                if self.worker.is_paused():
                    self._resume()
                else:
                    self._pause()
                return
        super().keyPressEvent(event)

    def closeEvent(self, event):
        if self._global_escape_listener is not None and self._global_escape_listener.is_alive():
            self._global_escape_listener.stop()
        super().closeEvent(event)


def main():
    pyautogui.FAILSAFE = True
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
