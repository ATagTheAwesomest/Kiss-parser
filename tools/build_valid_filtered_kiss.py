from __future__ import annotations

import re
import sys
from pathlib import Path

# Usage: python build_valid_filtered_kiss.py [job folder]   (defaults to the current folder)
root = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else Path.cwd()
allowed_path = root / 'assemblies.txt'

# Load allowed assembly IDs from the whitelist file.
allowed = set()
if allowed_path.exists():
    text = allowed_path.read_text(encoding='utf-8', errors='ignore')
    for raw in re.split(r'[,;|\s]+', text):
        token = raw.strip().strip('"\'[](){}')
        if token:
            allowed.add(token)
else:
    raise FileNotFoundError(f'Whitelist file not found: {allowed_path}')

# Source files are all real KSS detail exports, excluding generated and filtered outputs.
source_files = sorted(
    p for p in root.glob('*.kss')
    if 'Filtered Assemblies' not in p.name and '(Generated)' not in p.name
)

if not source_files:
    raise FileNotFoundError('No source KSS files found in the folder.')

project_name = 'HYPER-WARRENTON-B2'
job_number = '26-418'
header_line = None
current_assm = None
assembly_rows: dict[str, dict[tuple, dict]] = {}
drawings: dict[str, str] = {}

for path in source_files:
    for line in path.read_text(encoding='utf-8', errors='ignore').splitlines():
        if not line or not line.strip():
            continue
        parts = line.split(',')
        kind = parts[0].strip().upper()

        if kind == 'H' and len(parts) >= 3:
            if header_line is None:
                job_number = parts[1].strip() or job_number
                project_name = parts[2].strip() or project_name
                header_line = line.strip()
            continue

        if kind == 'D' and len(parts) >= 12:
            assembly = parts[1].strip()
            if not assembly or assembly not in allowed:
                current_assm = assembly or current_assm
                continue

            qty = 0
            try:
                qty = int(parts[5].strip())
            except ValueError:
                current_assm = assembly
                continue

            weight = 0.0
            for value in reversed(parts):
                text = value.strip()
                if not text:
                    continue
                try:
                    weight = float(text)
                    break
                except ValueError:
                    pass

            main_mark = parts[3].strip()
            raw_part = parts[4].strip()
            shape = parts[6].strip()
            size = parts[7].strip()
            grade = parts[8].strip()
            length = parts[9].strip()
            finish = parts[10].strip() if len(parts) > 10 else ''
            description = parts[11].strip() if len(parts) > 11 else ''
            key = (main_mark, raw_part, shape, size, grade, length, finish, description)

            assembly_map = assembly_rows.setdefault(assembly, {})
            rec = assembly_map.setdefault(key, {
                'assembly': assembly,
                'main_mark': main_mark,
                'raw_part': raw_part,
                'shape': shape,
                'size': size,
                'grade': grade,
                'length': length,
                'finish': finish,
                'description': description,
                'qty': 0,
                'weight': weight,
            })
            rec['qty'] += qty
            rec['weight'] = weight
            current_assm = assembly
            continue

        if kind == 'A' and len(parts) >= 2 and current_assm and current_assm in allowed:
            drawings[current_assm] = parts[1].strip()
            continue

        if kind in {'S', '*'}:
            current_assm = None
            continue

# Build the output file.
out_path = root / f'{project_name} - Job {job_number} - Filtered Assemblies - VALID.kss'
lines: list[str] = []
lines.append('KISS,1.1,Tekla Structures')
if header_line:
    lines.append(header_line)
else:
    lines.append(f'H,{job_number},{project_name}')

for assembly in sorted(assembly_rows.keys()):
    rows = assembly_rows[assembly]
    total_qty = sum(item['qty'] for item in rows.values())
    prefix_digits = re.match(r'^(\d+)', assembly)
    prefix = prefix_digits.group(1)[:4] if prefix_digits else assembly

    for item in sorted(rows.values(), key=lambda r: (r['raw_part'] or r['main_mark'], r['shape'], r['size'])):
        desc = item['description'].replace(',', ';')
        weight = item['weight']
        d_line = (
            f"D,{assembly},0,{item['main_mark']},{item['raw_part']},{item['qty']},"
            f"{item['shape']},{item['size']},{item['grade']},{item['length']},{item['finish']},"
            f"{desc},,{weight}"
        )
        lines.append(d_line)

    lines.append(f'S,{prefix},{total_qty}')
    if assembly in drawings and drawings[assembly]:
        lines.append(f'A,{drawings[assembly]}')
    lines.append('*')

out_path.write_text('\n'.join(lines) + '\n', encoding='ascii')
print(f'Wrote valid filtered KISS export: {out_path}')
print(f'Assemblies kept: {len(assembly_rows)}')
print(f'Files used: {len(source_files)}')
