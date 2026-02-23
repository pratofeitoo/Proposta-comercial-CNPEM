#!/usr/bin/env python3
"""
Convert CNPEM - PCD.timeline.csv to XLSX and add an "Estimated Hours" column.
"""
from pathlib import Path
import pandas as pd

csv_path = Path(__file__).parent / 'CNPEM - PCD.timeline.csv'
if not csv_path.exists():
    raise SystemExit(f"CSV not found: {csv_path}")

df = pd.read_csv(csv_path)

# Simple heuristic estimates (hours) by milestone keywords
def estimate_hours(row):
    m = str(row.get('Milestone', '')).lower()
    if 'kickoff' in m or 'alinhamento' in m:
        return 8
    if 'diagn' in m:
        return 120
    if 'análise' in m or 'analise' in m or 'prior' in m:
        return 40
    if 'capacita' in m:
        return 120
    if 'implement' in m or 'implementa' in m:
        return 160
    if 'consol' in m or 'relat' in m:
        return 60
    return 40

df['Estimated Hours'] = df.apply(estimate_hours, axis=1)

xlsx_path = csv_path.with_suffix('.xlsx')
df.to_excel(xlsx_path, index=False)
print(f"Wrote {xlsx_path}")
