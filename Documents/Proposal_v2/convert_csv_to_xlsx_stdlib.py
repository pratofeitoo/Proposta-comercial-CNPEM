#!/usr/bin/env python3
"""
Convert CSV timeline to XLSX using only Python standard library (no external dependencies).
Adds an 'Estimated Hours' column using simple heuristics.
"""
import csv
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.sax.saxutils as sax

# Helpers
COL_LETTERS = [
    'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'
]

def col_letter(n):
    # 0-based n
    s = ''
    while True:
        s = COL_LETTERS[n % 26] + s
        n = n // 26 - 1
        if n < 0:
            break
    return s

# Estimate heuristic
def estimate_hours_for_milestone(milestone_text):
    m = (milestone_text or '').lower()
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

# Paths
csv_path = Path(__file__).parent / 'CNPEM - PCD.timeline.csv'
if not csv_path.exists():
    raise SystemExit(f"CSV not found: {csv_path}")

xlsx_path = csv_path.with_suffix('.xlsx')

# Read CSV and augment
rows = []
with csv_path.open(newline='', encoding='utf-8-sig') as f:
    reader = csv.reader(f)
    header = next(reader)
    for r in reader:
        rows.append(r)

# Build new header and data rows
new_header = header + ['Estimated Hours']
new_rows = []
for r in rows:
    milestone = r[0] if len(r) > 0 else ''
    est = estimate_hours_for_milestone(milestone)
    new_rows.append(r + [str(est)])

# Build sheet1.xml content
def make_sheet_xml(header, data_rows):
    sb = []
    sb.append('<?xml version="1.0" encoding="UTF-8"?>')
    sb.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"')
    sb.append(' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    sb.append('<sheetData>')
    # Header row (row 1)
    rnum = 1
    sb.append(f'<row r="{rnum}">')
    for ci, col in enumerate(header):
        cell_ref = f"{col_letter(ci)}{rnum}"
        text = sax.escape(col)
        sb.append(f'<c r="{cell_ref}" t="inlineStr"><is><t>{text}</t></is></c>')
    sb.append('</row>')
    # Data rows
    for i, row in enumerate(data_rows, start=2):
        sb.append(f'<row r="{i}">')
        for ci, val in enumerate(row):
            cell_ref = f"{col_letter(ci)}{i}"
            # Try numeric
            try:
                if val is None or val == '':
                    sb.append(f'<c r="{cell_ref}"/>')
                else:
                    num = float(val)
                    # integer? keep as number
                    if num.is_integer():
                        sb.append(f'<c r="{cell_ref}"><v>{int(num)}</v></c>')
                    else:
                        sb.append(f'<c r="{cell_ref}"><v>{num}</v></c>')
            except Exception:
                s = sax.escape(str(val))
                sb.append(f'<c r="{cell_ref}" t="inlineStr"><is><t>{s}</t></is></c>')
        sb.append('</row>')
    sb.append('</sheetData>')
    sb.append('</worksheet>')
    return '\n'.join(sb)

sheet_xml = make_sheet_xml(new_header, new_rows)

# Other required xml parts
content_types = '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''

rels_rels = '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

workbook_rels = '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

workbook_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Timeline" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>'''

styles_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="0"/>
  <fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>'''

# Write zip
with ZipFile(xlsx_path, 'w', compression=ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', content_types)
    z.writestr('_rels/.rels', rels_rels)
    z.writestr('xl/workbook.xml', workbook_xml)
    z.writestr('xl/_rels/workbook.xml.rels', workbook_rels)
    z.writestr('xl/worksheets/sheet1.xml', sheet_xml)
    z.writestr('xl/styles.xml', styles_xml)

print(f'Wrote {xlsx_path}')
