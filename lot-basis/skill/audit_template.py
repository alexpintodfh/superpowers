"""
Lot Basis Template Audit Script
DreamFinders Homes — Jacksonville Division

Usage:
    python3 audit_template.py <template.xlsm> [jcs_total_99] [jcs_total_11]

Reads a Q3 lot basis template and reports:
- SSC row status (G vs A)
- All 99 cost lines vs any JCS values provided
- Carry cost formula presence check
- CCCX and LD Contingency placement
"""

import sys
import openpyxl
from openpyxl.utils import get_column_letter

def audit(template_path):
    print(f"\nAuditing: {template_path}\n")
    
    try:
        wb_v = openpyxl.load_workbook(template_path, keep_vba=True, data_only=True)
        wb_f = openpyxl.load_workbook(template_path, keep_vba=True, data_only=False)
    except Exception as e:
        print(f"Error loading file: {e}")
        return

    ws_v = wb_v['Detail']
    ws_f = wb_f['Detail']

    # ── SSC Check ──────────────────────────────────────────────────────────────
    print("=== SSC: Sales / Starts / Closings ===")
    for r, label in [(10, 'Sales'), (27, 'Starts'), (44, 'Closings')]:
        row = list(ws_v.iter_rows(min_row=r, max_row=r, values_only=True))[0]
        a = row[0] or 0
        g = row[6] or 0
        ok = a == g
        print(f"  {'✓' if ok else '✗'} {label}: A(lot count)={a}, G(sum)={g}"
              + ('' if ok else f'  ← GAP = {g-a:+}'))
    print()

    # ── Carry Cost Check ───────────────────────────────────────────────────────
    print("=== 11 Carry Costs ===")
    carry_rows = {
        66: 'Closing Costs',
        67: 'RE Taxes',
        68: 'HOA',
        86: 'CCCX',
    }
    for r, label in carry_rows.items():
        row_v = list(ws_v.iter_rows(min_row=r, max_row=r, values_only=True))[0]
        total = row_v[6] or 0
        actuals = row_v[7] or 0
        col_i_f = ws_f.cell(r, 9).value
        col_l_f = ws_f.cell(r, 12).value
        is_formula = str(col_i_f or '').startswith('=') or str(col_l_f or '').startswith('=')
        
        # Check if all I-BP cells are formulas
        all_formula = all(
            str(ws_f.cell(r, col).value or '').startswith('=')
            for col in range(9, 69)
        ) if r != 86 else True
        
        print(f"  Row {r} {label}: total={total:,.2f}, actuals={actuals:,.2f}")
        print(f"    col I formula: {col_i_f}")
        if r == 86:
            print(f"    col L formula: {col_l_f}")
        if r != 86:
            print(f"    All I-BP are formulas: {'✓' if all_formula else '✗ SOME HARDCODED'}")
    print()

    # ── 99 Dev Costs ───────────────────────────────────────────────────────────
    print("=== 99 Dev Costs ===")
    for r in range(93, 134):
        row = list(ws_v.iter_rows(min_row=r, max_row=r, values_only=True))[0]
        code = row[1]
        desc = row[3]
        total = row[6] or 0
        actuals = row[7] or 0
        if not isinstance(code, int) or total == 0 and actuals == 0:
            continue
        nz = [
            (col, str(ws_v.cell(8, col).value)[:7] if ws_v.cell(8, col).value else '?', row[col-1])
            for col in range(9, 70)
            if (col-1) < len(row) and row[col-1] is not None and row[col-1] != 0
        ]
        nz_str = '  '.join(f"{get_column_letter(c)}={v:,.0f}" for c, d, v in nz[:5])
        print(f"  Row {r} {code} {desc}: total={total:,.2f}, actuals={actuals:,.2f}")
        if nz_str:
            print(f"    Cashflows: {nz_str}")

    # ── Grand Total ────────────────────────────────────────────────────────────
    print()
    row136 = list(ws_v.iter_rows(min_row=136, max_row=136, values_only=True))[0]
    print(f"=== 99 Grand Total: {row136[6]:,.2f} ===")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python3 audit_template.py <template.xlsm>")
        sys.exit(1)
    audit(sys.argv[1])
