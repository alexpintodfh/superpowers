#!/usr/bin/env python3
"""Apply coded invoices to a draw sheet and rebuild Exhibit D.

Reads invoices.json (produced by the agent from the rendered invoice images),
then on the target 'Draw Request N' sheet:
  * zeroes the Current Period + Retainage input columns on every cost row,
  * writes each invoice's GROSS into the Current Period cell of its cost-code row
    (summing when several invoices share a code) and sets the Retainage rate, so
    the 'Current Period Less Retainage' column resolves to the net,
  * routes retention-release invoices to the matching '(No Retainage)' row so
    retainage is never deducted twice,
  * rebuilds both the standalone 'Exhibit D' sheet and the Exhibit D block
    embedded in the draw sheet (one line per invoice, at net).

Draw total = sum of net amounts = the figure requested from the land banker.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter as col_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
import draw_lib as dl


def _clear_inputs(ws, cols, code_map):
    rows = set()
    for entry in code_map.values():
        rows.update(v for v in entry.values() if v)
    for r in rows:
        ws.cell(row=r, column=cols["cp"]).value = 0
        ws.cell(row=r, column=cols["ret"]).value = 0


def _apply_to_rows(ws, cols, code_map, invoices, warnings):
    """Accumulate gross/rate per target row."""
    acc: dict[int, dict] = {}
    for inv in invoices:
        code = inv["major_code"]
        entry = code_map.get(code)
        if entry is None:
            warnings.append(f"major code {code} ({inv['exhibit_line']}) has no row in the tracker")
            continue
        if inv["is_retention_release"]:
            row = entry.get("no_ret") or entry.get("main")
            amount, rate = inv["net_amount"], 0.0
            if not entry.get("no_ret"):
                warnings.append(
                    f"code {code} retention release routed to main row "
                    f"(no '(No Retainage)' row exists)")
        else:
            row = entry.get("main")
            amount, rate = inv["gross_amount"], inv["retainage_rate"]
        if row is None:
            warnings.append(f"code {code} has no target row for {inv['exhibit_line']}")
            continue
        slot = acc.setdefault(row, {"gross": 0.0, "rate": None})
        slot["gross"] += amount
        if slot["rate"] is None:
            slot["rate"] = rate
        elif abs(slot["rate"] - rate) > 1e-9:
            warnings.append(
                f"row {row} received invoices with different retainage rates "
                f"({slot['rate']} vs {rate}); using {slot['rate']}")
    for row, slot in acc.items():
        ws.cell(row=row, column=cols["cp"]).value = round(slot["gross"], 2)
        ws.cell(row=row, column=cols["ret"]).value = slot["rate"] or 0
        # keep the canonical summary formulas intact / refreshed
        for role, formula in dl.summary_formulas(row, cols).items():
            if role in ("cplr", "tcd", "pct", "rtd", "bal"):
                ws.cell(row=row, column=cols[role]).value = formula


def _find_embedded_exhibit(ws):
    for r in range(1, ws.max_row + 1):
        if dl._norm(ws.cell(row=r, column=2).value) == 'exhibit "d"':
            return r
    return None


def _write_exhibit_rows(ws, start_row, cols_map, invoices, draw_number):
    """Shared writer for both the standalone sheet and the embedded block.

    cols_map maps logical fields to column indices:
      idx, code, desc, amount, lien   (and draw_cell as (row,col) for the draw #)
    """
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    money = dl.MONEY_FMT
    for i, inv in enumerate(invoices):
        r = start_row + i
        ws.cell(row=r, column=cols_map["idx"], value=i + 1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=cols_map["code"], value=inv["major_code"]).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=cols_map["desc"], value=inv["exhibit_line"]).alignment = Alignment(horizontal="center")
        amt = ws.cell(row=r, column=cols_map["amount"], value=round(inv["net_amount"], 2))
        amt.number_format = money
        ws.cell(row=r, column=cols_map["lien"], value=inv.get("lien_release", "N/A")).alignment = Alignment(horizontal="center")
        for key in ("idx", "code", "desc", "amount", "lien"):
            ws.cell(row=r, column=cols_map[key]).border = border
    total_row = start_row + len(invoices)
    tot_label = ws.cell(row=total_row, column=cols_map["idx"], value="TOTAL")
    tot_label.font = Font(bold=True)
    amt_col = col_letter(cols_map["amount"])
    tot = ws.cell(row=total_row, column=cols_map["amount"],
                  value=f"=SUM({amt_col}{start_row}:{amt_col}{total_row - 1})")
    tot.number_format = money
    tot.font = Font(bold=True)
    for key in ("idx", "code", "desc", "amount", "lien"):
        ws.cell(row=total_row, column=cols_map[key]).border = border
    return total_row


def rebuild_standalone_exhibit(wb, invoices, draw_number):
    name = next((n for n in wb.sheetnames if dl._norm(n) == "exhibit d"), None)
    if name is None:
        return
    ws = wb[name]
    ws["E2"] = draw_number
    # clear old invoice + total rows (row 5 downward, cols A..E)
    for r in range(5, ws.max_row + 2):
        for c in range(1, 6):
            ws.cell(row=r, column=c).value = None
            ws.cell(row=r, column=c).border = Border()
    _write_exhibit_rows(
        ws, 5,
        {"idx": 1, "code": 2, "desc": 3, "amount": 4, "lien": 5},
        invoices, draw_number)


def rebuild_embedded_exhibit(ws, invoices, draw_number):
    head = _find_embedded_exhibit(ws)
    if head is None:
        return
    # header block: B(exhibit) / row+1 Invoices,MajorCode,Draw#,N / row+3 col headers
    draw_hdr = head + 1
    ws.cell(row=draw_hdr, column=6, value=draw_number)  # F col holds the number
    data_start = head + 4
    # clear existing embedded invoice + total rows (cols B..F)
    for r in range(data_start, ws.max_row + 2):
        first = ws.cell(row=r, column=2).value
        for c in range(2, 7):
            ws.cell(row=r, column=c).value = None
            ws.cell(row=r, column=c).border = Border()
        if dl._norm(first) == "total":
            break
    _write_exhibit_rows(
        ws, data_start,
        {"idx": 2, "code": 3, "desc": 4, "amount": 5, "lien": 6},
        invoices, draw_number)


def main():
    ap = argparse.ArgumentParser(description="Apply invoices.json to a draw sheet and rebuild Exhibit D.")
    ap.add_argument("workbook", help="Path to the tracker .xlsx (edited in place unless --out)")
    ap.add_argument("invoices_json", help="Path to invoices.json")
    ap.add_argument("--draw", type=int, help="Draw number to target (default: latest sheet)")
    ap.add_argument("--out", help="Write to this path instead of editing in place")
    args = ap.parse_args()

    data = dl.load_invoices(args.invoices_json)
    invoices = data["invoices"]
    wb = openpyxl.load_workbook(args.workbook)

    if args.draw:
        target = f"Draw Request {args.draw}"
        name = next((n for n in wb.sheetnames if dl._norm(n) == dl._norm(target)), None)
        if name is None:
            sys.exit(f"Sheet '{target}' not found. Run new_draw_sheet.py first.")
        ws, draw_number = wb[name], args.draw
    else:
        ws, draw_number = dl.latest_draw_sheet(wb)

    cols = dl.find_columns(ws)
    code_map = dl.build_code_map(ws)
    warnings: list[str] = []

    _clear_inputs(ws, cols, code_map)
    _apply_to_rows(ws, cols, code_map, invoices, warnings)
    rebuild_embedded_exhibit(ws, invoices, draw_number)
    rebuild_standalone_exhibit(wb, invoices, draw_number)

    out = args.out or args.workbook
    wb.save(out)

    net_total = round(sum(i["net_amount"] for i in invoices), 2)
    print(f"Draw Request {draw_number}: {len(invoices)} invoice(s) applied to sheet {ws.title!r}")
    print(f"Draw total (net / amount to draw): ${net_total:,.2f}")
    for w in warnings:
        print(f"  WARNING: {w}")
    print(f"Saved -> {out}")


if __name__ == "__main__":
    main()
