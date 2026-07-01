"""Shared helpers for the draw-request-package skill.

Everything that reads or mutates the tracker workbook flows through here so the
layout logic lives in one place. The tracker's summary columns shift one to the
right with every new draw, so we always locate columns by their header text in
row 4 rather than by fixed letters.
"""
from __future__ import annotations

import re
from openpyxl.utils import get_column_letter as col_letter

# Styling lifted from the real tracker so generated cells match the original.
ORANGE_FILL = "FFFFCC99"                     # the "Current Period" / "Retainage" input columns
MONEY_FMT = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
MONEY_FMT_INT = '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"??_);_(@_)'
PCT_FMT = "0%"

HEADER_ROW = 4
FIRST_COST_ROW = 7
LAST_COST_ROW = 49          # data rows live in 7..49; row 51 is TOTAL COSTS
CURRENT_BUDGET_COL = 6      # column F is always "Current Budget"
FIRST_DRAW_COL = 7          # column G is always "Draw 1"

# Header text -> the summary column we care about. Matched case-insensitively
# after collapsing whitespace.
SUMMARY_HEADERS = {
    "total previously drawn": "prev",
    "current period": "cp",
    "retainage": "ret",
    "current period less retainage": "cplr",
    "total completed to date": "tcd",
    "% complete": "pct",
    "retainage to date": "rtd",
    "balance to finish": "bal",
}


def _norm(text) -> str:
    return re.sub(r"\s+", " ", str(text or "").strip()).lower()


def find_columns(ws) -> dict:
    """Return {role: column_index} for the summary columns, located by header."""
    found = {}
    for c in range(1, ws.max_column + 1):
        key = _norm(ws.cell(row=HEADER_ROW, column=c).value)
        if key in SUMMARY_HEADERS:
            role = SUMMARY_HEADERS[key]
            # "retainage" also matches "retainage to date"; only take the first
            # bare "retainage" as the input column.
            if role == "ret" and "ret" in found:
                continue
            found.setdefault(role, c)
    missing = set(SUMMARY_HEADERS.values()) - set(found)
    if missing:
        raise ValueError(f"Could not locate summary columns {missing} in {ws.title!r}")
    return found


def last_draw_col(cols: dict) -> int:
    """Column index of the last frozen 'Draw N' history column (just left of prev)."""
    return cols["prev"] - 1


def latest_draw_sheet(wb):
    """Return (sheet, number) for the highest-numbered 'Draw Request N' sheet."""
    best = None
    for name in wb.sheetnames:
        m = re.fullmatch(r"draw request\s*(\d+)", _norm(name))
        if m:
            n = int(m.group(1))
            if best is None or n > best[1]:
                best = (wb[name], n)
    if best is None:
        raise ValueError("No 'Draw Request N' sheet found in workbook")
    return best


def build_code_map(ws) -> dict:
    """Map major code -> {'main': row, 'no_ret': row_or_None}.

    A '(No Retainage)' row is associated with the most recent code seen above it
    (retention releases route there so retainage is not deducted a second time).
    """
    code_map: dict[str, dict] = {}
    last_code = None
    for r in range(FIRST_COST_ROW, LAST_COST_ROW + 1):
        code_raw = ws.cell(row=r, column=2).value
        desc = _norm(ws.cell(row=r, column=3).value)
        code = str(code_raw).strip() if code_raw not in (None, "") else None
        is_no_ret = "no retainage" in desc
        if is_no_ret:
            anchor = code or last_code
            if anchor is not None:
                code_map.setdefault(anchor, {"main": None, "no_ret": None})
                code_map[anchor]["no_ret"] = r
            if code:
                last_code = code
        elif code is not None:
            code_map.setdefault(code, {"main": None, "no_ret": None})
            if code_map[code]["main"] is None:
                code_map[code]["main"] = r
            last_code = code
    return code_map


def summary_formulas(row: int, cols: dict) -> dict:
    """Canonical per-row summary formulas, keyed by role, with resolved letters."""
    F = col_letter(CURRENT_BUDGET_COL)
    G = col_letter(FIRST_DRAW_COL)
    last_draw = col_letter(last_draw_col(cols))
    prev = col_letter(cols["prev"])
    cp = col_letter(cols["cp"])
    ret = col_letter(cols["ret"])
    tcd = col_letter(cols["tcd"])
    return {
        "prev": f"=SUM({G}{row}:{last_draw}{row})",
        "cplr": f"={cp}{row}-({cp}{row}*{ret}{row})",
        "tcd": f"=IF({ret}{row}>0,{prev}{row}/(1-{ret}{row})+{cp}{row},{prev}{row}+{cp}{row})",
        "pct": f'=IF({F}{row}=0,"-",({tcd}{row}/{F}{row}))',
        "rtd": f"={tcd}{row}*{ret}{row}",
        "bal": f"={F}{row}-{tcd}{row}",
    }


def load_invoices(path: str) -> dict:
    import json
    with open(path) as fh:
        data = json.load(fh)
    if "invoices" not in data:
        raise ValueError("invoices.json must have an 'invoices' array")
    for inv in data["invoices"]:
        inv.setdefault("retainage_rate", 0.0)
        inv.setdefault("is_retention_release", False)
        inv.setdefault("lien_release", "N/A")
        if "net_amount" not in inv:
            inv["net_amount"] = round(inv["gross_amount"] * (1 - inv["retainage_rate"]), 2)
        if "exhibit_line" not in inv:
            label = inv.get("invoice_number", "")
            inv["exhibit_line"] = f"{inv.get('vendor','').strip()} - {label}".strip(" -")
        inv["major_code"] = str(inv["major_code"]).strip()
    return data
