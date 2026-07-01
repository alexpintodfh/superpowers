#!/usr/bin/env python3
"""Update a community's PRIOR cover sheet .docx into the new draw's cover.

Cover formats differ per community/land banker, so the reliable approach is to
take the previous draw's cover as a template and substitute three things:
  * the draw/application number (#N),
  * the signature date (always the 10th of the current month), and
  * the draw amount (the Exhibit D net total).

Only paragraphs that actually contain one of these is rewritten, so the rest of
the community-specific wording and formatting is preserved.
"""
from __future__ import annotations

import argparse
import json
import re
from datetime import date

from docx import Document

MONTHS = ("January February March April May June July August September "
          "October November December").split()
DATE_RE = re.compile(
    r"(" + "|".join(MONTHS) + r")\s+(\d{1,2})(st|nd|rd|th)?,\s*(\d{4})")
SLASH_DATE_RE = re.compile(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b")
CURRENCY_RE = re.compile(r"\$\s?[\d,]+\.\d{2}")


def _ordinal(n: int) -> str:
    if 10 <= n % 100 <= 20:
        return f"{n}th"
    return f"{n}" + {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")


def signature_date(today=None) -> date:
    today = today or date.today()
    return today.replace(day=10)


def _para_text(p) -> str:
    return "".join(r.text for r in p.runs)


def _set_para_text(p, text):
    """Replace a paragraph's text, keeping the first run's formatting."""
    if not p.runs:
        p.add_run(text)
        return
    p.runs[0].text = text
    for r in p.runs[1:]:
        r.text = ""


def transform_text(text, old_number, new_number, sig, total):
    long_date = f"{sig.strftime('%B')} {sig.day}, {sig.year}"
    long_date_ord = f"{sig.strftime('%B')} {_ordinal(sig.day)}, {sig.year}"
    slash_date = f"{sig.month}/{sig.day}/{sig.year}"

    # draw / application number: "#13" -> "#14"
    text = re.sub(rf"#\s*{old_number}\b", f"#{new_number}", text)

    # Dates: only the signature/draw date changes. Leave fixed reference dates
    # (e.g. the Construction/Development Agreement date) untouched — those live in
    # a sentence mentioning the agreement, so skip date edits in such paragraphs.
    if "agreement" not in text.lower():
        def _date_sub(m):
            return long_date_ord if m.group(3) else long_date
        text = DATE_RE.sub(_date_sub, text)
        text = SLASH_DATE_RE.sub(slash_date, text)

    # currency: only meaningful in an "amount" sentence
    if "amount" in text.lower():
        text = CURRENCY_RE.sub(f"${total:,.2f}", text)
    return text


def update_cover(template, invoices_json, out_path, old_number=None, total=None, today=None):
    with open(invoices_json) as fh:
        cfg = json.load(fh)
    new_number = cfg["draw_number"]
    if old_number is None:
        old_number = new_number - 1
    if total is None:
        total = round(sum(i.get("net_amount", i["gross_amount"] * (1 - i.get("retainage_rate", 0)))
                          for i in cfg["invoices"]), 2)
    sig = signature_date(today)

    doc = Document(template)
    for p in doc.paragraphs:
        original = _para_text(p)
        updated = transform_text(original, old_number, new_number, sig, total)
        if updated != original:
            _set_para_text(p, updated)
    # tables, if any
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    original = _para_text(p)
                    updated = transform_text(original, old_number, new_number, sig, total)
                    if updated != original:
                        _set_para_text(p, updated)
    doc.save(out_path)
    print(f"Cover sheet -> {out_path}  (#{new_number}, {sig.strftime('%-m/%-d/%Y')}, ${total:,.2f})")


def main():
    ap = argparse.ArgumentParser(description="Update a prior cover .docx into the new draw's cover.")
    ap.add_argument("template", help="Prior draw's cover .docx")
    ap.add_argument("invoices_json")
    ap.add_argument("out", help="Output .docx path")
    ap.add_argument("--old-number", type=int, help="Draw number in the template (default: new-1)")
    ap.add_argument("--total", type=float, help="Override draw total")
    args = ap.parse_args()
    update_cover(args.template, args.invoices_json, args.out, args.old_number, args.total)


if __name__ == "__main__":
    main()
