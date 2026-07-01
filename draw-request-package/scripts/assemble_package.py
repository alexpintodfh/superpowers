#!/usr/bin/env python3
"""Assemble the draw package PDF: cover + Exhibit D + all invoice PDFs.

Uses only PyMuPDF (no LibreOffice), so it runs anywhere the rest of the skill
does. The cover and Exhibit D pages are drawn directly; the invoice PDFs are
appended in Exhibit D order. This is the land-banker-facing summary package
(the physically-signed original is produced after wet signatures).
"""
from __future__ import annotations

import argparse
import json
from datetime import date
from pathlib import Path

import fitz  # PyMuPDF

PAGE = fitz.paper_rect("letter")
MARGIN = 54
FONT = "helv"
FONT_B = "hebo"


def _ordinal(n: int) -> str:
    if 10 <= n % 100 <= 20:
        return f"{n}th"
    return f"{n}{{1:'st',2:'nd',3:'rd'}}".format() if False else f"{n}" + {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")


def draw_cover(doc, cfg, total):
    page = doc.new_page(width=PAGE.width, height=PAGE.height)
    cx = PAGE.width / 2
    y = 150
    for line, size in ((("APPLICATION FOR PAYMENT"), 13),
                       (cfg.get("project_name", "PROJECT").upper(), 13),
                       ("CERTIFICATION FOR PAYMENT", 13),
                       (f"Application {cfg['draw_number']}", 12)):
        w = fitz.get_text_length(line, fontname=FONT_B, fontsize=size)
        page.insert_text((cx - w / 2, y), line, fontname=FONT_B, fontsize=size)
        y += 28
    y += 10
    manager = cfg.get("project_manager", "Dream Finders Homes, LLC")
    owner = cfg.get("owner", "the Owner")
    agreement_date = cfg.get("agreement_date", "")
    body1 = (f'{manager} ("Project Manager"), the Project Manager under that certain '
             f'Development and Management Agreement dated {agreement_date}, between '
             f'{owner} ("Owner"), and Project Manager, hereby requests payment pursuant '
             f"to the attached Application for Payment.")
    exec_dt = cfg.get("application_date")
    d = date.fromisoformat(exec_dt) if exec_dt else date.today()
    body2 = (f"a) Application for Payment. The Project Manager hereby requests Owner to "
             f"make a payment in the amount of ${total:,.2f} (the \"Payment\") for the "
             f"Work set forth in the attached Application for Payment referred to as "
             f'"Application No. {cfg["draw_number"]}".')
    body3 = (f"b) Certification. The undersigned certifies that the foregoing information "
             f"is true and correct of his own knowledge. Executed as of "
             f"{_ordinal(d.day)} day of {d.strftime('%B')}, {d.year}.")
    box = fitz.Rect(MARGIN, y, PAGE.width - MARGIN, y + 320)
    text = body1 + "\n\n" + body2 + "\n\n" + body3
    page.insert_textbox(box, text, fontname=FONT, fontsize=11, align=0)
    sy = y + 340
    for line in ("By: ____________________________",
                 f"Name: {cfg.get('signer_name','')}",
                 f"Title: {cfg.get('signer_title','')}",
                 f"Date: {d.strftime('%-m/%-d/%Y')}"):
        page.insert_text((cx, sy), line, fontname=FONT, fontsize=11)
        sy += 26


def draw_exhibit_d(doc, cfg, invoices, total):
    page = doc.new_page(width=PAGE.width, height=PAGE.height)
    cx = PAGE.width / 2
    title = 'EXHIBIT "D"'
    w = fitz.get_text_length(title, fontname=FONT_B, fontsize=12)
    page.insert_text((cx - w / 2, 70), title, fontname=FONT_B, fontsize=12)
    page.insert_text((PAGE.width - MARGIN - 170, 95),
                     f"Draw Request #   {cfg['draw_number']}", fontname=FONT_B, fontsize=10)

    # column x-edges
    x0 = MARGIN
    cols = [x0, x0 + 62, x0 + 130, x0 + 344, x0 + 438, PAGE.width - MARGIN]
    headers = ["Invoices", "Major Code", "INVOICES SUBMITTED THIS REQUEST", "Amount", "Lien Release"]
    top = 115
    rowh = 22
    n = len(invoices)
    bottom = top + rowh * (n + 2)

    def cell_text(r_idx, c_idx, text, bold=False, align="c"):
        cy = top + rowh * r_idx + rowh - 6
        cxl, cxr = cols[c_idx], cols[c_idx + 1]
        fn = FONT_B if bold else FONT
        tw = fitz.get_text_length(str(text), fontname=fn, fontsize=9)
        if align == "c":
            tx = (cxl + cxr) / 2 - tw / 2
        elif align == "r":
            tx = cxr - 6 - tw
        else:
            tx = cxl + 6
        page.insert_text((tx, cy), str(text), fontname=fn, fontsize=9)

    # header row
    for c in range(len(headers)):
        cell_text(0, c, headers[c], bold=True, align="l" if c in (2,) else "c")
    for i, inv in enumerate(invoices):
        r = i + 1
        cell_text(r, 0, i + 1)
        cell_text(r, 1, inv["major_code"])
        cell_text(r, 2, inv["exhibit_line"])
        cell_text(r, 3, f"$ {inv['net_amount']:,.2f}", align="r")
        cell_text(r, 4, inv.get("lien_release", "N/A"))
    # total row
    tr = n + 1
    cell_text(tr, 0, "TOTAL", bold=True)
    cell_text(tr, 3, f"$ {total:,.2f}", bold=True, align="r")

    # grid
    for i in range(n + 3):
        yy = top + rowh * i
        page.draw_line((x0, yy), (cols[-1], yy))
    for cxx in cols:
        page.draw_line((cxx, top), (cxx, bottom))


def draw_cover_from_docx(doc, docx_path):
    """Render the actual cover .docx text into the package's first page(s).

    Keeps the package cover faithful to whatever community-specific cover was
    generated, without needing LibreOffice to convert docx -> PDF.
    """
    from docx import Document

    def _ascii(s):
        # the base-14 PDF font lacks curly quotes / dashes; map to ASCII so they
        # don't render as stray dots
        return (s.replace("‘", "'").replace("’", "'")
                 .replace("“", '"').replace("”", '"')
                 .replace("–", "-").replace("—", "-")
                 .replace("…", "...").replace("\xa0", " "))

    paras = [_ascii("".join(r.text for r in p.runs)).rstrip()
             for p in Document(docx_path).paragraphs]

    fontsize, leading = 11, 15
    left, right = MARGIN, PAGE.width - MARGIN
    top, bottom = 80, PAGE.height - 60
    width = right - left

    def wrap(par):
        if not par.strip():
            return [""]
        lines, cur = [], ""
        for word in par.split():
            trial = (cur + " " + word).strip()
            if fitz.get_text_length(trial, fontname=FONT, fontsize=fontsize) <= width:
                cur = trial
            else:
                if cur:
                    lines.append(cur)
                cur = word
        if cur:
            lines.append(cur)
        return lines

    page = doc.new_page(width=PAGE.width, height=PAGE.height)
    y = top
    for par in paras:
        for line in wrap(par):
            if y > bottom:
                page = doc.new_page(width=PAGE.width, height=PAGE.height)
                y = top
            if line:
                page.insert_text((left, y), line, fontname=FONT, fontsize=fontsize)
            y += leading


def main():
    ap = argparse.ArgumentParser(description="Assemble cover + Exhibit D + invoice PDFs into one package PDF.")
    ap.add_argument("invoices_json")
    ap.add_argument("invoices_dir", help="Folder holding the invoice PDFs referenced by source_pdf")
    ap.add_argument("out_pdf")
    ap.add_argument("--total", type=float)
    ap.add_argument("--cover-docx", help="Render this cover .docx as the package cover page(s)")
    args = ap.parse_args()

    with open(args.invoices_json) as fh:
        cfg = json.load(fh)
    invoices = cfg["invoices"]
    for inv in invoices:
        inv.setdefault("retainage_rate", 0.0)
        inv.setdefault("net_amount", round(inv["gross_amount"] * (1 - inv["retainage_rate"]), 2))
        inv.setdefault("lien_release", "N/A")
        inv.setdefault("exhibit_line", f"{inv.get('vendor','')} - {inv.get('invoice_number','')}".strip(" -"))
    total = args.total if args.total is not None else round(sum(i["net_amount"] for i in invoices), 2)

    doc = fitz.open()
    if args.cover_docx:
        draw_cover_from_docx(doc, args.cover_docx)
    else:
        draw_cover(doc, cfg, total)
    draw_exhibit_d(doc, cfg, invoices, total)

    inv_dir = Path(args.invoices_dir)
    missing = []
    for inv in invoices:
        src = inv.get("source_pdf")
        pdf_path = inv_dir / src if src else None
        if pdf_path and pdf_path.exists():
            with fitz.open(pdf_path) as ipdf:
                doc.insert_pdf(ipdf)
        else:
            missing.append(src or inv["exhibit_line"])
    doc.save(args.out_pdf, garbage=4, deflate=True)
    doc.close()
    print(f"Package -> {args.out_pdf}  ({len(invoices)} invoices, total ${total:,.2f})")
    if missing:
        print("  NOTE: invoice PDFs not found (cover+Exhibit D still generated):")
        for m in missing:
            print(f"    - {m}")


if __name__ == "__main__":
    main()
