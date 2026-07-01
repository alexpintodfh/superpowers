#!/usr/bin/env python3
"""Render every invoice PDF in a folder to PNG images for the agent to read.

The invoices are scanned images with handwritten CHECK ROUTING stamps, so text
extraction is unreliable — the agent reads the rendered pages visually. Each PDF
gets one PNG per page under <out>/<pdf-stem>/page_NNNN.png.
"""
from __future__ import annotations

import argparse
from pathlib import Path

import fitz  # PyMuPDF


def render_folder(invoices_dir: str, out_dir: str, dpi: int = 150) -> list[str]:
    src = Path(invoices_dir)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    written = []
    pdfs = sorted(p for p in src.glob("*.pdf"))
    if not pdfs:
        print(f"No PDFs found in {src}")
        return written
    for pdf in pdfs:
        dest = out / pdf.stem
        dest.mkdir(parents=True, exist_ok=True)
        doc = fitz.open(pdf)
        for i, page in enumerate(doc):
            f = dest / f"page_{i + 1:04d}.png"
            page.get_pixmap(matrix=mat).save(str(f))
            written.append(str(f))
        print(f"{pdf.name}: {len(doc)} page(s) -> {dest}/")
        doc.close()
    print(f"\nRendered {len(pdfs)} invoice PDF(s), {len(written)} page(s) total, into {out}/")
    return written


def main():
    ap = argparse.ArgumentParser(description="Render all invoice PDFs in a folder to PNGs.")
    ap.add_argument("invoices_dir", help="Folder containing the invoice PDFs")
    ap.add_argument("out_dir", help="Where to write the rendered PNGs")
    ap.add_argument("--dpi", type=int, default=150)
    args = ap.parse_args()
    render_folder(args.invoices_dir, args.out_dir, args.dpi)


if __name__ == "__main__":
    main()
