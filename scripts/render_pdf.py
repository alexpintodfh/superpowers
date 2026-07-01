#!/usr/bin/env python3
"""Render PDF pages to PNG images using PyMuPDF."""

import sys
import argparse
from pathlib import Path

import fitz  # PyMuPDF


def render_pdf(pdf_path: str, output_dir: str, dpi: int = 150) -> list[str]:
    doc = fitz.open(pdf_path)
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    outputs = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=mat)
        dest = out / f"page_{i + 1:04d}.png"
        pix.save(str(dest))
        outputs.append(str(dest))
        print(f"Rendered page {i + 1}/{len(doc)} -> {dest}")
    doc.close()
    return outputs


def main():
    parser = argparse.ArgumentParser(description="Render PDF pages to PNG images.")
    parser.add_argument("pdf", help="Path to the PDF file")
    parser.add_argument("output_dir", help="Directory to write PNG files into")
    parser.add_argument("--dpi", type=int, default=150, help="Resolution (default: 150)")
    args = parser.parse_args()

    paths = render_pdf(args.pdf, args.output_dir, args.dpi)
    print(f"\nDone: {len(paths)} page(s) written to {args.output_dir}")


if __name__ == "__main__":
    main()
