#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Rasterise a PDF into per-page PNGs so a human (or the model) can eyeball it.

Used after ``docx2pdf.ps1`` to verify pagination, caption placement, table width
and stray indentation - things XML assertions cannot see.

    python pdf2png.py out.pdf pages/ [--scale 1.6] [--pages 1,6,18]

Requires ``pypdfium2`` (``pip install pypdfium2``); falls back to ``pdftoppm``
if pypdfium2 is missing and poppler is on PATH.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys


def render(pdf: str, outdir: str, scale: float, only: set | None) -> int:
    os.makedirs(outdir, exist_ok=True)
    try:
        import pypdfium2 as pdfium  # type: ignore
    except ImportError:
        return render_poppler(pdf, outdir, scale, only)

    doc = pdfium.PdfDocument(pdf)
    n = len(doc)
    for i in range(n):
        page_no = i + 1
        if only and page_no not in only:
            continue
        img = doc[i].render(scale=scale).to_pil()
        path = os.path.join(outdir, "p%02d.png" % page_no)
        img.save(path)
        print(path)
    print("# %d pages rendered" % n)
    return 0


def render_poppler(pdf: str, outdir: str, scale: float, only: set | None) -> int:
    if not shutil.which("pdftoppm"):
        print("need pypdfium2 or poppler's pdftoppm on PATH", file=sys.stderr)
        return 2
    dpi = str(int(72 * scale))
    prefix = os.path.join(outdir, "p")
    subprocess.run(["pdftoppm", "-png", "-r", dpi, pdf, prefix], check=True)
    print("rendered with pdftoppm into %s" % outdir)
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="PDF -> per-page PNG")
    ap.add_argument("pdf")
    ap.add_argument("outdir")
    ap.add_argument("--scale", type=float, default=1.6)
    ap.add_argument("--pages", help="comma-separated 1-based page numbers")
    a = ap.parse_args(argv)
    only = None
    if a.pages:
        only = {int(x) for x in a.pages.split(",") if x.strip()}
    return render(a.pdf, a.outdir, a.scale, only)


if __name__ == "__main__":
    raise SystemExit(main())
