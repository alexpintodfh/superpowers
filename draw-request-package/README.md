# draw-request-package

A Claude Code **skill** that turns signed, coded invoice PDFs into a complete
land-development draw request — Exhibit D, an updated tracker/budget workbook,
a cover sheet, and the assembled package PDF — the way Draw Request #17 looks.

This is a **standalone, domain-specific skill** for the Anabelle Island draw
workflow. It is not part of the Superpowers core skill set.

## What it does

Drop signed & coded invoice PDFs into an `Invoices/` folder. The agent reads the
scanned invoices (including the handwritten CHECK ROUTING stamp that carries the
Major Code), then the scripts:

- add a new `Draw Request N` sheet to the tracker, freezing the prior draw into a
  history column (the workbook keeps one sheet per draw);
- write each invoice's **gross** into the **Current Period** column of its coded
  budget row and set the **Retainage** rate, so the tracker withholds retention
  and the **net** flows through;
- rebuild **Exhibit D** (both the standalone sheet and the copy embedded in the
  draw sheet), one line per invoice at net;
- generate the **cover sheet** with the amount to draw from the land banker;
- assemble **cover + Exhibit D + all invoice PDFs** into one package PDF.

## Layout

```
draw-request-package/
  SKILL.md                     agent-facing workflow (start here)
  reference/code_bible.md      Major Code -> budget line + retainage rules
  templates/invoices.schema.json   structured extraction schema
  scripts/
    render_invoices.py         PDFs -> PNGs so the agent can read scans
    new_draw_sheet.py          add Draw Request N+1 (column-shift + freeze)
    apply_invoices.py          invoices.json -> Current Period + Exhibit D
    build_cover.py             Application-for-Payment cover .docx
    assemble_package.py        cover + Exhibit D + invoices -> package PDF
    run_draw.py                orchestrates the mechanical steps
    draw_lib.py                shared tracker helpers
```

## Requirements

```
pip install pymupdf openpyxl python-docx
```

No LibreOffice required — PDFs are generated with PyMuPDF.

## Working folder convention

```
Draw #N/
  Invoices/                      <- drop invoice PDFs here
  Draw Package/
    tracker.xlsx                 <- the living tracker (Exhibit D + a sheet per draw)
    Cover_Sheet_N.docx           <- generated
    Draw_Request_N_package.pdf   <- generated
```

## Quick start (after invoices are in `Invoices/`)

```
# 1. render for reading
python3 scripts/render_invoices.py "Draw #18/Invoices" /tmp/render

# 2. agent reads the images and writes invoices.json (see templates/invoices.schema.json)

# 3. run the mechanical steps
python3 scripts/run_draw.py "Draw #18/Draw Package/tracker.xlsx" \
    invoices.json "Draw #18/Invoices" --out-dir "Draw #18/Draw Package"
```

The run prints the **draw total** (net of retainage) — the amount to request from
the land banker — and warns about any invoice whose Major Code didn't match a
tracker row.

## Notes on retention

Retention (retainage) is withheld on normal progress invoices (the "Less 10%
Retainage" line). It is **released only at the end of horizontal development**, on
its own invoice, entered at full amount with 0% retainage and routed to the
matching `(No Retainage)` row so it is never withheld twice.

## Validation

The scripts were validated against the real Draw Request #17 package: applying its
8 invoices reproduces the Exhibit D total of **$798,883.59** and a faithful
25-page package (cover + Exhibit D + invoices).
