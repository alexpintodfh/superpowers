---
name: draw-request-package
description: "Use when building a land-development draw request / Application for Payment from scanned, coded invoices — turns invoice PDFs dropped into an Invoices folder into an Exhibit D, updates the draw tracker/budget workbook (new sheet per draw, Current Period + retainage), and assembles the cover + Exhibit D + invoices package PDF. Triggers: 'draw request', 'Exhibit D', 'draw package', 'land banker draw', 'Application for Payment'."
---

# Draw Request Package

Build a draw request the way Draw Request #17 looks: drop signed, coded invoice
PDFs into an `Invoices/` folder and produce the Exhibit D, the updated tracker,
the cover sheet, and the assembled package PDF.

You (the agent) do the one thing scripts can't: **read the scanned invoices**
(they are images with handwritten CHECK ROUTING stamps — OCR is not reliable).
Everything else is deterministic Python.

## Folder layout

```
Draw #N/
  Invoices/                 <- drop the signed & coded invoice PDFs here
  Draw Package/
    tracker.xlsx            <- the living draw tracker/budget (Exhibit D + a sheet per draw)
    Cover_Sheet_N.docx      <- generated
    Draw_Request_N_package.pdf   <- generated (cover + Exhibit D + invoices)
```

The tracker is the **living document**. It already contains one `Draw Request K`
sheet per past draw plus a standalone `Exhibit D` sheet. Each new draw adds a
new sheet and freezes the prior draw into a history column.

## Workflow

Run these from the skill's `scripts/` directory. Requires `pymupdf`, `openpyxl`,
`python-docx` (`pip install pymupdf openpyxl python-docx`).

1. **Render the dropped invoices to images so you can read them.**
   ```
   python3 scripts/render_invoices.py "Draw #N/Invoices" /tmp/draw_render
   ```
   Read every rendered page.

2. **Extract each invoice into `invoices.json`** (schema in
   `templates/invoices.schema.json`). For each invoice capture, from the invoice
   body and the handwritten **CHECK ROUTING** stamp:
   - `vendor`, `invoice_number` → these form the Exhibit D line `"Vendor - Invoice#"`.
   - `major_code` — the handwritten `Major Code:` on the stamp (e.g. `1018`).
     Cross-check its meaning in `reference/code_bible.md`.
   - `gross_amount` — the current/gross billing **before** retainage.
   - `retainage_rate` — `0.1` if the invoice shows a "Less 10% Retainage" line,
     else `0`. (`net_amount` is then gross×(1−rate); it is what shows on Exhibit D.)
   - `is_retention_release` — `true` only for an end-of-development retention
     **release** invoice (often labeled `RET`). These carry `retainage_rate` 0.
   - `lien_release` — **check every page of the invoice PDF.** If any page is a
     "Conditional Waiver and Release of Lien" (upon progress OR final payment),
     set `"Yes"`; otherwise `"N/A"`. The waiver is usually the LAST page of a
     multi-page contractor packet — do not decide from the first page alone.
   - `source_pdf` — the invoice's filename in `Invoices/`.
   Set `draw_number` and the cover fields (`application_date`, `owner`, `signer_name`, …).

3. **Run the mechanical steps.**
   ```
   python3 scripts/run_draw.py "Draw #N/Draw Package/tracker.xlsx" invoices.json "Draw #N/Invoices" --out-dir "Draw #N/Draw Package"
   ```
   This creates the `Draw Request N` sheet (freezing the prior draw), writes each
   invoice's gross into the **Current Period** column of its coded row and sets the
   retainage rate, rebuilds both Exhibit D views, generates the cover sheet, and
   assembles the package PDF. It prints the **draw total** (net) — the amount to
   request from the land banker.

   (You can also run the steps individually: `new_draw_sheet.py`,
   `apply_invoices.py`, `build_cover.py`, `assemble_package.py`.)

4. **Verify before handing off.** Confirm:
   - The Exhibit D total equals the sum you get adding the invoice net amounts by hand.
   - Every `major_code` matched a tracker row (the script warns on any that didn't).
   - The cover-sheet amount equals the Exhibit D total.
   Report the total and any warnings to your human partner. Do not claim the draw
   is complete if any invoice failed to map or a total does not reconcile.

## How amounts land in the tracker

- **Current Period** column = **gross** billing; the **Retainage** column holds
  the rate; **Current Period Less Retainage** = the net. Exhibit D shows nets.
- Several invoices with the same major code sum into that code's one row.
- **Retention releases** route to the code's **"(No Retainage)"** row at full
  amount with 0% retainage, so retention is never withheld twice. Only ask for
  retention money at the end of horizontal development, on its own invoice.
- The columns shift one to the right every draw, so scripts locate them by the
  header text in row 4 — never hard-code column letters.

## Gotchas

- The invoices are **scanned images**; `render_invoices.py` + your own reading is
  the only reliable extraction. Do not trust `page.get_text()` on them.
- A draw's Exhibit D header must show the **current** draw number. (Draw 17's
  files shipped mislabeled "16" — always set `draw_number` correctly.)
- Read the handwritten stamp carefully; the Major Code decides the budget line.
