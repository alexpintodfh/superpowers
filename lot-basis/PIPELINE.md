# Lot Basis Quarterly Processing — Pipeline Design

**DreamFinders Homes — Jacksonville Division · LandDev**
**Companion to:** `skill/SKILL.md`, `skill/community_constants.json`, `skill/audit_template.py`

---

## 1. Goal

Each quarter, process 30+ Jacksonville community lot-basis workbooks (`.xlsm`) so that
**every cost line's Totals (col G) ties exactly to the JCS Adjusted Total Budget**, with
carry-cost formulas, SSC schedules, CCCX, and LD Contingency all correct — while keeping
every VP/AMC judgment call in front of a human, never auto-decided.

Success per community = the checklist in `SKILL.md` passes, independently verified, with a
VP/AMC review packet produced.

---

## 2. Constraints that shape the design

| Reality | Consequence for the pipeline |
|---|---|
| `openpyxl` writes formulas but does **not** recalculate them; `data_only` reads only Excel's last-saved cache | The pipeline must **independently compute** every intended value and verify `total = JCS`. Excel recalc is a *final human gate*, not our source of truth. |
| Source files are financial records on a mapped B: drive (no SharePoint) + manual JCS export | Never edit the source. Operate on a **copy**; require an audit pass before anyone saves back to B:. Files reach this environment by manual upload. |
| VP makes all budget calls; AMC reviews per-lot change > $1,500 | The tool **proposes + flags**; it must never silently decide a gap, hold-to-prior, or rate discrepancy. |
| `.xlsm` carries VBA + cross-sheet formulas (some pre-existing `#NAME?`) | Always load/save with `keep_vba=True`; distinguish our edits from pre-existing errors. |
| Community-specific constants (millage, HOA, CDD, closing rate, lot types, CCCX placement) | Config-driven, one record per community. **Template constant beats reference rate** — encoded and asserted. |

---

## 3. The processing pipeline (per community, per quarter)

Each stage is a **pure proposal** step (produces a list of intended cell edits + expected
values) until Stage G, which is the only step that writes a file. Everything is **dry-run by
default**; `--apply` writes to a copy.

```
  A. Intake & validate      →  B. SSC reconcile   →  C. Carry-cost formulas
        │                                                    │
        └──────────────► D. 99 dev-cost reconcile ◄──────────┘
                                   │
                         E. Independent audit (recompute, verify vs JCS)
                                   │
                         F. VP / AMC review packet (diffs + flags)
                                   │
                         G. Apply to copy → human recalc in Excel → archive
```

### A. Intake & validate inputs
- **Inputs:** Q3 working template, Q2 reference template, JCS export (adjusted total per cost code, 99 + 11).
- Confirm Detail sheet + expected row map; detect community code (`11XXXXXX`), lot-type count (1 vs 2 closings rows), and column map (H=actuals, I=quarter start, L=one-out, BP=last).
- Load the community config record. Fail loudly on any structural mismatch — never guess layout.

### B. SSC reconciliation (deterministic)
- For rows 10/27/44(/45): compute `G = SUM(H:BP)` and compare to `A`.
- `G<A`: extend the tail on the existing monthly pattern. `G>A`: trim the tail. **Never touch the core absorption pattern.**
- Output: proposed tail edits so `G=A`.

### C. Carry-cost formula build (11 lines)
- Read the **Q2 formulas exactly** and re-template to the Q3 column layout (do not approximate).
- Emit the community's formula for **every column I→BP**; compute the **BP trueup** to anchor the total to JCS (skip trueup when the small gap is cleanly absorbed by CCCX).
- **CCCX:** `G44=A44` → `MAX(SUM(F66:F86)-SUM(G66:G85),0)` in col L, clear col I; else stack col I.
- Constants come from config with the **template-precedence** rule enforced.

### D. 99 dev-cost reconciliation
- For each line with remaining budget: `col_I = JCS − Actuals − pre_set_future`.
- Preserve pre-set future cashflows (J, K, R, …) and straightline patterns (same # months, recompute per-month, last month absorbs rounding).
- **LD Contingency:** `MAX(SUM(F93:F133)-SUM(G93:G128,G130:G133),0)` in col L, clear col I.
- **Hold-to-prior** (col F > JCS, e.g. Site Contract): propose hold + **flag as VP decision** — do not decide it.
- Never write a negative col I. Verify row 136 grand total = JCS 99 total.

### E. Independent audit (the correctness gate)
Recompute intended values **without** relying on Excel, and assert:
- every line `total = JCS` (within rounding tolerance);
- SSC `G=A` on all three rows;
- row 136 grand total = JCS 99 total; carry total ties to JCS 11 total;
- all carry-cost cells I→BP are formulas (none silently hardcoded);
- no negative col I; pre-existing `#NAME?` cells are reported as pre-existing, not introduced.
- Output: per-line green/red audit report.

### F. VP / AMC review packet
- Diff vs Q2 (or prior submitted): **per-lot** cost change per code; flag every change **> $1,500** (AMC threshold).
- List open VP decisions: rate discrepancies (template vs master list), hold-to-prior candidates, gaps intentionally left to contingency.
- One-page-per-community summary a human can sign off on.

### G. Apply & finalize
- Write formulas into a **copy** of the Q3 template (`keep_vba=True`); never the source.
- Human opens in Excel → recalc → confirms the recalculated totals match our audit → saves to B:.
- Archive per community per quarter: processed template + audit report + VP packet.

---

## 4. Tooling architecture

```
lot-basis/
  skill/                     # the source-of-truth rules (uploaded)
  config/communities.json    # per-community constants & structure (grows from community_constants.json)
  lib/
    template_model.py        # parse Detail sheet: row map, column map, lot-type detection
    jcs.py                   # parse JCS export → {cost_code: adjusted_total}
    rules/
      ssc.py                 # G=A tail fix
      carry_costs.py         # 11-line formula engine (per-community templates)
      dev_costs.py           # 99-line col-I reconciliation
      contingency.py         # LD Contingency + CCCX placement
    writer.py                # apply edits to .xlsm copy (keep_vba)
    audit.py                 # independent recompute + verify (extends skill/audit_template.py)
    vp_packet.py             # diffs + AMC >$1,500 flags + open decisions
  run.py                     # orchestrate one community or batch; --dry-run (default) / --apply
  tests/
    golden/                  # Q2 inputs + known-good Q3 outputs for the 3 documented communities
```

Each rule module is a **pure function**: `(template_model, jcs, config) → [CellEdit] + [ExpectedValue] + [Flag]`.
The writer is the only impure step. The auditor consumes ExpectedValues so it never depends on Excel recalc.

---

## 5. Correctness guardrails (this is money)

- **Golden-master TDD.** The 3 documented communities (Anabelle Island, Bartram Commons, Ellis Cove) have exact formulas in `SKILL.md`. The engine must reproduce them byte-for-byte before we trust it on anything else.
- **Dry-run by default**; writing requires explicit `--apply` and targets a copy.
- **Template constant > reference rate** — asserted, not assumed.
- **Idempotence** — re-running yields identical output.
- **Human gates** — VP packet before finalize; Excel recalc before save-back.
- **Never** silently: leave a gap, hold-to-prior, resolve a rate discrepancy, or write a negative col I.

---

## 6. How we build it (development roadmap)

Each phase = TDD + golden-master against known-good + review + verification before moving on.

- **Phase 0 — Samples & environment.** Get real `.xlsm` for the 3 documented communities (Q2 + Q3) plus their JCS exports. *Nothing downstream can be validated without these.* Decide where processing runs (this container upload/download loop vs. local Windows).
- **Phase 1 — Model + parsers.** `template_model` + `jcs`; lock the row/column map across the 3 communities; reproduce `audit_template.py` reads on real files.
- **Phase 2 — SSC rule.** Simplest, fully deterministic; golden tests for extend/trim.
- **Phase 3 — Carry-cost engine.** Anabelle (2-type stack), Bartram (MAX), Ellis Cove (1% + MAX). Golden tests reproduce documented formulas exactly, incl. BP trueup.
- **Phase 4 — 99 reconciliation.** col-I, straightline, pre-set future, LD Contingency, hold-to-prior flags, row 136 tie-out.
- **Phase 5 — Audit + VP packet.** Independent recompute; AMC >$1,500 flag; open-decision list.
- **Phase 6 — Writer + end-to-end.** Apply to a copy for the 3 communities; human verifies in Excel.
- **Phase 7 — Scale to 30+.** Expand config per community; batch runner; per-community validation report.

---

## 7. Open decisions (need input before Phase 0/1)

1. **Trust model:** tool writes formulas into a copy for human review in Excel (recommended), vs. tool only emits a change-list a human applies by hand.
2. **Where it runs:** portable Python run locally on Windows (natural for B: drive + Excel recalc) vs. upload files into this environment per community.
3. **Starting scope:** build + validate on the 3 documented communities first (recommended), then expand — vs. attempt all 30+ now.
4. **JCS format:** what a JCS export actually looks like (columns, file type) so `jcs.py` can parse it reliably.
