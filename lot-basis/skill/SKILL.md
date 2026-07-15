# Lot Basis Quarterly Processing Skill
**DreamFinders Homes — Jacksonville Division**
**Maintained by: Alex Pinto, LandDev**
**Last updated: 26Q3 session**

---

## What This Skill Covers

Quarterly processing of lot basis templates (`.xlsm`) for 30+ Jacksonville communities. Each quarter, templates are updated so that every cost line's **Totals column matches the JCS Adjusted Total Budget** exactly. Corporate (AMC) reviews per-lot cost changes above $1,500. The VP of Land makes all final budget decisions.

---

## File Locations

| File Type | Location |
|-----------|----------|
| Q2 templates (reference) | `C:\Users\alex.pinto\OneDrive - DreamFinders Homes\Desktop\Lot Basis History\26Q2 Lot Basis Templates With Formulas` |
| Q3 templates (working) | `B:\Land\Lot Basis Tracker\26Q3 LB Workbooks\Jacksonville` |
| JCS exports | Manual upload from LandDev web app |
| HOA/CDD/Millage tracker | `HOA_CDD_Master_List_v2.xlsm` (outputs folder) |

The B: drive is a mapped network drive — not accessible via SharePoint. Q3 templates must be manually uploaded. JCS files must be manually exported from LandDev.

---

## Template Structure (Detail Sheet)

| Row(s) | Content |
|--------|---------|
| 8 | Date headers (monthly columns) |
| 10 | Sales schedule |
| 27 | Starts schedule |
| 44 | Closings schedule (single lot type) |
| 44–45 | Closings schedule (two lot types, e.g. 40ft + 55ft) |
| 66–86 | **11 Carry Costs** |
| 93–133 | **99 Dev Costs** |
| 135 | Land Purchase Price |
| 136 | 99 Grand Total |
| 144 | Lot Takedowns schedule |
| 144–145 | Lot Takedowns (two lot types) |
| 165 | Lot price (D165 = single type; D165+D166 = two types) |

**Column key (Q3 = Jul-26 start):**
- Col H = Actuals (06/30/26 for Q3)
- Col I = Jul-26 (first future cashflow = current quarter start)
- Col L = Oct-26 (one quarter out — used for LD Contingency and CCCX MAX formulas)
- Col BP = last date column in most templates

---

## Core Processing Rules

### Rule 1 — Totals Must Match JCS
Every line's **col G (Totals)** must equal the **JCS Adjusted Total Budget** for that cost code. This is non-negotiable. Verify every line. `Total = Actuals + SUM(col_I : last_col)`.

### Rule 2 — Pre-Set Future Cashflows Are Untouched
If a PM has already placed cashflows in future columns (J, K, R, etc.), those stay. Only adjust col I so that `Actuals + col_I + pre-set_future = JCS`.

### Rule 3 — Col I = JCS − Actuals − Pre-Set Future (for every line with remaining budget)
Set col I for **every** line that still has remaining budget, regardless of gap size. Do not decide on your own to leave gaps to LD Contingency — that is the VP's judgment call, not yours. The only reasons NOT to set col I:
  1. VP explicitly decided to leave a gap (e.g., Engineering in Ellis Cove — remaining not expected to be spent)
  2. Col I would be negative (impossible — never write a negative cashflow in col I)

### Rule 4 — LD Contingency Uses MAX Formula
Always place in col L (one quarter out from current quarter start):
```
=MAX(SUM(F93:F133)-SUM(G93:G128,G130:G133),0)
```
Clear col I for this row. The MAX formula absorbs whatever the VP leaves behind from other lines plus legitimate contingency budget.

**Exception:** Early-stage communities with incomplete closings schedules (G44 ≠ A44) use a hardcoded stack in col I instead of the MAX formula. Example: Anabelle Island Q3 (319/513 lots scheduled).

### Rule 5 — CCCX Uses MAX Formula When Schedule Is Complete
When G44 = A44 (all lots scheduled), place in col L:
```
=MAX(SUM(F66:F86)-SUM(G66:G85),0)
```
Clear col I. When schedule is incomplete, stack col I.

### Rule 6 — Straightline Cashflows: Preserve Pattern, Update Amount
When a PM has set identical monthly amounts across consecutive months, that's a straightline instruction. Keep the same number of months. Update only the per-month amount so the total hits JCS:
```
per_month = (JCS - Actuals) / number_of_months
```
Last month absorbs rounding. Example: Ellis Cove Amenities — 6 months at 136,011.33, last month 136,010.33.

### Rule 7 — Hold to Prior (Site Contract / large lines)
When the Previous (col F) budget exceeds the JCS Adjusted Total, set col I to hold at the Previous level. LD Contingency absorbs the gap to JCS. This is a VP decision — look for it on large cost codes like Site Contract where actuals already exceed JCS.

### Rule 8 — Small Gaps and LD Contingency
LD Contingency absorbs gaps left by VP judgment calls. The grand total (row 136) must still tie to JCS 99 total. Verify with: `row_136_total = JCS_99_total`.

---

## SSC (Sales, Starts, Closings) Schedule Rules

The SSC schedule lives in rows 10 (Sales), 27 (Starts), 44 (Closings).

**The one rule:** Col G must equal col A (lot count) for every row.
```
G = SUM(H:BP) = A
```

**How to fix when G ≠ A:**
- **G < A (short):** Extend the schedule at the tail. Add months continuing the existing monthly pattern (usually 4/month). Example: Sales G=145, A=150 → extend AP from 2→4, add AQ=3.
- **G > A (over):** Trim from the tail. Reduce or clear the last month(s). Example: Starts G=154, A=150 → AY: 4→2, AZ: clear.

**Never touch the core absorption pattern** — only adjust at the tail end. The orange-highlighted cell in the template marks the current period.

---

## 11 Carry Cost Formula Patterns

### Universal Pattern
- Formula runs in **every** column from col I through last col (BP)
- The formula evaluates to zero where there are no closings/takedowns — this is expected, not hardcoded
- Last column (BP) gets `base_formula - TRUEUP_CONSTANT` to anchor total to JCS
- If formula gap is small and absorbed by CCCX, no trueup needed in BP (clean formula end-to-end)

### Formula Templates by Community

#### Anabelle Island (11ANAISL) — Two lot types (40ft row 44, 55ft row 45)
```
Closing Costs:  =SUMPRODUCT($D$165:$D$166)*SUMPRODUCT({col}$144:{col}$145))*0.75%
RE Taxes:       =(SUM($G$44:$G$45)-SUM($H$44:{prev_col}$45))*($D$165*15.2187*0.75/1000/12)
CDD Fees:       =(SUM($A$44)-SUM($H44:{col}44))*(1221.14/12)+(SUM($A$45)-SUM($H45:{col}45))*(1690.83/12)
HOA:            =(SUMPRODUCT($G$44:$G$45)-SUMPRODUCT($H$44:{col}$45))*(386.11/12)
```
- Millage: 15.2187 (Clay County, Lake Asbury MSBD) — confirmed from template
- Lot price: $117,640.08
- CCCX: col I stack (schedule incomplete: 319/513 lots)

#### Bartram Commons (11BARTCM) — Single lot type (TH)
```
Closing Costs:  ={col}144*$D$165*0.75%
RE Taxes:       =(($D$165*0.0177412/12)*($A$44-SUM($H$44:{prev_col}$44)))*0.75
HOA:            =((125)*($A$44-SUM($H$44:{prev_col}$44)))
```
- Millage: 0.0177412 (Duval County USD1) — updated from template's 0.014
- Lot price: $110,330.99
- HOA rate: $125/month (template constant; master list says $173 — VP to confirm)
- CCCX: MAX formula col L (158/158 lots complete)

#### Ellis Cove (11ELLCOV) — Single lot type (TH)
```
Closing Costs:  =(0.01*$D$165)*{col}144          ← NOTE: 1%, not 0.75%
RE Taxes:       =($A$44-(SUM($H$44:{col}44)))*((((17.865/1000)*$D$165)*0.75)/12)
HOA:            =($A$44-SUM($H$44:{col}44))*(80)
```
- Millage: 17.865 per $1,000 (template constant from Q2; USD1 actual = 17.7412)
- Lot price: $119,876.28
- HOA rate: $80/month (template constant; master list says $175 — VP to confirm)
- Closing cost rate: **1%** (community-specific — different from all others)
- CCCX: MAX formula col L (150/150 lots complete)

---

## Millage Rate Reference

| County | District | Rate (per $1) | Per $1,000 | DFH Communities |
|--------|----------|---------------|------------|-----------------|
| Clay | Unincorporated (Lake Asbury MSBD) | 0.0152523 | 15.2523 | Anabelle Island (template: 15.2187), Amberly, Jennings Farm |
| Clay | Challenger Center MSTU | 0.0182523 | 18.2523 | Challenger Center |
| Duval | USD1 | 0.0177412 | 17.7412 | Bartram Commons, Ellis Cove, Owens Rd, Pinewalk, Bellbrooke, Dunns 4, Everlake, Everrange, McLaurin Rd |
| St. Johns | District 300 | 0.0125415 | 12.5415 | Beacon Lake, Rev at Trailmark, Shearwater, Silverleaf |
| St. Johns | District 450 | 0.0126013 | 12.6013 | Cordova Palms |
| Flagler | County | 0.0183217 | 18.3217 | Reveire at Palm Coast, Spring Lake |

Source: `Tax_lookup_millage_rates.xlsx` (coworker reference file). Template constants take precedence over reference rates.

---

## HOA / CDD Rate Reference

| Community | HOA ($/mo) | CDD | Notes |
|-----------|-----------|-----|-------|
| Anabelle Island | $386.11/yr ÷ 12 = $32.18/mo | 40ft=$1,221.14/yr, 55ft=$1,690.83/yr | Template confirmed |
| Bartram Commons | $125/mo | None | Template says $125; master list says $173 — VP decision pending |
| Ellis Cove | $80/mo | None | Template says $80; master list says $175 — VP decision pending |

---

## Open VP Decisions (as of 26Q3)

1. **Bartram Commons HOA:** Template $125/mo vs master list $173/mo
2. **Bartram Commons RE Tax millage:** Template 0.014 vs USD1 actual 0.0177412
3. **Ellis Cove HOA:** Template $80/mo vs master list $175/mo
4. **Ellis Cove Closing Cost rate:** 1% (confirmed from template — community-specific)

---

## Processing Checklist (per community)

```
[ ] Upload Q3 template, Q2 template, JCS export
[ ] Read Q2 carry cost formulas exactly (do not approximate)
[ ] Check SSC: Sales G=A, Starts G=A, Closings G=A — fix tail if needed
[ ] Determine CCCX placement: G44=A44 → MAX formula col L; else stack col I
[ ] Write carry cost formulas I through BP (every cell, formula-driven)
[ ] Compute trueup for BP: (actuals + formula_sum) - JCS → subtract from BP formula
[ ] If gap < ~$5k and formula is clean, let CCCX absorb (no trueup)
[ ] Audit all 99 lines vs JCS:
    [ ] Set col I = JCS - actuals - pre_set_future for every line with remaining budget
    [ ] Preserve straightline patterns (same # months, updated per-month amount)
    [ ] Leave pre-set future cashflows in place (J, K, R, etc.)
    [ ] LD Contingency: MAX formula col L, clear col I
    [ ] Verify row 136 grand total = JCS 99 total
[ ] Recalculate and verify all totals
[ ] Flag VP decisions (rate discrepancies, hold-to-prior lines)
```

---

## Common Mistakes to Avoid

| Mistake | Correct Approach |
|---------|-----------------|
| Leaving gaps to LD Contingency without VP instruction | Fill col I for every line with remaining budget |
| Treating quarterly cashflow patterns as hardcoded | They're formulas evaluating to zero — run formula in every col |
| Using reference millage rates without checking template | Template constant always takes precedence |
| Setting col I without checking pre-set future cashflows first | Compute: col_I = JCS − actuals − pre_set_future |
| Declaring a line "keep as-is" without verifying total = JCS | Always verify: total = JCS, no exceptions |
| Consolidating straightline cashflows | Preserve number of months, update per-month amount only |
| Not checking G=A for SSC rows | Always fix tail to make G=A before processing carry costs |
| Negative col I | Never — if result is negative, it's a VP/prior-period issue |
