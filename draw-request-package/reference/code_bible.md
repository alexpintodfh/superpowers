# Major Code "Code Bible" — Anabelle Island

The handwritten **CHECK ROUTING** stamp on each approved invoice carries the
`Major Code`. That code maps to a budget line (below) and drives which row of the
tracker's cost-code column the invoice's amount lands in.

Keep this file in sync with the client's official Code Bible. The tracker sheet
itself is the source of truth for which codes have rows; codes here that are not
present as rows in the tracker will be reported as warnings by `apply_invoices.py`.

## 99 Codes (Land Development)

| Major Code | Description |
|---|---|
| 1000 | Landbank Development Costs |
| 1005 | Due Diligence |
| 1006 | Commissions |
| 1007 | Investor DD & Underwriting |
| 1008 | Closing Costs |
| 1009 | Land Deposits |
| 1010 | Land Acquisition Costs |
| 1011 | Deal Extension Fees |
| 1012 | Bonds |
| 1013 | Zoning & Entitlements |
| 1014 | Legal |
| 1015 | Land Planning & Design |
| 1016 | Architecture |
| 1017 | Platting |
| 1018 | Surveying |
| 1019 | Geotechnical |
| 1020 | Engineering |
| 1021 | Environmental Consulting |
| 1022 | School Concurrency |
| 1023 | Permitting |
| 1024 | Traffic Concurrency |
| 1025 | Site Contract |
| 1027 | Off Sites |
| 1028 | Utility Connection Fees |
| 1029 | Water Rights (Credit) CO |
| 1030 | Environmental Impacts |
| 1032 | Lift Station |
| 1033 | Electrical |
| 1034 | Street Lights |
| 1035 | Wetland Impacts |
| 1036 | Tree Mitigation |
| 1037 | Gas |
| 1040 | Signage |
| 1042 | Fencing, Walls, & Docks |
| 1045 | Landscaping |
| 1050 | Amenities |
| 1053 | Mailboxes |
| 1055 | LD Contingency |
| 1056 | 2nd Lift Asphalt |
| 1075 | Development Reimbursements |
| 1085 | Development Management CEI |
| 1086 | Development Consulting |

## 11 Codes

| Major | Description |
|---|---|
| 1000 | Landbank Carrying Costs |
| 1008 | Closing Costs |
| 1060 | RE Taxes, Fees, etc. |
| 1062 | CDD Fees |
| 1063 | CDD Maintenance |
| 1065 | Home Owners Association |
| 1097 | Lot Options Fees |
| 1100 | Carrying Cost Alloc to Lots |
| 1107 | LOF Allocation to Lots |
| 1150 | Due Diligence |
| 1151 | Arch/Mech/Struct Eng-MstrPln |
| 1152 | Dirt Import/Export |
| 1153 | Over Excavation |
| 1154 | Underdrains |
| 1155 | Erosion Control |
| 1156 | Common Area Landscaping |
| 1157 | Final Acceptance Costs |
| 1158 | Overfunding |
| 1159 | Project Specific |
| 1160 | Common Area Maintenance |
| 1161 | Weather Conditions |
| 1162 | Transfer Price |

## GL routing strings (from the Code Bible footer)

- **11's** — `J,1000,11,(JOBCODE),999,(MAJORCODE),20` — Go into Land Dev
- **99's** — `J,1000,99,(JOBCODE),999,(MAJORCODE)` — Go into Land Dev
- **Before AMC approved** — `G,1000,035,1500,LDDD-???` — Go to Landpayables email
- **35's** — `1300,WIP,HJAX,035,(PROJECT CODE),0000,2038,15` — invoices DFH & Taheka Brown to upload to Concur
- **Dead Deal's** — `G,1000,035,6139` — Go to Landpayables email

## Retainage / retention rule

- Normal progress invoices that show a **"Less 10% Retainage"** line are entered
  at **gross** in the Current Period column with the retainage rate set, so the
  tracker withholds retention and the **net** flows to Exhibit D.
- **Retention is only released at the end of horizontal development**, on its own
  invoice (e.g. "Pipeline Constructors - RET"). A retention-release invoice is
  entered at its full amount with **0% retainage**, routed to the matching
  **"(No Retainage)"** row for that cost code, so retention is never withheld a
  second time.
