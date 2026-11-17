# Support ticket SLA showcase

Open [Support-SLA.exlt](Support-SLA.exlt) in Excel-Lite. The Dashboard opens first with four cached KPIs, a monthly volume and breach-rate chart, and a team-by-month breach heatmap.

Maximize the window for the walkthrough. In a shorter window, use **View → Zoom** at about 85% to keep the cards, charts, and scenario instructions in view together.

The register contains **610 synthetic tickets** from January through June 2026, across five teams and four priorities. April has a Workspace launch surge, May has a webhook delivery incident in Integrations, and June ends with open tickets at different points in their response windows. The fixed baseline snapshot is **July 1, 2026 at 00:00 UTC**.

## Walkthrough

1. On **Dashboard**, confirm **610 opened**, **24.8% breached**, **8 open** and **77.0 h P90 resolution**. Volume peaks in April and the heatmap shows the May Integrations hotspot.
2. Open **Tickets** and filter the register by **Integrations** or **Webhook delivery incident**. These sheet filters help inspect records. The Python analysis reads the full named table.
3. On a copy, enable **File → Options → Formulas → Python Automatic**. Scroll right on **Dashboard** to the first-response targets and change the mint **P9** input from **4.0** to **2.0** hours. The breach rate rises to **32.1%** and both charts refresh. Open count and resolution P90 stay at their baseline values. In Manual mode, use **Run All** to calculate the scenario.
4. Open **Analysis** to inspect the Python spill at **B6**, the breach-rate pivot at **B26:H31**, the P90 resolution pivot at **B35:H40**, and the priority breakdown below them. **B52** and **B53** hold the two Matplotlib chart scripts. Restore **Dashboard!P9** to **4.0** to return to the baseline.

The response SLA compares each ticket’s elapsed first-reply time with its priority target. A ticket with no reply after its deadline is a breach. A ticket with no reply and time remaining is pending, so it is excluded from the breach-rate denominator. P90 uses tickets resolved by the snapshot. The `pandas.pivot_table()` results spill into worksheet ranges. Native formulas link the Dashboard KPIs to the shared analysis output.

## Rebuild and verify

The test runner builds the native workbook from deterministic synthetic data and the [analysis](analysis.py), [trend](trend.py) and [heatmap](heatmap.py) scripts. It uses the private Python runtime and the app’s mixed calculation and persistence paths. Run from the repository root after preparing Python.

```powershell
./scripts/test.ps1 -Configuration Debug
dotnet run --project tests/ExcelLite.Tests -c Debug -- --sla-showcase-only
dotnet run --project tests/ExcelLite.Tests -c Debug -- --sla-showcase-only --output build/Support-SLA.exlt
```

The focused checks reconcile the source records with all four KPIs, six months, both 5 × 6 pivots, and priority counts at baseline and after a target edit. They also verify invalid targets, cached reopen, and the delivered workbook’s embedded scripts. The output command creates a separate rebuild under the ignored build directory.
