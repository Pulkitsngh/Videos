# Navjyoti Vermicompost Tracker

Cloud register for Navjyoti's vermicompost production, one batch after another.
It replaces the per-batch Excel register (Batch 4 layout) with a shared online app.

Live app: https://claude.ai/artifact/SL9xyoMd91JG3XVMCb2LWr

## What it covers

| Tab | Excel sheet it replaces | What you record |
|---|---|---|
| Overview | Dashboard & Traceability | KPIs, material flow, produced vs sold chart, alerts, cost breakdown, batch comparison |
| Reports | (new) | Pick a register and batches: bar chart, key findings, report table, Excel download. Quality compares RM trials, FG and EX parameter by parameter |
| Daily log | (new) | Watering, turning, temperature / moisture checks, feeding, inspections |
| Raw material | 1. Raw Material Purchase | Lots, supplier, vehicle, qty, rate, moisture, lab acceptance |
| Pre-compost | 2. Pre-Composting | Lots, culture dose, turning log (date / °C / % / watering) |
| Beds | 3. Bed Production & Ops | Bed board with status colours, bulk add (BED-01…BED-80), bulk update |
| Harvest | 4. Harvesting & Packing | Raw and net yield per bed, FG batch, packing |
| Quality | 5. Quality Control | Lab reports for RM (raw material / trial), FG (finished goods) and EX (export); FG and EX checked against the FCO reference |
| Sales | 6. Sales Register | Invoices, customer, FG batch, qty, price |
| Stock | 7. Stock Inventory | Running stock ledger with closing balance |
| Expenses | 8. Cost & Expenses | Expenses by category |
| Traceability | (new) | Trace any RM lot, PC lot, bed, VB code, FG batch or invoice end to end |
| Batches | (new) | One row per production cycle (B4, B5, …) |

Records live in the app's cloud database, so everyone the app is shared with
(as Contributor or Editor) sees and edits the same data live. The **Export Excel**
button downloads the selected batch as a workbook with one sheet per register.

## Files

- `index.html` – the app (published as a claude.ai artifact with the `db`, `user` and `downloads` capabilities).
- `convert_register.py` – converts a register workbook in the Batch 4 layout into JSON records
  (`python3 convert_register.py register.xlsx B4 seed.json`). It fixes dates that Excel read as mm/dd
  and converts QC cells that were formatted as percentages back to their values.
