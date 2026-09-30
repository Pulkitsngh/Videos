# Navjyoti Vermicompost on Microsoft 365

Step-by-step setup: **SharePoint lists** hold the data, **Power Apps** is the data-entry app on phone and PC,
**Power BI** gives the live dashboards and reports, **Power Automate** sends alerts. Everyone signs in with their
normal company Microsoft 365 account; no other accounts are needed.

Total effort: about 1–2 days for one person who knows Microsoft 365 basics.

## What is in this kit

| File | Use |
|---|---|
| `Create-NavjyotiLists.ps1` | Creates the 10 SharePoint lists with the right column types and loads the Batch 4 data |
| `data/*.csv` | Batch 4 records, one file per list (read by the script) |
| `PowerBI-Measures.dax` | Calculated columns and measures to paste into Power BI |
| `PowerApps-Formulas.md` | Screens and formulas for the data-entry app, including pop-up detail panels |
| `PowerAutomate-Flows.md` | The four alert flows, step by step |

The 10 lists: **Batches, RawMaterial, PreCompost, Beds, Harvest, QualityControl, Sales, StockLedger, Expenses, DailyLog**.
Column names are the same as in the Excel tracker, so reports and formulas line up.

---

## Stage 1 – SharePoint site (IT admin, 10 min)

1. office.com → **SharePoint** → **+ Create site** → **Team site** → name it **Navjyoti Production**.
2. Add supervisors as **Members** (can edit) and managers as **Visitors** (view only).
3. Note the site address, e.g. `https://yourcompany.sharepoint.com/sites/NavjyotiProduction`.

## Stage 2 – Create the lists and load Batch 4 (IT admin, 20 min)

On a Windows PC:

1. Install **PowerShell 7** (Microsoft Store) and open it.
2. `Install-Module PnP.PowerShell -Scope CurrentUser`
3. One-time app registration (needs a Microsoft 365 admin):
   `Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP Navjyoti" -Tenant yourcompany.onmicrosoft.com`
   and copy the **Client ID** it prints.
   Replace `yourcompany` with the name in your SharePoint address (`https://yourcompany.sharepoint.com`).
   A browser window opens: sign in as admin and accept the permissions. Copy the **Client ID** shown at the end.
4. Go to the unzipped kit folder (not `C:\Windows\System32`) and unblock the files, for example:
   `cd "$env:USERPROFILE\Downloads\Navjyoti_M365_Kit"; Get-ChildItem -Recurse | Unblock-File`
5. Run (replace `yourcompany` and `PASTE-CLIENT-ID-HERE` with your values; no < > brackets):
   `./Create-NavjyotiLists.ps1 -SiteUrl https://yourcompany.sharepoint.com/sites/NavjyotiProduction -ClientId PASTE-CLIENT-ID-HERE`
6. Check the counts on the site (Site contents): RawMaterial 15, PreCompost 3, Beds 80, Harvest 80, QualityControl 6,
   Sales 6, StockLedger 63, Expenses 30, Batches 2 (B4, B5), DailyLog 0.

The script can be run again safely: it keeps existing lists, adds missing columns, and never loads data into a list that already has rows.

No PowerShell available? Create each list with **+ New → List → From CSV** using the files in `data/`, then set the
column types by hand: dates → *Date and time (date only)*, quantities → *Number*, and the drop-down columns → *Choice*.

**Calculated in SharePoint:** RawMaterial *Amount*, Sales *Revenue*, Harvest *RecoveryPct*, PreCompost *DaysToTransfer*.
Figures that need another list (bed net yield, live bed status, closing stock, FCO check) are calculated in Power BI and the app.

## Stage 3 – Data-entry app in Power Apps (2–3 hours)

Follow `PowerApps-Formulas.md`. In short:

1. **make.powerapps.com → + Create → Start with data → SharePoint →** your site → **DailyLog**. Power Apps builds a
   working phone app with list, detail and edit screens.
2. **Data → + Add data → SharePoint →** add Beds, Harvest, QualityControl, RawMaterial, Sales, Expenses, StockLedger, PreCompost, Batches.
3. Add the Home screen (batch picker plus one tile per register), the bed board, the harvest and quality forms, and the
   pop-up detail panels, using the formulas in the guide.
4. **File → Save → Publish → Share** with the Navjyoti Production members group.
5. Supervisors install **Power Apps** from the Play Store / App Store and sign in with their company account.

## Stage 4 – Reports in Power BI (3–4 hours)

1. **Power BI Desktop → Get data → SharePoint Online list →** site URL → **Implementation 2.0** → tick all 10 lists → **Transform data**.
2. In Power Query, for each list keep the columns shown in the Excel tracker and set types (dates → Date, quantities → Decimal number).
   If a choice column arrives as a record, expand it to its *Value*.
3. **Close & apply.** In Model view link **Batches[BatchCode] → *every list*[Batch]** (one-to-many).
4. **New table → Calendar** and the calculated columns and measures from `PowerBI-Measures.dax`; link Calendar[Date] to
   Harvest[HarvestDate], Sales[SaleDate], RawMaterial[PurchaseDate], Expenses[ExpenseDate] and DailyLog[LogDate].
5. Build the pages (all with a **Batch slicer** that syncs across pages):
   - **Overview:** cards for RM received, Net yield, Conversion %, Sold, Revenue, Stock on hand, Expenses, Margin; clustered column *Net yield kg* and *Sold kg* by month; bar *Expenses* by category; bar *Beds* by LiveStatus.
   - **Raw material:** bar Quantity by Supplier; column by month; table of lots with lab result.
   - **Beds & harvest:** bar Net yield by BedBlock; column Net yield by month; matrix Beds × status.
   - **Quality (RM / FG / EXRM / EXFG):** clustered column of Avg Moisture, OC, C:N, N, P, K with **Product** as legend; constant lines at the FCO limits; table of lab reports with FCOCheck.
   - **Sales & stock:** bar by Customer; revenue by month; stock in / out / loss.
   - **Compare batches:** matrix with Batch on columns and the main measures on rows.
6. Tooltips and pop-ups: in each visual turn on **Tooltips**; add a **tooltip page** (Format → Page information → Tooltip) showing bed details,
   and a **drill-through page** so right-click → *Drill through* on a batch or bed opens its full detail.
7. **Publish** to a workspace (e.g. *Navjyoti Production*). In the Power BI service: **Semantic model → Settings → Data source
   credentials → OAuth2 (organisational account)** and **Scheduled refresh** up to 8 times a day.

## Stage 5 – Alerts in Power Automate (1 hour)

Follow `PowerAutomate-Flows.md`: overdue beds, watering gap, FG/EXFG lab report outside FCO, unpaid raw material.

## Stage 6 – Put it together in Teams (10 min)

In the **Navjyoti Production** team, click **+** on the channel tab bar and add:
**Power Apps** (the app) · **Power BI** (the report) · **Lists** (Beds, Harvest) for quick checks.

## Stage 7 – Go live with Batch 5

1. Add B5 details in the **Batches** list (start date, stage).
2. Supervisors enter everything in the app from now on.
3. Run the Excel tracker in parallel for 1–2 weeks and compare totals, then use Microsoft 365 as the main system.

## Licences

Most Microsoft 365 business plans (Business Standard/Premium, E3, E5) include SharePoint, Power Apps and Power Automate
for SharePoint data. Sharing Power BI reports needs **Power BI Pro** for each viewer (included in E5). Confirm with your IT admin.
