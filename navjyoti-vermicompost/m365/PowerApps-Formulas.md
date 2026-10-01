# Navjyoti app – Power Apps screens and formulas

Canvas app, **phone layout**, data sources: the 9 SharePoint lists.
Choice columns (Batch, Status, Product, PaymentStatus …) are records in Power Apps, so compare with `.Value`.

## App

**App.OnStart**
```
Set(varBatch, "B5");
Set(varShowDetail, false)
```

## 1. Home screen (`scrHome`)

| Control | Property | Formula |
|---|---|---|
| `ddBatch` (Dropdown) | Items | `Sort(Batches, StartDate, SortOrder.Descending).BatchCode` |
| | Default | `varBatch` |
| | OnChange | `Set(varBatch, ddBatch.Selected.BatchCode)` |
| `lblBedsSummary` (Label) | Text | `CountRows(Filter(Beds, Batch.Value = varBatch)) & " beds · " & CountRows(Filter(Beds, Batch.Value = varBatch, Status.Value = "Harvested")) & " harvested"` |
| `lblNet` (Label) | Text | `"Net yield: " & Text(Sum(Filter(Harvest, Batch.Value = varBatch), NetKg), "#,##0") & " kg"` |
| Tile buttons | OnSelect | `Navigate(scrWorms)`, `Navigate(scrBeds)`, `Navigate(scrHarvest)`, `Navigate(scrQC)`, `Navigate(scrRM)`, `Navigate(scrSales)`, `Navigate(scrExpenses)` |

`Sum` and `CountRows` on SharePoint are limited to the first 2 000 rows per batch (delegation). That is ample for one batch.

## 2. Earthworm purchases (`scrWorms`) – quick entry

Earthworm purchases are kept outside every batch (no Batch column).

| Control | Property | Formula |
|---|---|---|
| `txtInvoice`, `txtSupplier`, `txtSpecies`, `txtKg`, `txtRate`, `txtTransport` | — | text inputs (numbers: Format = Number) |
| `ddPay` | Items | `Choices(Earthworm.PaymentStatus)` |
| `btnSave` | OnSelect | see below |
| `galWorms` | Items | `Sort(Earthworm, PurchaseDate, SortOrder.Descending)` |
| `lblWormTotal` | Text | `Text(Sum(Earthworm, WormsKg), "#,##0") & " kg · ₹" & Text(Sum(Earthworm, WormsKg * RatePerKg) + Sum(Earthworm, TransportCost), "#,##0")` |

```
Patch(Earthworm, Defaults(Earthworm), {
    PurchaseDate: Today(),
    InvoiceNo: txtInvoice.Text,
    Supplier: txtSupplier.Text,
    Species: txtSpecies.Text,
    WormsKg: Value(txtKg.Text),
    RatePerKg: Value(txtRate.Text),
    TransportCost: Value(txtTransport.Text),
    PaymentStatus: ddPay.Selected
});
Reset(txtInvoice); Reset(txtKg); Reset(txtRate); Reset(txtTransport);
Notify("Earthworm purchase saved", NotificationType.Success)
```

## 3. Bed board (`scrBeds`) with pop-up detail

| Control | Property | Formula |
|---|---|---|
| `galBeds` (Blank flexible-height gallery, **Wrap count 4**) | Items | `SortByColumns(Filter(Beds, Batch.Value = varBatch), "BedNo")` |
| `recStatus` (Rectangle inside the gallery, left strip) | Fill | `Switch(true, ThisItem.Status.Value = "Harvested", RGBA(107,68,36,1), !IsBlank(ThisItem.ExpectedHarvest) && ThisItem.ExpectedHarvest < Today(), RGBA(232,151,0,1), ThisItem.Status.Value = "Filled", RGBA(91,91,214,1), RGBA(30,138,90,1))` |
| `lblBed` | Text | `ThisItem.BedNo & Char(10) & If(ThisItem.Status.Value <> "Harvested" && ThisItem.ExpectedHarvest < Today(), "Overdue " & DateDiff(ThisItem.ExpectedHarvest, Today()) & "d", ThisItem.Status.Value)` |
| `galBeds` | OnSelect | `Set(varBed, ThisItem); Set(varShowDetail, true)` |

**Pop-up panel** – a `Container` (`conDetail`) on top of the gallery:

| Control | Property | Formula |
|---|---|---|
| `conDetail` | Visible | `varShowDetail` |
| | Fill | `RGBA(0,0,0,0.45)` (dims the screen behind) |
| `frmBed` (Edit form, inside a white rectangle in the container) | DataSource / Item | `Beds` / `varBed` |
| `lblBedNet` | Text | `"Net yield: " & Text(Sum(Filter(Harvest, Batch.Value = varBed.Batch.Value, BedNo = varBed.BedNo), NetKg), "#,##0") & " kg"` |
| `btnSaveBed` | OnSelect | `SubmitForm(frmBed); Set(varShowDetail, false)` |
| `btnClose` | OnSelect | `Set(varShowDetail, false)` |

Use the same pattern (gallery `OnSelect` → set a variable → show a container) for pop-ups on any list.

## 4. Harvest entry (`scrHarvest`)

| Control | Property | Formula |
|---|---|---|
| `frmHarvest` | DataSource | `Harvest` |
| | Item | `galHarvest.Selected` |
| Batch card | Default | `If(frmHarvest.Mode = FormMode.New, {Value: varBatch}, ThisItem.Batch)` |
| BedNo card (make it a Dropdown) | Items | `SortByColumns(Filter(Beds, Batch.Value = varBatch, Status.Value <> "Harvested"), "BedNo").BedNo` |
| `btnNew` | OnSelect | `NewForm(frmHarvest)` |
| `btnSubmit` | OnSelect | `SubmitForm(frmHarvest)` |
| `frmHarvest` | OnSuccess | see below |

```
// mark the bed as harvested when a net weight is saved
If(!IsBlank(frmHarvest.LastSubmit.NetKg),
    Patch(Beds,
        LookUp(Beds, Batch.Value = frmHarvest.LastSubmit.Batch.Value && BedNo = frmHarvest.LastSubmit.BedNo),
        {Status: {Value: "Harvested"}, HarvestComplete: frmHarvest.LastSubmit.HarvestDate}));
Notify("Harvest saved", NotificationType.Success);
ResetForm(frmHarvest)
```

## 5. Quality (`scrQC`) – RM, FG, EXRM, EXFG

| Control | Property | Formula |
|---|---|---|
| `frmQC` | DataSource / Item | `QualityControl` / `galQC.Selected` |
| Product card | Items | `Choices(QualityControl.Product)` (RM, FG, EXRM, EXFG) |
| `lblFCO` (warning label on the form) | Text | see below |
| | Color | `If(StartsWith(lblFCO.Text, "Outside"), Color.Red, Color.Green)` |

```
With({p: DataCardValueProduct.Selected.Value,
      m: Value(DataCardValueMoisture.Text), n: Value(DataCardValueNitrogen.Text),
      ph: Value(DataCardValuepH.Text), cn: Value(DataCardValueCNRatio.Text),
      ph2: Value(DataCardValuePhosphorus.Text), k: Value(DataCardValuePotassium.Text),
      oc: Value(DataCardValueOrganicCarbon.Text), ec: Value(DataCardValueEC.Text)},
    If(!(p in ["FG", "EXFG"]), "Trial (RM / EXRM) – not checked against FCO",
        With({f: Concatenate(
                If(m < 15 || m > 25, "Moisture ", ""), If(ph < 6.5 || ph > 7.5, "pH ", ""),
                If(ec > 4, "EC ", ""), If(oc < 18, "OC ", ""), If(n < 1, "N ", ""),
                If(ph2 < 0.8, "P ", ""), If(k < 0.8, "K ", ""), If(cn > 20, "C:N ", ""))},
            If(f = "", "Within FCO", "Outside FCO: " & Trim(f)))))
```
(The control names `DataCardValue…` differ in each app; pick the text input inside each card.)

**Compare RM / FG / EXRM / EXFG** – three labels per parameter:
`Text(Average(Filter(QualityControl, Batch.Value = varBatch, Product.Value = "EXFG"), Moisture), "0.00")`

## 6. Raw material, Sales, Expenses

Use **+ New screen → Form** or the generated browse/edit screens. For each, set:
- Gallery Items: `Sort(Filter(<List>, Batch.Value = varBatch), <DateColumn>, SortOrder.Descending)`
- Batch card Default: `If(<Form>.Mode = FormMode.New, {Value: varBatch}, ThisItem.Batch)`

## 7. Share

**File → Save → Publish → Share →** Navjyoti Production members. The SharePoint site permissions decide who can edit.
