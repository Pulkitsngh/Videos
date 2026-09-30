"""Build the Navjyoti cloud tracker workbook (Excel for the web / SharePoint / OneDrive).

Usage: python3 build_cloud_workbook.py <Batch_register.xlsx> <out.xlsx> [batch-code]

Every register is an Excel table (named like the SharePoint lists) so Power Automate,
Power BI and SharePoint can read it. Formulas use structured references so they follow
the tables as rows are added.
"""
import datetime as dt
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.table import Table, TableFormula, TableStyleInfo

from convert_register import convert

FONT = "Arial"
GREEN, SOIL, INK, MUTED = "1C7148", "6B4424", "16201B", "58685F"
SERIES = ["2A78D6", "EB6834", "1BAF7A"]
F = lambda **k: Font(name=FONT, **{"size": 10, **k})
FILL = lambda c: PatternFill("solid", start_color=c, end_color=c)
HDR_IN, HDR_CALC = FILL(GREEN), FILL("5B6B62")
THIN = Side(style="thin", color="D3DDD4")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WARN_FILL, BAD_FILL, OK_FILL = FILL("FBECD2"), FILL("F9DEDB"), FILL("DCEFE2")
DATE_FMT, KG_FMT, INR_FMT, PCT_FMT = "dd-mm-yyyy", "#,##0", "₹#,##0", "0.0%"

# Dropdown lists (Lists sheet). Batch codes come from the Batches table.
LISTS = {
    "Product": ["RM", "FG", "EX"],
    "BedStatus": ["Filled", "Inoculated", "Active", "Ready to harvest", "Harvested", "Empty"],
    "PaymentStatus": ["Paid", "Pending", "Partial"],
    "LabResult": ["Accepted", "Awaiting", "Rejected"],
    "YesNo": ["Yes", "No"],
    "Packaging": ["Loose", "Bags 50 kg", "Bags 25 kg", "Bags 5 kg"],
    "ExpenseCategory": ["Raw material", "Culture & inputs", "Bed setup", "Labour", "Sieving", "Transport", "Machinery",
                        "Packaging", "Lab testing", "Sales & dispatch", "Utilities", "Other"],
    "PaymentMode": ["Online", "UPI", "Cash", "Cheque"],
    "Activity": ["Watering", "Turning", "Temperature check", "Moisture check", "Feeding", "Shuffling", "Worm inoculation",
                 "Pest / ant control", "Shade / cover repair", "Harvest", "Sieving", "Inspection", "Other"],
    "Stage": ["Planning", "Raw material", "Pre-composting", "In beds", "Harvesting", "Closing", "Closed"],
    "Audit": ["Verified (Match)", "Pending", "Mismatch"],
    "SalePayment": ["Received", "Pending", "Partial"],
    "Packing": ["Bulk / Loose", "Bags"],
    "FCOStatus": ["Pass (FCO Compliant)", "Awaiting report", "Fail"],
}
# FCO reference for finished vermicompost; applied to FG and EX lab reports.
FCO = [("Moisture", "Moisture %", 15, 25), ("pH", "pH", 6.5, 7.5), ("EC", "EC dS/m", None, 4), ("OrganicCarbon", "Organic carbon %", 18, None),
       ("Nitrogen", "Nitrogen %", 1.0, None), ("Phosphorus", "Phosphorus % (as P2O5)", 0.8, None), ("Potassium", "Potassium % (as K2O)", 0.8, None), ("CNRatio", "C:N ratio", None, 20)]


def d(s):
    try:
        return dt.datetime.strptime(s, "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return s or None


def this(t, c):
    return f"{t}[[#This Row],[{c}]]"


# Register definitions: (header, source key or None, kind, extra)
# kind: text | date | kg | num | inr | pct | calc ; extra: dropdown list name, or formula for calc
def registers():
    return {
        "Batches": ("batches", [("BatchCode", "code", "text", None), ("BatchName", "name", "text", None), ("Site", "site", "text", None),
                    ("StartDate", "start", "date", None), ("Stage", "status", "text", "Stage"), ("Notes", "notes", "text", None)]),
        "RawMaterial": ("rm", [("Batch", "batch", "text", "Batch"), ("PurchaseDate", "date", "date", None), ("RMLot", "lot", "text", None),
            ("Supplier", "supplier", "text", "Supplier"), ("Material", "material", "text", None), ("VehicleNo", "vehicle", "text", None),
            ("InvoiceRef", "invoice", "text", None), ("QtyKg", "qtyKg", "kg", None), ("RatePerKg", "rate", "num", None),
            ("Amount", None, "calc_inr", "=IF(OR({QtyKg}=\"\",{RatePerKg}=\"\"),\"\",{QtyKg}*{RatePerKg})"),
            ("TransportCost", "transport", "inr", None), ("PaymentStatus", "payment", "text", "PaymentStatus"), ("MoisturePct", "moisture", "num", None),
            ("TempC", "temp", "num", None), ("Odour", "odour", "text", "YesNo"), ("ContaminationCheck", "contamination", "text", "YesNo"),
            ("SampleDate", "sampleDate", "date", None), ("SentToLabDate", "labDate", "date", None), ("LabResult", "labResult", "text", "LabResult"),
            ("SealNo", "seal", "text", None), ("Remarks", "remarks", "text", None)]),
        "PreCompost": ("precompost", [("Batch", "batch", "text", "Batch"), ("PCLot", "lot", "text", None), ("RMLot", "rmLot", "text", None),
            ("StartDate", "start", "date", None), ("CultureDate", "cultureDate", "date", None), ("CultureDose", "cultureDose", "text", None),
            ("TransferDate", "end", "date", None), ("WeightMT", "weightMT", "num", None),
            *[(f"Turn{i}{s}", None, k, None) for i in (1, 2, 3) for s, k in (("Date", "date"), ("TempC", "num"), ("MoisturePct", "num"), ("WateringDate", "date"))],
            ("PeakTempC", None, "calc_num", "=IF(COUNT({Turn1TempC},{Turn2TempC},{Turn3TempC})=0,\"\",MAX({Turn1TempC},{Turn2TempC},{Turn3TempC}))"),
            ("DaysToTransfer", None, "calc_num", "=IF(OR({StartDate}=\"\",{TransferDate}=\"\"),\"\",{TransferDate}-{StartDate})"),
            ("LinkedBeds", "beds", "text", None), ("Remarks", "remarks", "text", None)]),
        "Beds": ("beds", [("Batch", "batch", "text", "Batch"), ("BedNo", "bed", "text", None), ("Status", "status", "text", "BedStatus"),
            ("Dimensions", "dims", "text", None), ("FillDate", "fillDate", "date", None), ("DungT", "dungT", "num", None), ("PCLot", "pcLot", "text", None),
            ("WormSpecies", "species", "text", None), ("WormDate", "wormDate", "date", None), ("WormsKg", "wormsKg", "num", None),
            ("FeedDate", "feedDate", "date", None), ("Watering", "watering", "text", None), ("ShuffleDate", "shuffleDate", "date", None),
            ("ExpectedHarvest", "expHarvest", "date", None), ("ActualHarvest", "actHarvest", "date", None), ("HarvestComplete", "harvestDone", "date", None),
            ("ProdCode", "prodCode", "text", None), ("Remark", "remark", "text", None),
            ("NetYieldKg", None, "calc_kg", "=IF({BedNo}=\"\",\"\",SUMIFS(Harvest[NetKg],Harvest[Batch],{Batch},Harvest[BedNo],{BedNo}))"),
            ("LiveStatus", None, "calc", "=IF({BedNo}=\"\",\"\",IF(OR(N({NetYieldKg})>0,{Status}=\"Harvested\"),\"Harvested\",IF(AND(ISNUMBER({ExpectedHarvest}),{ExpectedHarvest}<TODAY()),\"Overdue\",IF({Status}=\"\",\"Active\",{Status}))))"),
            ("BedNum", None, "calc_num", "=IFERROR(VALUE(SUBSTITUTE(UPPER({BedNo}),\"BED-\",\"\")),\"\")")]),
        "Harvest": ("harvest", [("Batch", "batch", "text", "Batch"), ("HarvestDate", "date", "date", None), ("BedNo", "bed", "text", None),
            ("RawKg", "rawKg", "kg", None), ("NetKg", "netKg", "kg", None),
            ("RecoveryPct", None, "calc_pct", "=IF(OR(N({RawKg})=0,{NetKg}=\"\"),\"\",{NetKg}/{RawKg})"),
            ("ProdBatch", "prodBatch", "text", None), ("FGBatch", "fgBatch", "text", None), ("Packaging", "packaging", "text", "Packaging"),
            ("Packets", "packets", "num", None), ("MoisturePct", "moisture", "num", None), ("Tagging", "tagging", "text", None)]),
        "QualityControl": ("qc", [("Batch", "batch", "text", "Batch"), ("SamplingDate", "date", "date", None), ("Product", "product", "text", "Product"),
            ("SealNo", "seal", "text", None), ("Lab", "lab", "text", None), ("BatchNo", "batchNo", "text", None), ("ReportDate", "reportDate", "date", None),
            ("ReportNo", "reportNo", "text", None), ("pH", "ph", "num", None), ("EC", "ec", "num", None), ("Moisture", "moisture", "num", None),
            ("OrganicCarbon", "oc", "num", None), ("OrganicMatter", "om", "num", None), ("CNRatio", "cn", "num", None), ("Nitrogen", "n", "num", None),
            ("Phosphorus", "p", "num", None), ("Potassium", "k", "num", None), ("Ash", "ash", "num", None), ("Calcium", "ca", "num", None),
            ("Manganese", "mn", "num", None), ("Iron", "fe", "num", None), ("Manganese2", "mn2", "num", None), ("Zinc", "zn", "num", None),
            ("Copper", "cu", "num", None), ("FCOStatus", "status", "text", "FCOStatus"),
            ("FCOCheck", None, "calc", "FCO")]),
        "Sales": ("sales", [("Batch", "batch", "text", "Batch"), ("InvoiceNo", "invoice", "text", None), ("SaleDate", "date", "date", None),
            ("Customer", "customer", "text", "Customer"), ("DealerDRC", "drc", "text", None), ("FGBatch", "fgBatch", "text", None),
            ("QtyKg", "qtyKg", "kg", None), ("PricePerKg", "price", "num", None),
            ("Revenue", None, "calc_inr", "=IF(OR({QtyKg}=\"\",{PricePerKg}=\"\"),\"\",{QtyKg}*{PricePerKg})"),
            ("Payment", "payment", "text", "SalePayment")]),
        "StockLedger": ("stock", [("Batch", "batch", "text", "Batch"), ("EntryNo", "seq", "num", None), ("EntryDate", "date", "date", None),
            ("Product", "product", "text", None), ("Packing", "pack", "text", "Packing"), ("InKg", "inKg", "kg", None), ("OutKg", "outKg", "kg", None),
            ("LossKg", "lossKg", "kg", None),
            ("ClosingKg", None, "calc_kg", "=IF({Batch}=\"\",\"\",SUMIFS(StockLedger[InKg],StockLedger[Batch],{Batch},StockLedger[EntryNo],\"<=\"&{EntryNo})-SUMIFS(StockLedger[OutKg],StockLedger[Batch],{Batch},StockLedger[EntryNo],\"<=\"&{EntryNo})-SUMIFS(StockLedger[LossKg],StockLedger[Batch],{Batch},StockLedger[EntryNo],\"<=\"&{EntryNo}))"),
            ("Audit", "audit", "text", "Audit"), ("Note", "note", "text", None)]),
        "Expenses": ("expenses", [("Batch", "batch", "text", "Batch"), ("ExpenseDate", "date", "date", None), ("Category", "category", "text", "ExpenseCategory"),
            ("Description", "desc", "text", None), ("PaymentMode", "mode", "text", "PaymentMode"), ("Amount", "amount", "inr", None), ("Note", "note", "text", None)]),
        "DailyLog": ("logs", [("Batch", "batch", "text", "Batch"), ("LogDate", "date", "date", None), ("Activity", "activity", "text", "Activity"),
            ("BedsArea", "area", "text", None), ("TempC", "temp", "num", None), ("MoisturePct", "moisture", "num", None), ("Qty", "qty", "num", None),
            ("RecordedBy", "by", "text", None), ("Observations", "notes", "text", None)]),
    }


def fco_formula(t):
    parts = []
    for i, (col, label, lo, hi) in enumerate(FCO):
        v = this(t, col)
        conds = []
        if lo is not None:
            conds.append(f"{v}<Lists!$S${i + 2}")
        if hi is not None:
            conds.append(f"{v}>Lists!$T${i + 2}")
        cond = conds[0] if len(conds) == 1 else f"OR({','.join(conds)})"
        parts.append(f'IF(AND(ISNUMBER({v}),{cond}),"{label.split(" ")[0]} ","")')
    flags = "&".join(parts)
    p = this(t, "Product")
    return (f'=IF({this(t, "Batch")}="","",IF(OR({p}="FG",{p}="EX"),IF({flags}="","Within FCO","Outside: "&TRIM({flags})),"Trial (not checked)"))')


def build(src, out, batch):
    data = convert(src, batch)
    data["logs"] = []
    wb = Workbook()
    wb.remove(wb.active)
    ws_readme = wb.create_sheet("How to use")
    ws_dash = wb.create_sheet("Dashboard")
    ws_rep = wb.create_sheet("Reports")
    regs = registers()
    sheets = {name: wb.create_sheet(name) for name in regs}
    ws_list = wb.create_sheet("Lists")

    # ---------- Lists ----------
    lists = dict(LISTS)
    lists["Supplier"] = sorted({r["supplier"] for r in data["rm"] if r["supplier"]})
    lists["Customer"] = sorted({r["customer"] for r in data["sales"] if r["customer"]})
    ws_list["A1"] = "BatchPick"
    ws_list["A2"] = "All"
    for i in range(3, 53):
        ws_list[f"A{i}"] = f'=IFERROR(IF(INDEX(Batches[BatchCode],{i - 2})="","",INDEX(Batches[BatchCode],{i - 2})),"")'
    col = 2
    list_ranges = {"Batch": "Batches!$A$2:$A$200", "BatchPick": "Lists!$A$2:$A$52"}
    for name, vals in lists.items():
        c = ws_list.cell(1, col, name)
        n = len(vals) + (15 if name in ("Supplier", "Customer") else 0)
        for j, v in enumerate(vals):
            ws_list.cell(j + 2, col, v)
        L = c.column_letter
        list_ranges[name] = f"Lists!${L}$2:${L}${n + 1}"
        col += 1
    # FCO table at S:T with labels in R
    ws_list["R1"], ws_list["S1"], ws_list["T1"], ws_list["U1"] = "FCO parameter", "Min", "Max", "Column"
    for i, (colname, label, lo, hi) in enumerate(FCO):
        ws_list.cell(i + 2, 18, label)
        ws_list.cell(i + 2, 19, lo)
        ws_list.cell(i + 2, 20, hi)
        ws_list.cell(i + 2, 21, colname)
    ws_list["R11"] = "Source: FCO 1985 vermicompost specification (P as P2O5, K as K2O). Edit Min / Max if your buyer's export spec differs."
    for name, ref in list_ranges.items():
        wb.defined_names[f"L_{name}"] = DefinedName(f"L_{name}", attr_text=ref)
    for c in ws_list[1]:
        c.font = F(bold=True, color="FFFFFF")
        c.fill = HDR_CALC
    ws_list.column_dimensions["A"].width = 12
    for L in "BCDEFGHIJKLMNOPQ":
        ws_list.column_dimensions[L].width = 20
    ws_list.column_dimensions["R"].width = 18
    ws_list["A1"].comment = Comment("Filled automatically from the Batches sheet. Used by the batch selectors.", "Navjyoti")

    # ---------- registers ----------
    calc_cols = {}
    for tname, (key, cols) in regs.items():
        ws = sheets[tname]
        rows = data.get(key, [])
        for ci, (h, _, kind, extra) in enumerate(cols, 1):
            c = ws.cell(1, ci, h)
            c.font = F(bold=True, color="FFFFFF")
            c.fill = HDR_CALC if kind.startswith("calc") else HDR_IN
            c.alignment = Alignment(vertical="center", wrap_text=True)
            ws.column_dimensions[c.column_letter].width = max(11, min(26, len(h) + 4))
        nrows = max(1, len(rows))
        for ri in range(nrows):
            r = rows[ri] if rows else {}
            # pre-compost turns are nested in the source
            turns = r.get("turns", []) if key == "precompost" else []
            for ci, (h, k, kind, extra) in enumerate(cols, 1):
                cell = ws.cell(ri + 2, ci)
                cell.font = F()
                if kind.startswith("calc"):
                    f = fco_formula(tname) if extra == "FCO" else extra
                    for hh, *_ in cols:
                        f = f.replace("{" + hh + "}", this(tname, hh))
                    cell.value = f
                elif h.startswith("Turn") and turns:
                    i = int(h[4]) - 1
                    if i < len(turns):
                        t = turns[i]
                        v = {"Date": t.get("date"), "TempC": t.get("temp"), "MoisturePct": t.get("moisture"), "WateringDate": t.get("watering")}[h[5:]]
                        cell.value = d(v) if kind == "date" else v
                elif k:
                    v = r.get(k)
                    cell.value = d(v) if kind == "date" else (None if v in ("", None) else v)
                cell.number_format = {"date": DATE_FMT, "kg": KG_FMT, "inr": INR_FMT, "calc_inr": INR_FMT, "calc_kg": KG_FMT,
                                      "calc_pct": PCT_FMT, "num": "0.##", "calc_num": "0.#"}.get(kind, "General")
        last = ws.cell(1, len(cols)).column_letter
        tab = Table(displayName=tname, ref=f"A1:{last}{nrows + 1}")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        tab._initialise_columns()
        for tc, (h, _, kind, extra) in zip(tab.tableColumns, cols):
            if kind.startswith("calc"):
                f = fco_formula(tname) if extra == "FCO" else extra
                for hh, *_ in cols:
                    f = f.replace("{" + hh + "}", this(tname, hh))
                tc.calculatedColumnFormula = TableFormula(attr_text=f[1:])
        ws.add_table(tab)
        ws.freeze_panes = "C2" if tname != "Batches" else "B2"
        ws.row_dimensions[1].height = 30
        calc_cols[tname] = [h for h, _, kind, _ in cols if kind.startswith("calc")]
        # dropdowns on the input columns (to row 5000 so new rows keep them)
        for ci, (h, _, kind, extra) in enumerate(cols, 1):
            if extra and not kind.startswith("calc") and extra in list_ranges:
                L = ws.cell(1, ci).column_letter
                dv = DataValidation(type="list", formula1=f"=L_{extra}", allow_blank=True,
                                    showErrorMessage=extra not in ("Supplier", "Customer"),
                                    errorTitle="Pick from the list", error="Choose a value from the drop-down. Add new options on the Lists sheet.")
                dv.add(f"{L}2:{L}5000")
                ws.add_data_validation(dv)
            if kind == "date":
                L = ws.cell(1, ci).column_letter
                dv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                                    error="Enter a date, for example 14-02-2026.", errorTitle="Date needed")
                dv.add(f"{L}2:{L}5000")
                ws.add_data_validation(dv)
        # status highlights
        hdr = [h for h, *_ in cols]
        if tname == "Beds":
            L = ws.cell(1, hdr.index("LiveStatus") + 1).column_letter
            ws.conditional_formatting.add(f"{L}2:{L}5000", CellIsRule(operator="equal", formula=['"Overdue"'], fill=WARN_FILL, font=F(bold=True, color="9A5B00")))
            ws.conditional_formatting.add(f"{L}2:{L}5000", CellIsRule(operator="equal", formula=['"Harvested"'], fill=OK_FILL))
        if tname == "QualityControl":
            L = ws.cell(1, hdr.index("FCOCheck") + 1).column_letter
            ws.conditional_formatting.add(f"{L}2:{L}5000", FormulaRule(formula=[f'LEFT({L}2,7)="Outside"'], fill=BAD_FILL, font=F(bold=True, color="B3261E")))
            ws.conditional_formatting.add(f"{L}2:{L}5000", CellIsRule(operator="equal", formula=['"Within FCO"'], fill=OK_FILL))
            ws.column_dimensions[L].width = 34
        if tname == "Harvest":
            L = ws.cell(1, hdr.index("NetKg") + 1).column_letter
            B = ws.cell(1, hdr.index("BedNo") + 1).column_letter
            ws.conditional_formatting.add(f"{L}2:{L}5000", FormulaRule(formula=[f'AND({B}2<>"",{L}2="")'], fill=WARN_FILL))
        if tname == "RawMaterial":
            L = ws.cell(1, hdr.index("PaymentStatus") + 1).column_letter
            ws.conditional_formatting.add(f"{L}2:{L}5000", FormulaRule(formula=[f'AND({L}2<>"",{L}2<>"Paid")'], fill=WARN_FILL))

    # B5 is in process: add it to Batches so it can be selected
    wsb = sheets["Batches"]
    wsb.append(["B5", "Batch 5", "Akola", None, "In beds", "In process – add dates and records as they happen"])
    for c in wsb[3]:
        c.font = F()
    wsb.tables["Batches"].ref = "A1:F3"
    wsb["D2"].number_format = wsb["D3"].number_format = DATE_FMT
    wsb.column_dimensions["F"].width = 50

    # ---------- Dashboard ----------
    ws = ws_dash
    ws.sheet_view.showGridLines = False
    for L, w in zip("ABCDEFGHIJKL", (2, 34, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16)):
        ws.column_dimensions[L].width = w
    ws["B1"] = "Navjyoti · Vermicompost production dashboard"
    ws["B1"].font = F(bold=True, size=18, color=INK)
    ws["B2"] = "Live from the register sheets. Pick a batch in C4 (or All). Everything on this sheet is a formula — do not type over it."
    ws["B2"].font = F(color=MUTED, italic=True)
    ws["B4"] = "Selected batch"
    ws["B4"].font = F(bold=True)
    ws["C4"] = "All"
    ws["C4"].fill = FILL("FFF3B0")
    ws["C4"].font = F(bold=True, size=12)
    ws["C4"].border = BOX
    dv = DataValidation(type="list", formula1="=L_BatchPick", allow_blank=False)
    dv.add("C4")
    ws.add_data_validation(dv)
    ws["D4"] = '=IF(C4="All","All batches combined",IFERROR(INDEX(Batches[BatchName],MATCH(C4,Batches[BatchCode],0))&" · "&INDEX(Batches[Stage],MATCH(C4,Batches[BatchCode],0)),""))'
    ws["D4"].font = F(color=MUTED)
    ws["H4"] = "Today"
    ws["I4"] = "=TODAY()"
    ws["I4"].number_format = DATE_FMT
    wb.defined_names["SelBatch"] = DefinedName("SelBatch", attr_text='Dashboard!$C$4')
    K = 'IF(SelBatch="All","*",SelBatch)'  # SUMIFS criterion: * matches every batch
    wb.defined_names["BatchCrit"] = DefinedName("BatchCrit", attr_text='IF(Dashboard!$C$4="All","*",Dashboard!$C$4)')

    def section(row, title):
        ws.cell(row, 2, title).font = F(bold=True, size=12, color=SOIL)
        ws.cell(row, 2).border = Border(bottom=Side(style="medium", color=SOIL))

    section(6, "KEY FIGURES")
    kpis = [
        ("Raw material received (t)", f"=SUMIFS(RawMaterial[QtyKg],RawMaterial[Batch],{K})/1000", "#,##0.0"),
        ("Raw material cost", f"=SUMIFS(RawMaterial[Amount],RawMaterial[Batch],{K})", INR_FMT),
        ("Pre-compost decomposed (MT)", f"=SUMIFS(PreCompost[WeightMT],PreCompost[Batch],{K})", "#,##0.0"),
        ("Beds filled", f'=COUNTIFS(Beds[Batch],{K},Beds[BedNo],"<>")', "#,##0"),
        ("Beds harvested", f'=COUNTIFS(Beds[Batch],{K},Beds[LiveStatus],"Harvested")', "#,##0"),
        ("Dung / compost filled in beds (t)", f"=SUMIFS(Beds[DungT],Beds[Batch],{K})", "#,##0.0"),
        ("Raw harvest (kg)", f"=SUMIFS(Harvest[RawKg],Harvest[Batch],{K})", KG_FMT),
        ("Net yield after sieving (kg)", f"=SUMIFS(Harvest[NetKg],Harvest[Batch],{K})", KG_FMT),
        ("Conversion (net ÷ dung filled)", "=IFERROR(C15/(C13*1000),\"\")", PCT_FMT),
        ("Sold (kg)", f"=SUMIFS(Sales[QtyKg],Sales[Batch],{K})", KG_FMT),
        ("Revenue", f"=SUMIFS(Sales[Revenue],Sales[Batch],{K})", INR_FMT),
        ("Average price (₹/kg)", "=IFERROR(C18/C17,\"\")", "₹#,##0.00"),
        ("Stock on hand (kg, stock ledger)", f"=SUMIFS(StockLedger[InKg],StockLedger[Batch],{K})-SUMIFS(StockLedger[OutKg],StockLedger[Batch],{K})-SUMIFS(StockLedger[LossKg],StockLedger[Batch],{K})", KG_FMT),
        ("Harvested − sold (kg)", "=C15-C17", KG_FMT),
        ("Expenses", f"=SUMIFS(Expenses[Amount],Expenses[Batch],{K})", INR_FMT),
        ("Expense per kg produced (₹)", "=IFERROR(C22/C15,\"\")", "₹#,##0.00"),
        ("Revenue − expenses", "=C18-C22", '₹#,##0;[Red]-₹#,##0'),
    ]
    for i, (label, f, fmt) in enumerate(kpis):
        r = 8 + i
        ws.cell(r, 2, label).font = F()
        c = ws.cell(r, 3, f)
        c.font = F(bold=True, size=11)
        c.number_format = fmt
        for cc in (ws.cell(r, 2), c):
            cc.border = Border(bottom=THIN)

    ws.cell(6, 5, "NEEDS ATTENTION").font = F(bold=True, size=12, color=SOIL)
    ws.cell(6, 5).border = Border(bottom=Side(style="medium", color=SOIL))
    ws.merge_cells("E7:H7")
    ws["E7"] = "Count for the selected batch · status"
    ws["E7"].font = F(color=MUTED, italic=True)
    alerts = [
        ("Beds past expected harvest", f'=COUNTIFS(Beds[Batch],{K},Beds[LiveStatus],"Overdue")'),
        ("Harvest entries without net weight", f'=COUNTIFS(Harvest[Batch],{K},Harvest[BedNo],"<>",Harvest[NetKg],"")'),
        ("Raw material payments not settled", f'=COUNTIFS(RawMaterial[Batch],{K},RawMaterial[RMLot],"<>",RawMaterial[PaymentStatus],"<>Paid")'),
        ("Raw material lots without lab acceptance", f'=COUNTIFS(RawMaterial[Batch],{K},RawMaterial[RMLot],"<>",RawMaterial[LabResult],"<>Accepted")'),
        ("FG / EX lab reports outside FCO", f'=COUNTIFS(QualityControl[Batch],{K},QualityControl[FCOCheck],"Outside*")'),
        ("Sales priced under ₹1 per kg", f'=COUNTIFS(Sales[Batch],{K},Sales[QtyKg],">0",Sales[PricePerKg],"<1")'),
        ("Days since watering was logged", f'=IF(_xlfn.MAXIFS(DailyLog[LogDate],DailyLog[Batch],{K},DailyLog[Activity],"Watering")=0,"none logged",TODAY()-_xlfn.MAXIFS(DailyLog[LogDate],DailyLog[Batch],{K},DailyLog[Activity],"Watering"))'),
    ]
    for i, (label, f) in enumerate(alerts):
        r = 8 + i
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
        ws.cell(r, 5, label).font = F()
        c = ws.cell(r, 8, f)
        c.font = F(bold=True)
        c.alignment = Alignment(horizontal="right")
        s = ws.cell(r, 9, f'=IF(H{r}="none logged","Log watering",IF(H{r}>{3 if i == 6 else 0},"Check","OK"))')
        s.font = F(bold=True)
        for cc in (ws.cell(r, 5), c, s):
            cc.border = Border(bottom=THIN)
    ws.conditional_formatting.add("I8:I14", CellIsRule(operator="equal", formula=['"OK"'], fill=OK_FILL, font=F(bold=True, color="1D7443")))
    ws.conditional_formatting.add("I8:I14", CellIsRule(operator="notEqual", formula=['"OK"'], fill=WARN_FILL, font=F(bold=True, color="9A5B00")))

    # batch comparison
    r0 = 28
    section(r0, "BATCH COMPARISON")
    heads = ["Batch", "Stage", "Beds harvested / filled", "RM received (t)", "Net yield (kg)", "Conversion", "Sold (kg)", "Revenue", "Expenses", "Revenue − expenses"]
    for j, h in enumerate(heads):
        c = ws.cell(r0 + 1, 2 + j, h)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = HDR_IN
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i in range(10):
        r = r0 + 2 + i
        b = f"B{r}"
        ws[b] = f"=Lists!A{3 + i}"
        vals = [f'=IF({b}="","",IFERROR(INDEX(Batches[Stage],MATCH({b},Batches[BatchCode],0)),""))',
                f'=IF({b}="","",COUNTIFS(Beds[Batch],{b},Beds[LiveStatus],"Harvested")&" / "&COUNTIFS(Beds[Batch],{b},Beds[BedNo],"<>"))',
                f'=IF({b}="","",SUMIFS(RawMaterial[QtyKg],RawMaterial[Batch],{b})/1000)',
                f'=IF({b}="","",SUMIFS(Harvest[NetKg],Harvest[Batch],{b}))',
                f'=IF({b}="","",IFERROR(F{r}/(SUMIFS(Beds[DungT],Beds[Batch],{b})*1000),""))',
                f'=IF({b}="","",SUMIFS(Sales[QtyKg],Sales[Batch],{b}))',
                f'=IF({b}="","",SUMIFS(Sales[Revenue],Sales[Batch],{b}))',
                f'=IF({b}="","",SUMIFS(Expenses[Amount],Expenses[Batch],{b}))',
                f'=IF({b}="","",I{r}-J{r})']
        fmts = ["General", "General", "#,##0.0", KG_FMT, PCT_FMT, KG_FMT, INR_FMT, INR_FMT, '₹#,##0;[Red]-₹#,##0']
        ws[b].font = F(bold=True)
        for j, (f, fm) in enumerate(zip(vals, fmts)):
            c = ws.cell(r, 3 + j, f)
            c.font = F()
            c.number_format = fm
            c.border = Border(bottom=THIN)
    ws.row_dimensions[r0 + 1].height = 30

    # monthly produced vs sold
    r1 = 42
    section(r1, "MONTHLY NET YIELD VS SALES (selected batch)")
    ws.cell(r1 + 1, 2, "Chart starts from month").font = F()
    ws.cell(r1 + 1, 3, dt.date(2026, 4, 1)).number_format = "mmm yyyy"
    ws.cell(r1 + 1, 3).fill = FILL("FFF3B0")
    ws.cell(r1 + 1, 3).border = BOX
    ws.cell(r1 + 1, 4, "← change to move the 12-month window").font = F(color=MUTED, italic=True)
    for j, h in enumerate(["Month", "Net yield (kg)", "Sold (kg)"]):
        c = ws.cell(r1 + 3, 2 + j, h)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = HDR_IN
    for i in range(12):
        r = r1 + 4 + i
        ws.cell(r, 2, f"=EDATE($C${r1 + 1},{i})").number_format = "mmm yy"
        ws.cell(r, 3, f'=SUMIFS(Harvest[NetKg],Harvest[Batch],{K},Harvest[HarvestDate],">="&B{r},Harvest[HarvestDate],"<"&EDATE(B{r},1))').number_format = KG_FMT
        ws.cell(r, 4, f'=SUMIFS(Sales[QtyKg],Sales[Batch],{K},Sales[SaleDate],">="&B{r},Sales[SaleDate],"<"&EDATE(B{r},1))').number_format = KG_FMT
        for j in range(3):
            ws.cell(r, 2 + j).font = F()
    ch = BarChart()
    ch.type = "col"
    ch.title = "Net yield vs sold per month (kg)"
    ch.add_data(Reference(ws, min_col=3, max_col=4, min_row=r1 + 3, max_row=r1 + 15), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=r1 + 4, max_row=r1 + 15))
    style_chart(ch, 2)
    ch.height, ch.width = 8, 18
    ws.add_chart(ch, f"F{r1 + 2}")
    ws.freeze_panes = "A5"

    build_reports(wb, ws_rep, list_ranges)
    build_readme(ws_readme)
    for w in wb.worksheets:
        w.sheet_properties.tabColor = {"How to use": SOIL, "Dashboard": GREEN, "Reports": GREEN, "Lists": "9AA79F"}.get(w.title, "C9D6CB")
    for w in wb.worksheets:
        w.page_setup.orientation = "landscape"
        w.page_setup.paperSize = w.PAPERSIZE_A4
        w.sheet_properties.pageSetUpPr.fitToPage = True
        w.page_setup.fitToWidth, w.page_setup.fitToHeight = 1, 0
    wb.calculation.fullCalcOnLoad = True
    wb.active = 1
    wb.save(out)


def style_chart(ch, n):
    for i, s in enumerate(ch.series[:n]):
        s.graphicalProperties = GraphicalProperties(solidFill=SERIES[i % 3])
        s.graphicalProperties.line.solidFill = SERIES[i % 3]
    ch.legend.position = "b"
    ch.gapWidth = 60
    ch.y_axis.delete = False
    ch.x_axis.delete = False


def build_reports(wb, ws, list_ranges):
    ws.sheet_view.showGridLines = False
    for L, w in zip("ABCDEFG", (2, 30, 15, 15, 15, 15, 3)):
        ws.column_dimensions[L].width = w
    ws["B1"] = "Reports · compare batches"
    ws["B1"].font = F(bold=True, size=18, color=INK)
    ws["B2"] = "Pick up to three batches in C4:E4 and the first month in C5. Click a report name below to jump to it. Every table and chart updates live."
    ws["B2"].font = F(color=MUTED, italic=True)
    ws["B4"], ws["B5"], ws["B6"] = "Compare batches", "Months start from", "Quality report batch"
    for c in (ws["B4"], ws["B5"], ws["B6"]):
        c.font = F(bold=True)
    ws["C4"], ws["D4"], ws["E4"] = "B4", "B5", None
    ws["C5"] = dt.date(2025, 12, 1)
    ws["C5"].number_format = "mmm yyyy"
    ws["C6"] = "All"
    for a in ("C4", "D4", "E4", "C5", "C6"):
        ws[a].fill = FILL("FFF3B0")
        ws[a].border = BOX
        ws[a].font = F(bold=True)
    dv = DataValidation(type="list", formula1="=L_Batch", allow_blank=True)
    dv.add("C4:E4")
    ws.add_data_validation(dv)
    dv2 = DataValidation(type="list", formula1="=L_BatchPick", allow_blank=False)
    dv2.add("C6")
    ws.add_data_validation(dv2)

    sections = []
    row = 13

    def table(title, short, cat_label, cats, fn, fmt, note=None, chart_title=None, total=None):
        """cats: list of (label or formula, criteria builder). fn(batch_cell, cat_row) -> formula"""
        nonlocal row
        start = row
        ws.cell(row, 2, title).font = F(bold=True, size=13, color=SOIL)
        ws.cell(row, 2).border = Border(bottom=Side(style="medium", color=SOIL))
        sections.append((short, row))
        if note:
            ws.cell(row + 1, 2, note).font = F(color=MUTED, italic=True)
        hr = row + 2
        heads = [cat_label, "=IF($C$4=\"\",\"(pick)\",$C$4)", "=IF($D$4=\"\",\"(none)\",$D$4)", "=IF($E$4=\"\",\"(none)\",$E$4)", "Total"]
        for j, h in enumerate(heads):
            c = ws.cell(hr, 2 + j, h)
            c.font = F(bold=True, color="FFFFFF")
            c.fill = HDR_IN
        for i, cat in enumerate(cats):
            r = hr + 1 + i
            ws.cell(r, 2, cat).font = F()
            if isinstance(cat, str) and cat.startswith("=EDATE"):
                ws.cell(r, 2).number_format = "mmm yy"
            for j, bc in enumerate(("$C$4", "$D$4", "$E$4")):
                c = ws.cell(r, 3 + j, f'=IF({bc}="","",{fn(bc, r)})')
                c.number_format = fmt
                c.font = F()
            t = ws.cell(r, 6, f"=SUM(C{r}:E{r})")
            t.number_format = fmt
            t.font = F(bold=True)
            for j in range(5):
                ws.cell(r, 2 + j).border = Border(bottom=THIN)
        end = hr + len(cats)
        tr = end + 1
        ws.cell(tr, 2, total[0] if total else "Total").font = F(bold=True)
        for j in range(4):
            L = "CDEF"[j]
            c = ws.cell(tr, 3 + j, total[1].format(L=L, a=hr + 1, b=hr + 2, c=hr + 3) if total else f"=SUM({L}{hr + 1}:{L}{end})")
            c.number_format = fmt
            c.font = F(bold=True)
            c.fill = FILL("E8EEE7")
        ws.cell(tr, 2).fill = FILL("E8EEE7")
        ch = BarChart()
        ch.type = "bar"
        ch.title = chart_title or title
        ch.add_data(Reference(ws, min_col=3, max_col=5, min_row=hr, max_row=end), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=2, min_row=hr + 1, max_row=end))
        ch.y_axis.numFmt = "#,##0"
        ch.x_axis.scaling.orientation = "maxMin"
        style_chart(ch, 3)
        ch.height = max(7, 0.55 * len(cats) + 3)
        ch.width = 17
        ws.add_chart(ch, f"H{start}")
        row = max(tr + 3, start + int(ch.height * 2) + 3)
        return hr, end

    months = [f"=EDATE($C$5,{i})" for i in range(12)]
    mrange = lambda tb, dc, r: f'{tb}[{dc}],">="&$B${r},{tb}[{dc}],"<"&EDATE($B${r},1)'
    sup = list_ranges["Supplier"]
    nsup = int(sup.split("$")[-1]) - 1
    table("Raw material – quantity by supplier (kg)", "RM by supplier", "Supplier", [f"=IF(INDEX(L_Supplier,{i + 1})=\"\",\"\",INDEX(L_Supplier,{i + 1}))" for i in range(min(nsup, 10))],
          lambda b, r: f"SUMIFS(RawMaterial[QtyKg],RawMaterial[Batch],{b},RawMaterial[Supplier],$B${r})", KG_FMT,
          "Supplier names come from the Lists sheet. Add new suppliers there.")
    table("Raw material – quantity by purchase month (kg)", "RM by month", "Month", months,
          lambda b, r: f"SUMIFS(RawMaterial[QtyKg],RawMaterial[Batch],{b},{mrange('RawMaterial', 'PurchaseDate', r)})", KG_FMT)
    table("Pre-compost – weight decomposed by start month (MT)", "Pre-compost", "Month", months,
          lambda b, r: f"SUMIFS(PreCompost[WeightMT],PreCompost[Batch],{b},{mrange('PreCompost', 'StartDate', r)})", "#,##0.0")
    blocks = [f"BED-{s:02d}–{s + 9:02d}" for s in range(1, 100, 10)]
    table("Beds – net yield by bed block (kg)", "Beds by block", "Bed block", blocks,
          lambda b, r: f'SUMIFS(Beds[NetYieldKg],Beds[Batch],{b},Beds[BedNum],">="&VALUE(MID($B${r},5,2)),Beds[BedNum],"<="&VALUE(MID($B${r},5,2))+9)', KG_FMT)
    table("Harvest – net yield by month (kg)", "Harvest by month", "Month", months,
          lambda b, r: f"SUMIFS(Harvest[NetKg],Harvest[Batch],{b},{mrange('Harvest', 'HarvestDate', r)})", KG_FMT)
    table("Sales – quantity by customer (kg)", "Sales by customer", "Customer", [f"=IF(INDEX(L_Customer,{i + 1})=\"\",\"\",INDEX(L_Customer,{i + 1}))" for i in range(8)],
          lambda b, r: f"SUMIFS(Sales[QtyKg],Sales[Batch],{b},Sales[Customer],$B${r})", KG_FMT, "Customer names come from the Lists sheet.")
    table("Sales – revenue by month (₹)", "Revenue by month", "Month", months,
          lambda b, r: f"SUMIFS(Sales[Revenue],Sales[Batch],{b},{mrange('Sales', 'SaleDate', r)})", INR_FMT)
    table("Stock – movement (kg)", "Stock", "Movement", ["In (production)", "Out (sales)", "Loss / adjustment"],
          lambda b, r: f'SUMIFS(CHOOSE(MATCH($B${r},{{"In (production)","Out (sales)","Loss / adjustment"}},0),StockLedger[InKg],StockLedger[OutKg],StockLedger[LossKg]),StockLedger[Batch],{b})', KG_FMT,
          total=("Closing stock", "=N({L}{a})-N({L}{b})-N({L}{c})"))
    table("Expenses by category (₹)", "Expenses", "Category", [f"=INDEX(L_ExpenseCategory,{i + 1})" for i in range(len(LISTS['ExpenseCategory']))],
          lambda b, r: f"SUMIFS(Expenses[Amount],Expenses[Batch],{b},Expenses[Category],$B${r})", INR_FMT)
    table("Daily log – entries by activity", "Daily log", "Activity", [f"=INDEX(L_Activity,{i + 1})" for i in range(len(LISTS['Activity']))],
          lambda b, r: f"COUNTIFS(DailyLog[Batch],{b},DailyLog[Activity],$B${r})", "#,##0")

    # Quality: RM trial vs FG vs EX
    start = row
    ws.cell(row, 2, "Quality – RM trial vs FG vs EX (average of lab reports)").font = F(bold=True, size=13, color=SOIL)
    ws.cell(row, 2).border = Border(bottom=Side(style="medium", color=SOIL))
    sections.append(("Quality RM / FG / EX", row))
    ws.cell(row + 1, 2, "Batch chosen in C6 (All = every batch). FCO limits apply to FG and EX; cells in red fall outside them.").font = F(color=MUTED, italic=True)
    hr = row + 2
    for j, h in enumerate(["Parameter", "RM trial", "FG finished", "EX export", "FCO min", "FCO max"]):
        c = ws.cell(hr, 2 + j, h)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = HDR_IN
    params = [("Moisture", "Moisture %"), ("OrganicCarbon", "Organic carbon %"), ("OrganicMatter", "Organic matter %"), ("Ash", "Ash %"), ("CNRatio", "C:N ratio"),
              ("pH", "pH"), ("EC", "EC dS/m"), ("Nitrogen", "Nitrogen %"), ("Phosphorus", "Phosphorus %"), ("Potassium", "Potassium %"),
              ("Calcium", "Calcium"), ("Manganese", "Manganese"), ("Iron", "Iron"), ("Manganese2", "Manganese (2nd)"), ("Zinc", "Zinc"), ("Copper", "Copper")]
    crit = 'IF($C$6="All","*",$C$6)'
    fco_idx = {c: i for i, (c, *_) in enumerate(FCO)}
    for i, (colname, label) in enumerate(params):
        r = hr + 1 + i
        ws.cell(r, 2, label).font = F()
        for j, p in enumerate(("RM", "FG", "EX")):
            c = ws.cell(r, 3 + j, f'=IFERROR(AVERAGEIFS(QualityControl[{colname}],QualityControl[Batch],{crit},QualityControl[Product],"{p}"),"")')
            c.number_format = "0.00"
            c.font = F()
        if colname in fco_idx:
            k = fco_idx[colname] + 2
            ws.cell(r, 6, f'=IF(Lists!$S${k}="","",Lists!$S${k})').font = F(color=MUTED)
            ws.cell(r, 7, f'=IF(Lists!$T${k}="","",Lists!$T${k})').font = F(color=MUTED)
        for j in range(6):
            ws.cell(r, 2 + j).border = Border(bottom=THIN)
    qend = hr + len(params)
    rr = qend + 1
    ws.cell(rr, 2, "Reports included").font = F(color=MUTED)
    for j, p in enumerate(("RM", "FG", "EX")):
        ws.cell(rr, 3 + j, f'=COUNTIFS(QualityControl[Batch],{crit},QualityControl[Product],"{p}")').font = F(color=MUTED)
    ws.column_dimensions["G"].width = 10
    rng = f"D{hr + 1}:E{qend}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER(D{hr + 1}),OR(AND(ISNUMBER($F{hr + 1}),D{hr + 1}<$F{hr + 1}),AND(ISNUMBER($G{hr + 1}),D{hr + 1}>$G{hr + 1})))'],
                                                   fill=BAD_FILL, font=F(bold=True, color="B3261E")))
    for title, a, b, anchor in (("Quality – moisture, carbon, ash, C:N", hr + 1, hr + 5, f"I{start}"), ("Quality – pH, EC, N, P, K", hr + 6, hr + 10, f"I{start + 17}")):
        ch = BarChart()
        ch.type = "col"
        ch.title = title
        for j in range(3):
            s = Series(Reference(ws, min_col=3 + j, min_row=a, max_row=b), title=["RM trial", "FG finished", "EX export"][j])
            ch.series.append(s)
        ch.set_categories(Reference(ws, min_col=2, min_row=a, max_row=b))
        style_chart(ch, 3)
        ch.height, ch.width = 8, 16
        ws.add_chart(ch, anchor)
    row = max(rr + 3, start + 34)

    # index with links
    ws["B8"] = "Jump to:"
    ws["B8"].font = F(bold=True)
    col = 3
    r = 8
    for title, rr_ in sections:
        c = ws.cell(r, col, title)
        c.hyperlink = Hyperlink(ref=c.coordinate, location=f"Reports!B{rr_}", display=c.value)
        c.font = F(color="1C5FB8", underline="single")
        col += 1
        if col > 6:
            col, r = 3, r + 1
    ws.freeze_panes = "A7"


def build_readme(ws):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 100
    lines = [
        ("h", "Navjyoti Vermicompost Tracker – cloud workbook"),
        ("p", "One workbook for every batch. Keep it on OneDrive or SharePoint and open it in Excel for the web so the whole team works in the same file, live."),
        ("h2", "Put it on the cloud (once)"),
        ("1", "Upload this file to your Navjyoti SharePoint site (Documents) or OneDrive for Business."),
        ("2", "Open it in the browser (Excel for the web). Click Share → give supervisors 'Can edit', managers 'Can view'."),
        ("3", "Supervisors open the same link on phone (Excel app) or PC. Everyone's changes appear within seconds (co-authoring)."),
        ("4", "Optional: in Teams, add a tab → Excel → pick this file, so the team finds it in the Navjyoti channel."),
        ("5", "Optional: Power BI (Get data → Excel workbook on SharePoint) and Power Automate (Excel Online (Business) → List rows present in a table) can read every register, because each one is an Excel table."),
        ("h2", "Daily use"),
        ("•", "Type new records in the first empty row directly under a table. The table grows by itself and grey-header columns fill in their formulas."),
        ("•", "Always fill the Batch column (drop-down). Start Batch 5 records with B5."),
        ("•", "Use drop-downs where they appear. To add a new supplier, customer or category, type it on the Lists sheet."),
        ("•", "Dates: type as 14-02-2026 or pick from the date picker."),
        ("•", "Dashboard: choose a batch (or All) in cell C4. Reports: choose up to three batches in C4:E4 to compare, and the quality batch in C6."),
        ("h2", "Colour legend"),
        ("green", "Green header = you type here."),
        ("grey", "Grey header = formula column. Do not type in it; it calculates by itself."),
        ("yellow", "Yellow cell = a selector you can change (batch, month)."),
        ("h2", "Sheets"),
        ("Dashboard", "Key figures, alerts that need attention, batch comparison, monthly net yield vs sales."),
        ("Reports", "Bar charts and tables by supplier, month, bed block, customer, category and activity, for up to three batches side by side; Quality RM vs FG vs EX comparison against FCO limits."),
        ("Batches", "One row per batch (B4, B5, …). Batch drop-downs everywhere read from here."),
        ("RawMaterial", "Excel sheet 1 – purchases, supplier, vehicle, qty, rate, moisture, lab acceptance. Amount is calculated."),
        ("PreCompost", "Excel sheet 2 – lots, culture dose, up to three turnings. Peak temperature and days to transfer are calculated."),
        ("Beds", "Excel sheet 3 – bed lifecycle. Net yield (from Harvest), live status (Harvested / Overdue / …) and bed number are calculated."),
        ("Harvest", "Excel sheet 4 – raw and net yield per bed, FG batch, packing. Recovery % is calculated."),
        ("QualityControl", "Excel sheet 5 – lab reports. Product: RM = raw material / trial before final material, FG = finished goods, EX = export material. FCO check is calculated for FG and EX."),
        ("Sales", "Excel sheet 6 – invoices. Revenue is calculated."),
        ("StockLedger", "Excel sheet 7 – stock in / out / loss. Closing stock runs per batch in entry-number order."),
        ("Expenses", "Excel sheet 8 – expenses by category."),
        ("DailyLog", "Watering, turning, temperature and moisture checks, feeding, inspections. Example: B5 · 01-10-2026 · Watering · BED-01 to BED-40 · 400 L · Ramesh."),
        ("Lists", "Drop-down values and the FCO reference limits (editable)."),
        ("h2", "Notes on the imported Batch 4 data"),
        ("•", "Dates that Excel had read as month/day were corrected (for example 02-09-2026 for bed 51's harvest)."),
        ("•", "Quality values were stored as percentages in the old register; they are now plain numbers (pH 7.22, not 7.22%)."),
        ("•", "Please check: beds 59 and 61–66 have no net weight, invoice CBWS/0007 is priced at ₹0.001/kg."),
    ]
    r = 2
    for kind, text in lines:
        if kind == "h":
            ws.cell(r, 2, text).font = F(bold=True, size=18, color=INK)
        elif kind == "h2":
            r += 1
            ws.cell(r, 2, text).font = F(bold=True, size=12, color=SOIL)
            ws.cell(r, 2).border = Border(bottom=Side(style="medium", color=SOIL))
        elif kind == "p":
            ws.cell(r, 2, text).font = F(color=MUTED)
        else:
            label = {"green": "", "grey": "", "yellow": ""}.get(kind, kind)
            c = ws.cell(r, 2, label)
            c.font = F(bold=True)
            if kind in ("green", "grey", "yellow"):
                c.fill = {"green": HDR_IN, "grey": HDR_CALC, "yellow": FILL("FFF3B0")}[kind]
            t = ws.cell(r, 3, text)
            t.font = F()
            t.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1


if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "B4")
