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
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RichTextProperties
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
import re

from openpyxl.worksheet.filters import AutoFilter, FilterColumn

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from convert_register import convert

FONT = "Arial"
GREEN, SOIL, INK, MUTED = "1C7148", "6B4424", "16201B", "58685F"
SERIES = ["2A78D6", "EB6834", "1BAF7A"]
# Quality product types: code, label, colour (RM/EXRM are trials; FG/EX/EXFG are checked against FCO)
PRODUCTS = [("RM", "RM trial", "2A78D6"), ("FG", "FG finished", "EB6834"),
            ("EXRM", "EXRM export trial", "1BAF7A"), ("EXFG", "EXFG export finished", "C98500")]
FINISHED = ("FG", "EXFG")
MATERIALS = ["SMC", "Cow Dung", "Cow Dung Slurry", "Coco Peat", "FOM", "Press Mud", "Crop Residue", "Others"]
# Worm breeding unit: dung used to multiply earthworms is booked here, not to any batch
BREED_ENTRY = ["Earthworm purchase", "Dung / feed added", "Worms harvested", "Worms issued to beds", "Material out"]
BREED_DEST = ["Stock for sale (after mixing)", "Mixed into a batch", "New beds / worm stock", "Discarded"]
KG_AXIS = '[>=1000000]#,##0.0,,"M";[>=1000]#,##0,"k";0'
INR_AXIS = '[>=10000000]"₹"#,##0.0,,,"Cr";[>=100000]"₹"#,##0.0,,"L";[>=1000]"₹"#,##0,"k";"₹"0'
PASSWORD = "Navjyoti2026"  # sheet protection password (told to the owner; change in Review → Unprotect Sheet)
CAPACITY = {"Batches": 200, "DailyLog": 5000}  # rows ready for entry per register; others use PREFILL
F = lambda **k: Font(name=FONT, **{"size": 10, **k})
FILL = lambda c: PatternFill("solid", start_color=c, end_color=c)
HDR_IN, HDR_CALC = FILL(GREEN), FILL("5B6B62")
THIN = Side(style="thin", color="D3DDD4")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WARN_FILL, BAD_FILL, OK_FILL = FILL("FBECD2"), FILL("F9DEDB"), FILL("DCEFE2")
PREFILL, LAST_ROW = 1500, 5000  # formula rows ready below each table / rows the reports read
DATE_FMT, KG_FMT, INR_FMT, PCT_FMT = "dd-mm-yyyy", "#,##0", "₹#,##0", "0.0%"

# Dropdown lists (Lists sheet). Batch codes come from the Batches table.
LISTS = {
    "Product": [p[0] for p in PRODUCTS],
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
    "BreedEntry": BREED_ENTRY,
    "BreedDest": BREED_DEST,
    "Stage": ["Planned", "Raw material", "Pre-composting", "In beds", "Harvesting", "Closing", "Closed"],
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
    regs = {
        "Batches": ("batches", [("BatchCode", "code", "text", None), ("BatchName", "name", "text", None), ("Site", "site", "text", None),
                    ("StartDate", "start", "date", None), ("Stage", "status", "text", "Stage"), ("Notes", "notes", "text", None)]),
        "RawMaterial": ("rm", [("Batch", "batch", "text", "Batch"), ("PurchaseDate", "date", "date", None), ("RMLot", "lot", "text", None),
            ("Supplier", "supplier", "text", "Supplier"), ("Material", "material", "text", "Material"), ("VehicleNo", "vehicle", "text", None),
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
        "Earthworm": ("earthworm", [("EntryDate", "date", "date", None), ("Unit", "unit", "text", None), ("Entry", "entry", "text", "BreedEntry"),
            ("Supplier", "supplier", "text", None), ("Species", "species", "text", None), ("WormsKg", "wormsKg", "num", None),
            ("Material", "material", "text", "Material"), ("QtyKg", "qtyKg", "kg", None), ("CostRs", "cost", "inr", None),
            ("MaterialOutKg", "outKg", "kg", None), ("Destination", "dest", "text", "BreedDest"),
            ("ToBatch", "toBatch", "text", "Batch"), ("Notes", "notes", "text", None)]),
        "DailyLog": ("logs", [("Batch", "batch", "text", "Batch"), ("LogDate", "date", "date", None), ("Activity", "activity", "text", "Activity"),
            ("BedsArea", "area", "text", None), ("TempC", "temp", "num", None), ("MoisturePct", "moisture", "num", None), ("Qty", "qty", "num", None),
            ("RecordedBy", "by", "text", None), ("Observations", "notes", "text", None)]),
    }
    add_links(regs)
    return regs


# "Go to" link columns: (table, new column, key columns in this table, target table, target key columns)
LINKS = [
    ("RawMaterial", "GoToPreCompost", ["RMLot"], "PreCompost", ["RMLot"]),
    ("PreCompost", "GoToRMLot", ["RMLot"], "RawMaterial", ["RMLot"]),
    ("PreCompost", "GoToBeds", ["PCLot"], "Beds", ["PCLot"]),
    ("Beds", "GoToPCLot", ["PCLot"], "PreCompost", ["PCLot"]),
    ("Beds", "GoToHarvest", ["Batch", "BedNo"], "Harvest", ["Batch", "BedNo"]),
    ("Beds", "GoToBatch", ["Batch"], "Batches", ["BatchCode"]),
    ("Harvest", "GoToBed", ["Batch", "BedNo"], "Beds", ["Batch", "BedNo"]),
    ("Harvest", "GoToSale", ["FGBatch"], "Sales", ["FGBatch"]),
    ("Sales", "GoToHarvest", ["FGBatch"], "Harvest", ["FGBatch"]),
    ("Sales", "GoToBatch", ["Batch"], "Batches", ["BatchCode"]),
    ("QualityControl", "GoToBatch", ["Batch"], "Batches", ["BatchCode"]),
    ("Earthworm", "GoToBatch", ["ToBatch"], "Batches", ["BatchCode"]),
    ("Batches", "GoToBeds", ["BatchCode"], "Beds", ["Batch"]),
]


def add_links(regs):
    """Append locked formula columns whose cells are clickable links to the related record."""
    letters = {t: {h: get_column_letter(i + 1) for i, (h, *_) in enumerate(spec)} for t, (_, spec) in regs.items()}
    for t, name, keys, target, tkeys in LINKS:
        key = "{" + keys[-1] + "}"
        L = letters[target][tkeys[-1]]
        if len(keys) == 1:
            pos = f"MATCH({key},{target}[{tkeys[0]}],0)"
        else:
            cond = "*".join(f"({target}[{tk}]={{{k}}})" for k, tk in zip(keys, tkeys))
            pos = f"MATCH(1,INDEX({cond},0),0)"
        f = (f'=IF({key}="","",IFERROR(HYPERLINK("#\'{target}\'!{L}"&({pos}+1),"▸ {target}: "&{key}),"– not in {target}"))')
        regs[t][1].append((name, None, "calc_link", f))


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
    return (f'=IF({this(t, "Batch")}="","",IF(OR({p}="FG",{p}="EXFG"),IF({flags}="","Within FCO","Outside: "&TRIM({flags})),"Trial (not checked)"))')


def build(src, out, batch):
    data = convert(src, batch)
    data["logs"] = []
    # B5 is in process: add it to Batches so it can be selected
    data["batches"].append({"code": "B5", "name": "Batch 5", "site": "Akola", "start": "", "status": "In beds",
                            "notes": "In process – add dates and records as they happen"})
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
    lists["Supplier"] = [None] * 40   # filled by formula from RawMaterial[Supplier]
    lists["Customer"] = [None] * 40   # filled by formula from Sales[Customer]
    ws_list["A1"] = "BatchPick"
    ws_list["A2"] = "All"
    for i in range(3, 53):
        ws_list[f"A{i}"] = f'=IFERROR(IF(INDEX(Batches[BatchCode],{i - 2})="","",INDEX(Batches[BatchCode],{i - 2})),"")'
    col = 2
    list_ranges = {"Batch": "Batches!$A$2:$A$200", "BatchPick": "Lists!$A$2:$A$52"}
    for name, vals in lists.items():
        c = ws_list.cell(1, col, name)
        n = len(vals)
        L = c.column_letter
        src = {"Supplier": "RawMaterial!$D$2:$D$5000", "Customer": "Sales!$D$2:$D$5000"}.get(name)
        for j, v in enumerate(vals):
            if src:  # distinct names typed in the register, in order of first use
                v = (f'=IFERROR(INDEX({src},MATCH(0,INDEX(COUNTIF({L}$1:{L}{j + 1},{src})+({src}=""),0),0)),"")')
            ws_list.cell(j + 2, col, v)
        list_ranges[name] = f"Lists!${L}$2:${L}${n + 1}"
        col += 1
    # FCO table at S:T with labels in R
    ws_list["R1"], ws_list["S1"], ws_list["T1"], ws_list["U1"] = "FCO parameter", "Min", "Max", "Column"
    for i, (colname, label, lo, hi) in enumerate(FCO):
        ws_list.cell(i + 2, 18, label)
        ws_list.cell(i + 2, 19, lo)
        ws_list.cell(i + 2, 20, hi)
        ws_list.cell(i + 2, 21, colname)
    ws_list["V1"] = "Material"
    for j, m in enumerate(MATERIALS):
        ws_list.cell(j + 2, 22, m)
    ws_list["V1"].comment = Comment("Material types for the RawMaterial drop-down.", "Navjyoti")
    ws_list["R11"] = "Source: FCO 1985 vermicompost specification (P as P2O5, K as K2O). Edit Min / Max if your buyer's export spec differs."
    for name, ref in list_ranges.items():
        wb.defined_names[f"L_{name}"] = DefinedName(f"L_{name}", attr_text=ref)
    for c in ws_list[1]:
        c.font = F(bold=True, color="FFFFFF")
        c.fill = HDR_CALC
    ws_list.column_dimensions["A"].width = 12
    for L in "BCDEFGHIJKLMNOPQVW":
        ws_list.column_dimensions[L].width = 20
    ws_list.column_dimensions["R"].width = 18
    ws_list["A1"].comment = Comment("Filled automatically from the Batches sheet. Used by the batch selectors.", "Navjyoti")
    ws_list["P1"].comment = Comment("Filled automatically from the Supplier column of RawMaterial.", "Navjyoti")
    ws_list["Q1"].comment = Comment("Filled automatically from the Customer column of Sales.", "Navjyoti")
    for r in range(2, 10):
        for c in (19, 20):
            ws_list.cell(r, c).protection = Protection(locked=False)
            ws_list.cell(r, c).fill = FILL("FFF3B0")
    protect(ws_list)

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
        cap = CAPACITY.get(tname, PREFILL)
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
                    if kind == "calc_link":
                        cell.font = F(color="1C5FB8", underline="single")
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
        tab = Table(displayName=tname, ref=f"A1:{last}{cap}")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        if tname == "Batches":  # free-text notes: no filter arrow on the heading
            tab.autoFilter = AutoFilter(ref=f"A1:{last}{cap}", filterColumn=[FilterColumn(colId=[h for h, *_ in cols].index("Notes"), hiddenButton=True, showButton=False)])
        ws.add_table(tab)
        # formula columns are pre-filled below the table, so a row typed under the table
        # already calculates when Excel extends the table over it
        fmts = {"date": DATE_FMT, "kg": KG_FMT, "inr": INR_FMT, "calc_inr": INR_FMT, "calc_kg": KG_FMT, "calc_pct": PCT_FMT, "num": "0.##", "calc_num": "0.#"}
        for ri in range(nrows + 2, cap + 1):
            for ci, (h, _, kind, extra) in enumerate(cols, 1):
                c = ws.cell(ri, ci)
                c.font = F()
                c.number_format = fmts.get(kind, "General")
                if kind.startswith("calc"):
                    f = fco_formula(tname) if extra == "FCO" else extra
                    for hh, *_ in cols:
                        f = f.replace("{" + hh + "}", this(tname, hh))
                    c.value = f
                    if kind == "calc_link":
                        c.font = F(color="1C5FB8", underline="single")
        # typing cells are open; headers and formula columns stay locked
        for ci, (h, _, kind, extra) in enumerate(cols, 1):
            if not kind.startswith("calc"):
                for ri in range(2, cap + 1):
                    ws.cell(ri, ci).protection = Protection(locked=False)
        protect(ws)
        ws.freeze_panes = "C2" if tname != "Batches" else "B2"
        ws.row_dimensions[1].height = 30
        calc_cols[tname] = [h for h, _, kind, _ in cols if kind.startswith("calc")]
        # dropdowns on the input columns (to row 5000 so new rows keep them)
        for ci, (h, _, kind, extra) in enumerate(cols, 1):
            if extra and not kind.startswith("calc") and extra in list_ranges:
                L = ws.cell(1, ci).column_letter
                dv = DataValidation(type="list", formula1=f"L_{extra}", allow_blank=True,
                                    showErrorMessage=extra not in ("Supplier", "Customer"),
                                    errorTitle="Pick from the list", error="Choose a value from the drop-down list.")
                dv.add(f"{L}2:{L}{cap}")
                ws.add_data_validation(dv)
            if kind == "date":
                L = ws.cell(1, ci).column_letter
                dv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                                    error="Enter a date, for example 14-02-2026.", errorTitle="Date needed")
                dv.add(f"{L}2:{L}{cap}")
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

    sheets["Batches"].column_dimensions["F"].width = 50

    # ---------- Dashboard ----------
    ws = ws_dash
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 34
    for L in "CDEFGHIJKL":
        ws.column_dimensions[L].width = 15
    banner(ws, 1, 2, 12, "Navjyoti · Vermicompost Production Dashboard",
           "Live from the register sheets · pick a batch in C5 (or All) · every figure here is a formula")
    ws["B5"] = "Selected batch  ▸"
    ws["B5"].font = F(bold=True, size=11, color=INK)
    ws["B5"].alignment = Alignment(horizontal="right")
    ws["C5"] = "All"
    ws["C5"].fill = FILL("FFF3B0")
    ws["C5"].font = F(bold=True, size=12, color=INK)
    ws["C5"].border = Border(left=Side(style="medium", color="D9A400"), right=Side(style="medium", color="D9A400"),
                             top=Side(style="medium", color="D9A400"), bottom=Side(style="medium", color="D9A400"))
    ws["C5"].alignment = Alignment(horizontal="center")
    dv = DataValidation(type="list", formula1="L_BatchPick", allow_blank=False)
    dv.add("C5")
    ws.add_data_validation(dv)
    ws["D5"] = '=IF(C5="All","All batches combined",IFERROR(INDEX(Batches[BatchName],MATCH(C5,Batches[BatchCode],0))&"  ·  "&INDEX(Batches[Stage],MATCH(C5,Batches[BatchCode],0)),""))'
    ws["D5"].font = F(color=MUTED, italic=True)
    ws["J5"] = "Today"
    ws["J5"].font = F(color=MUTED)
    ws["J5"].alignment = Alignment(horizontal="right")
    ws["K5"] = "=TODAY()"
    ws["K5"].number_format = "dd mmm yyyy"
    ws["K5"].font = F(bold=True)
    wb.defined_names["SelBatch"] = DefinedName("SelBatch", attr_text='Dashboard!$C$5')
    K = 'IF(SelBatch="All","*",SelBatch)'  # SUMIFS criterion: * matches every batch

    # detailed figures (B:C) — the tiles above read from these cells
    R = 18
    kpis = [
        ("rm", "Raw material received (t)", f"=SUMIFS(RawMaterial[QtyKg],RawMaterial[Batch],{K})/1000", "#,##0.0"),
        ("rmcost", "Raw material cost", f"=SUMIFS(RawMaterial[Amount],RawMaterial[Batch],{K})", INR_FMT),
        ("pc", "Pre-compost decomposed (MT)", f"=SUMIFS(PreCompost[WeightMT],PreCompost[Batch],{K})", "#,##0.0"),
        ("beds", "Beds filled", f'=COUNTIFS(Beds[Batch],{K},Beds[BedNo],"<>")', "#,##0"),
        ("harv", "Beds harvested", f'=COUNTIFS(Beds[Batch],{K},Beds[LiveStatus],"Harvested")', "#,##0"),
        ("dung", "Dung / compost filled in beds (t)", f"=SUMIFS(Beds[DungT],Beds[Batch],{K})", "#,##0.0"),
        ("raw", "Raw harvest (kg)", f"=SUMIFS(Harvest[RawKg],Harvest[Batch],{K})", KG_FMT),
        ("net", "Net yield after sieving (kg)", f"=SUMIFS(Harvest[NetKg],Harvest[Batch],{K})", KG_FMT),
        ("conv", "Conversion (net yield ÷ raw material)", '=IFERROR(C{net}/(C{rm}*1000),"")', PCT_FMT),
        ("sold", "Sold (kg)", f"=SUMIFS(Sales[QtyKg],Sales[Batch],{K})", KG_FMT),
        ("rev", "Revenue", f"=SUMIFS(Sales[Revenue],Sales[Batch],{K})", INR_FMT),
        ("price", "Average price (₹/kg)", '=IFERROR(C{rev}/C{sold},"")', "₹#,##0.00"),
        ("stock", "Stock on hand (kg, stock ledger)", f"=SUMIFS(StockLedger[InKg],StockLedger[Batch],{K})-SUMIFS(StockLedger[OutKg],StockLedger[Batch],{K})-SUMIFS(StockLedger[LossKg],StockLedger[Batch],{K})", KG_FMT),
        ("exp", "Expenses", f"=SUMIFS(Expenses[Amount],Expenses[Batch],{K})", INR_FMT),
        ("cpk", "Expense per kg produced (₹)", '=IFERROR(C{exp}/C{net},"")', "₹#,##0.00"),
        ("margin", "Revenue − expenses", "=C{rev}-C{exp}", '₹#,##0;[Red]-₹#,##0'),
    ]
    at = {k: R + i for i, (k, *_) in enumerate(kpis)}
    subhead(ws, R - 2, 2, 3, "DETAILED FIGURES", GREEN)
    for i, (k, label, f, fmt) in enumerate(kpis):
        r = R + i
        band = FILL("F4F7F3") if i % 2 else FILL("FFFFFF")
        a = ws.cell(r, 2, label)
        c = ws.cell(r, 3, f.format(**at))
        a.font, c.font = F(color=INK), F(bold=True, size=11, color=INK)
        c.number_format = fmt
        c.alignment = Alignment(horizontal="right")
        for cc in (a, c):
            cc.fill, cc.border = band, Border(bottom=THIN)

    # KPI tiles: (label, value cell, format, sub-line formula, accent, tint)
    tiles = [
        ("RAW MATERIAL IN", f"=C{at['rm']}", '#,##0.0" t"', f'=TEXT(C{at["rmcost"]},"₹#,##0")&" cost"', "2A78D6", "E6F0FB"),
        ("NET YIELD", f"=C{at['net']}", '#,##0" kg"', f'=IF(C{at["conv"]}="","",TEXT(C{at["conv"]},"0.0%")&" of raw material")', GREEN, "E3F1E8"),
        ("BEDS HARVESTED", f'=C{at["harv"]}&" / "&C{at["beds"]}', "General", f'=COUNTIFS(Beds[Batch],{K},Beds[LiveStatus],"Overdue")&" overdue"', "7A4FB5", "EFE9F8"),
        ("SOLD", f"=C{at['sold']}", '#,##0" kg"', f'=IF(C{at["price"]}="","no sales yet","avg ₹"&TEXT(C{at["price"]},"0.00")&" / kg")', "EB6834", "FDEDE5"),
        ("REVENUE", f"=C{at['rev']}", INR_FMT, f'=TEXT(C{at["exp"]},"₹#,##0")&" expenses"', "138A60", "E1F4EC"),
        ("STOCK ON HAND", f"=C{at['stock']}", '#,##0" kg"', None, "B7791F", "FCF2DD"),
    ]
    tile_cols = [(2, 2), (3, 4), (5, 6), (7, 8), (9, 10), (11, 12)]
    for (label, val, fmt, sub, accent, tint), (c1, c2) in zip(tiles, tile_cols):
        for r in (7, 8, 9, 10):
            for cc in range(c1, c2 + 1):
                ws.cell(r, cc).fill = FILL(tint)
            if c2 > c1:
                ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        for cc in range(c1, c2 + 1):
            ws.cell(7, cc).border = Border(top=Side(style="thick", color=accent))
        ws.cell(7, c1, label).font = F(bold=True, size=9, color=accent)
        v = ws.cell(8, c1, val)
        v.font = F(bold=True, size=18, color=INK)
        v.number_format = fmt
        s = ws.cell(9, c1, sub)
        s.font = F(size=9, color=MUTED)
        for r in (7, 8, 9):
            ws.cell(r, c1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[8].height = 30
    # second strip: finance
    tiles2 = [
        ("EXPENSES", f"=C{at['exp']}", INR_FMT, "D14B3B", "FBE6E3"),
        ("EXPENSE PER KG", f"=C{at['cpk']}", "₹#,##0.00", "D14B3B", "FBE6E3"),
        ("REVENUE − EXPENSES", f"=C{at['margin']}", '₹#,##0;[Red]-₹#,##0', "138A60", "E1F4EC"),
        ("PRE-COMPOST", f"=C{at['pc']}", '#,##0.0" MT"', "7A4FB5", "EFE9F8"),
        ("DUNG IN BEDS", f"=C{at['dung']}", '#,##0.0" t"', GREEN, "E3F1E8"),
        ("RAW HARVEST", f"=C{at['raw']}", '#,##0" kg"', "2A78D6", "E6F0FB"),
    ]
    for (label, val, fmt, accent, tint), (c1, c2) in zip(tiles2, tile_cols):
        for r in (12, 13):
            for cc in range(c1, c2 + 1):
                ws.cell(r, cc).fill = FILL(tint)
            if c2 > c1:
                ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        for cc in range(c1, c2 + 1):
            ws.cell(12, cc).border = Border(top=Side(style="medium", color=accent))
        ws.cell(12, c1, label).font = F(bold=True, size=9, color=accent)
        v = ws.cell(13, c1, val)
        v.font = F(bold=True, size=13, color=INK)
        v.number_format = fmt
        for r in (12, 13):
            ws.cell(r, c1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[13].height = 22

    # earthworm purchases and breeding (kept outside every batch figure)
    B0 = R - 2
    subhead(ws, B0, 5, 9, "EARTHWORM  ·  purchases & breeding (not counted in any batch)", "7A4FB5")
    breed = [
        ("Earthworms purchased (kg)", '=SUMIFS(Earthworm[WormsKg],Earthworm[Entry],"Earthworm purchase")', KG_FMT),
        ("Earthworm purchase cost (₹)", '=SUMIFS(Earthworm[CostRs],Earthworm[Entry],"Earthworm purchase")', INR_FMT),
        ("Dung / feed added for breeding (kg)", '=SUMIFS(Earthworm[QtyKg],Earthworm[Entry],"Dung / feed added")', KG_FMT),
        ("Breeding cost (₹)", '=SUMIFS(Earthworm[CostRs],Earthworm[Entry],"<>Earthworm purchase")', INR_FMT),
        ("Earthworms produced by breeding (kg)", '=SUMIFS(Earthworm[WormsKg],Earthworm[Entry],"Worms harvested")', KG_FMT),
        ("Earthworms issued to beds (kg)", '=SUMIFS(Earthworm[WormsKg],Earthworm[Entry],"Worms issued to beds")', KG_FMT),
        ("Material sent for sale after mixing (kg)", f'=SUMIFS(Earthworm[MaterialOutKg],Earthworm[Destination],"{BREED_DEST[0]}")', KG_FMT),
        ("Total earthworm spend (₹)", "=SUM(Earthworm[CostRs])", INR_FMT),
    ]
    for i, (label, f, fmt) in enumerate(breed):
        r = B0 + 1 + i
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
        ws.cell(r, 5, label).font = F(color=INK)
        c = ws.cell(r, 8, f)
        c.font, c.number_format, c.alignment = F(bold=True, color=INK), fmt, Alignment(horizontal="right")
        for cc in range(5, 10):
            ws.cell(r, cc).border = Border(bottom=THIN)
            ws.cell(r, cc).fill = FILL("F3EEFA")

    # charts
    T = 37
    subhead(ws, T, 2, 12, "TRENDS", GREEN)
    r0 = 56          # batch comparison table
    BC = 10          # batch rows (new batches fill in automatically)
    G0 = r0 + BC + 4  # batch comparison chart
    D = G0 + 20      # chart data
    subhead(ws, D - 2, 2, 12, "CHART DATA (calculated · feeds the charts above)", "9AA79F")
    ws.cell(D - 1, 2, "Monthly chart starts from (blank = automatic)").font = F(color=INK)
    ms = ws.cell(D - 1, 3)
    ms.number_format, ms.fill, ms.border, ms.font = "mmm yyyy", FILL("FFF3B0"), BOX, F(bold=True)
    ms.protection = Protection(locked=False)
    first = first_date("RawMaterial[PurchaseDate]", '(((SelBatch="All")+(RawMaterial[Batch]=SelBatch))>0)')
    first2 = first_date("Harvest[HarvestDate]", '(((SelBatch="All")+(Harvest[Batch]=SelBatch))>0)')
    eff = ws.cell(D - 1, 4, start_month(f"C{D - 1}", first, first2))
    eff.number_format, eff.font = '"from "mmm yyyy', F(color=MUTED, italic=True)
    hdr_row(ws, D, 2, ["Month", "Net yield (kg)", "Sold (kg)"])
    for i in range(12):
        r = D + 1 + i
        ws.cell(r, 2, f"=EDATE($D${D - 1},{i})").number_format = "mmm yy"
        ws.cell(r, 3, f'=SUMIFS(Harvest[NetKg],Harvest[Batch],{K},Harvest[HarvestDate],">="&B{r},Harvest[HarvestDate],"<"&EDATE(B{r},1))').number_format = KG_FMT
        ws.cell(r, 4, f'=SUMIFS(Sales[QtyKg],Sales[Batch],{K},Sales[SaleDate],">="&B{r},Sales[SaleDate],"<"&EDATE(B{r},1))').number_format = KG_FMT
    hdr_row(ws, D, 6, ["Expense category", "Amount (₹)"])
    cats = LISTS["ExpenseCategory"]
    for i, cat in enumerate(cats):
        r = D + 1 + i
        ws.cell(r, 6, cat)
        ws.cell(r, 7, f"=SUMIFS(Expenses[Amount],Expenses[Batch],{K},Expenses[Category],F{r})").number_format = INR_FMT
    for row in ws.iter_rows(min_row=D + 1, max_row=D + 12, min_col=2, max_col=7):
        for c in row:
            c.font = F(color=INK)

    ch = BarChart()
    ch.type = "col"
    ch.add_data(Reference(ws, min_col=3, max_col=4, min_row=D, max_row=D + 12), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=D + 1, max_row=D + 12))
    style_chart(ch, 2, colors=[GREEN, "EB6834"])
    ch.y_axis.numFmt = KG_AXIS
    ch.height, ch.width = 7.5, 15.5
    ws.cell(T + 1, 2, "Net yield vs sold per month (kg)").font = F(bold=True, color=INK)
    ws.add_chart(ch, f"B{T + 2}")
    ch2 = BarChart()
    ch2.type = "bar"
    ch2.add_data(Reference(ws, min_col=7, min_row=D, max_row=D + len(cats)), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=6, min_row=D + 1, max_row=D + len(cats)))
    style_chart(ch2, 1, colors=["D14B3B"], reverse=True)
    ch2.legend = None
    ch2.y_axis.numFmt = INR_AXIS
    ch2.height, ch2.width = 7.5, 20.5
    ws.cell(T + 1, 6, "Where the money went (₹)").font = F(bold=True, color=INK)
    ws.add_chart(ch2, f"F{T + 2}")

    # batch comparison (every batch in the Batches sheet appears here by itself)
    subhead(ws, r0, 2, 11, "BATCH COMPARISON  (new batches appear automatically)", "2A78D6")
    heads = ["Batch", "Stage", "Beds harvested / filled", "RM received (t)", "Net yield (kg)", "Conversion (net ÷ RM)", "Sold (kg)", "Revenue", "Expenses", "Revenue − expenses"]
    hdr_row(ws, r0 + 1, 2, heads, fill="2A78D6")
    ws.row_dimensions[r0 + 1].height = 30
    for i in range(BC):
        r = r0 + 2 + i
        b = f"B{r}"
        ws[b] = (f'=IF(Lists!A{3 + i}="","",IFERROR(HYPERLINK("#\'Batches\'!A"&(MATCH(Lists!A{3 + i},Batches[BatchCode],0)+1),Lists!A{3 + i}),Lists!A{3 + i}))')
        vals = [f'=IF({b}="","",IFERROR(INDEX(Batches[Stage],MATCH({b},Batches[BatchCode],0)),""))',
                f'=IF({b}="","",COUNTIFS(Beds[Batch],{b},Beds[LiveStatus],"Harvested")&" / "&COUNTIFS(Beds[Batch],{b},Beds[BedNo],"<>"))',
                f'=IF({b}="","",SUMIFS(RawMaterial[QtyKg],RawMaterial[Batch],{b})/1000)',
                f'=IF({b}="","",SUMIFS(Harvest[NetKg],Harvest[Batch],{b}))',
                f'=IF({b}="","",IFERROR(F{r}/(E{r}*1000),""))',
                f'=IF({b}="","",SUMIFS(Sales[QtyKg],Sales[Batch],{b}))',
                f'=IF({b}="","",SUMIFS(Sales[Revenue],Sales[Batch],{b}))',
                f'=IF({b}="","",SUMIFS(Expenses[Amount],Expenses[Batch],{b}))',
                f'=IF({b}="","",I{r}-J{r})']
        fmts = ["General", "General", "#,##0.0", KG_FMT, PCT_FMT, KG_FMT, INR_FMT, INR_FMT, '₹#,##0;[Red]-₹#,##0']
        band = FILL("EEF4FC") if i % 2 == 0 else FILL("FFFFFF")
        ws[b].font = F(bold=True, color="1C5FB8", underline="single")
        ws[b].fill = band
        for j, (f, fm) in enumerate(zip(vals, fmts)):
            c = ws.cell(r, 3 + j, f)
            c.font, c.number_format, c.fill, c.border = F(color=INK), fm, band, Border(bottom=THIN)
    ws.cell(G0, 2, "Net yield and sold by batch (kg)").font = F(bold=True, color=INK)
    ch4 = BarChart()
    ch4.type = "col"
    for col, title in ((6, "Net yield (kg)"), (8, "Sold (kg)")):
        ch4.series.append(Series(Reference(ws, min_col=col, min_row=r0 + 2, max_row=r0 + 1 + BC), title=title))
    ch4.set_categories(Reference(ws, min_col=2, min_row=r0 + 2, max_row=r0 + 1 + BC))
    style_chart(ch4, 2, colors=[GREEN, "EB6834"])
    ch4.y_axis.numFmt = KG_AXIS
    ch4.height, ch4.width = 7.5, 36
    ws.add_chart(ch4, f"B{G0 + 1}")
    ws.freeze_panes = "A6"
    ws["C5"].protection = Protection(locked=False)
    protect(ws)

    build_reports(wb, ws_rep, list_ranges, lists["Supplier"], lists["Customer"])
    build_readme(ws_readme)
    protect(ws_readme)
    for w in wb.worksheets:
        w.sheet_properties.tabColor = {"How to use": SOIL, "Dashboard": GREEN, "Reports": GREEN, "Lists": "9AA79F"}.get(w.title, "C9D6CB")
    for w in wb.worksheets:
        w.page_setup.orientation = "landscape"
        w.page_setup.paperSize = w.PAPERSIZE_A4
        w.sheet_properties.pageSetUpPr.fitToPage = True
        w.page_setup.fitToWidth, w.page_setup.fitToHeight = 1, 0
    to_a1(wb, regs)
    wb.calculation.fullCalcOnLoad = True
    wb.active = 1
    wb.save(out)


LASTWATER = ('SUMPRODUCT(MAX(DailyLog[LogDate]*(DailyLog[Activity]="Watering")'
             '*(((SelBatch="All")+(DailyLog[Batch]=SelBatch))>0)))')


def to_a1(wb, regs):
    """Rewrite table references (Harvest[NetKg], Beds[[#This Row],[BedNo]]) as plain A1 ranges,
    which every Excel version (2010 onwards, web, phone) reads without #REF!."""
    cols = {}
    for tname, (_, spec) in regs.items():
        ws = wb[tname]
        cols[tname] = {h: ws.cell(1, i + 1).column_letter for i, (h, *_) in enumerate(spec)}
    this_re = re.compile(r"(\w+)\[\[#This Row\],\[([^\]]+)\]\]")
    col_re = re.compile(r"(\w+)\[([A-Za-z0-9]+)\]")
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith("=") and "[" in v):
                    continue
                v = this_re.sub(lambda m: f"{cols[m.group(1)][m.group(2)]}{c.row}", v)
                v = col_re.sub(lambda m: f"{m.group(1)}!${cols[m.group(1)][m.group(2)]}$2:${cols[m.group(1)][m.group(2)]}${LAST_ROW}"
                               if m.group(1) in cols else m.group(0), v)
                assert "[" not in v.replace('"[', ''), v
                c.value = v


def style_chart(ch, n, colors=None, reverse=False):
    colors = colors or SERIES
    for i, s in enumerate(ch.series[:n]):
        s.graphicalProperties = GraphicalProperties(solidFill=colors[i % len(colors)])
        s.graphicalProperties.line.solidFill = colors[i % len(colors)]
    if ch.legend is not None:
        ch.legend.position = "t"      # above the plot, clear of the axis numbers
        ch.legend.overlay = False
    ch.gapWidth = 50
    ch.overlap = -10 if n > 1 else 0
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    ch.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill="E3E8E4")))
    ch.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.x_axis.spPr = GraphicalProperties(ln=LineProperties(solidFill="9AA79F"))
    ch.y_axis.txPr = axis_text()
    ch.x_axis.txPr = axis_text()
    if reverse:
        # horizontal bars top-to-bottom in list order, value axis kept at the bottom
        ch.x_axis.scaling.orientation = "maxMin"
        ch.y_axis.crosses = "max"


def axis_text(size=900, color="3E4A43"):
    return RichText(bodyPr=RichTextProperties(), p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=size, solidFill=color)), endParaRPr=CharacterProperties(sz=size))])


def banner(ws, row, c1, c2, title, sub):
    for r in (row, row + 1, row + 2):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).fill = FILL(GREEN if r < row + 2 else "2E8A5F")
    t = ws.cell(row, c1, title)
    t.font = F(bold=True, size=20, color="FFFFFF")
    t.alignment = Alignment(vertical="center", indent=1)
    s = ws.cell(row + 2, c1, sub)
    s.font = F(size=10, color="E3F1E8", italic=True)
    s.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 26
    ws.row_dimensions[row + 1].height = 8
    ws.row_dimensions[row + 2].height = 18


def subhead(ws, row, c1, c2, text, color):
    for c in range(c1, c2 + 1):
        ws.cell(row, c).border = Border(bottom=Side(style="medium", color=color))
    ws.cell(row, c1, text).font = F(bold=True, size=11, color=color)


def hdr_row(ws, row, c1, heads, fill=GREEN):
    for j, h in enumerate(heads):
        c = ws.cell(row, c1 + j, h)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = FILL(fill)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left" if j == 0 else "center")


def build_reports(wb, ws, list_ranges, lists_sup, lists_cus):
    ws.sheet_view.showGridLines = False
    for L, w in zip("ABCDEFGH", (2, 30, 15, 15, 15, 15, 3, 3)):
        ws.column_dimensions[L].width = w
    for L in "IJKLMNOP":
        ws.column_dimensions[L].width = 11
    banner(ws, 1, 2, 16, "Reports · Compare Batches",
           "Pick up to three batches in C5:E5 · every table and chart updates live · click a report name to jump to it")
    ws["B5"], ws["B6"], ws["B7"] = "Compare batches  ▸", "Months start from  ▸", "Quality report batch  ▸"
    ws["F6"] = "leave blank = starts from the first purchase of the chosen batches"
    ws["F6"].font = F(color=MUTED, italic=True, size=9)
    for c in (ws["B5"], ws["B6"], ws["B7"]):
        c.font = F(bold=True, color=INK)
    ws["C5"], ws["D5"], ws["E5"] = "B4", "B5", None
    ws["C6"].number_format = "mmm yyyy"
    rmsel = ('((RawMaterial[Batch]=$C$5)+(RawMaterial[Batch]=$D$5)+(RawMaterial[Batch]=$E$5)>0)')
    hvsel = ('((Harvest[Batch]=$C$5)+(Harvest[Batch]=$D$5)+(Harvest[Batch]=$E$5)>0)')
    ws["D6"] = start_month("C6", first_date("RawMaterial[PurchaseDate]", rmsel), first_date("Harvest[HarvestDate]", hvsel))
    ws["D6"].number_format = '"→ "mmm yyyy'
    ws["D6"].font = F(color=MUTED, italic=True)
    ws["C7"] = "All"
    for a in ("C5", "D5", "E5", "C6", "C7"):
        ws[a].protection = Protection(locked=False)
        ws[a].fill = FILL("FFF3B0")
        ws[a].border = BOX
        ws[a].font = F(bold=True, color=INK)
        ws[a].alignment = Alignment(horizontal="center")
    for a, colr in zip(("C4", "D4", "E4"), SERIES):
        ws[a].fill = FILL(colr)
    ws["F5"] = "← colour = bar colour in charts"
    ws["F5"].font = F(color=MUTED, italic=True, size=9)
    dv = DataValidation(type="list", formula1="L_Batch", allow_blank=True)
    dv.add("C5:E5")
    ws.add_data_validation(dv)
    dv2 = DataValidation(type="list", formula1="L_BatchPick", allow_blank=False)
    dv2.add("C7")
    ws.add_data_validation(dv2)

    sections = []
    row = 14

    def table(title, short, cat_label, cats, fn, fmt, note=None, chart_title=None, total=None):
        """cats: list of (label or formula, criteria builder). fn(batch_cell, cat_row) -> formula"""
        nonlocal row
        start = row
        for cc in range(2, 17):
            ws.cell(row, cc).fill = FILL("E3F1E8")
        ws.cell(row, 2, title).font = F(bold=True, size=12, color=GREEN)
        ws.cell(row, 2).alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[row].height = 22
        sections.append((short, row))
        top_button(ws, row, 6)
        if note:
            ws.cell(row + 1, 2, note).font = F(color=MUTED, italic=True)
        hr = row + 2
        heads = [cat_label, "=IF($C$5=\"\",\"(pick)\",$C$5)", "=IF($D$5=\"\",\"—\",$D$5)", "=IF($E$5=\"\",\"—\",$E$5)", "Total"]
        for j, h in enumerate(heads):
            c = ws.cell(hr, 2 + j, h)
            c.font = F(bold=True, color="FFFFFF")
            c.fill = FILL([INK, *SERIES, INK][j])
            c.alignment = Alignment(horizontal="left" if j == 0 else "right")
        for i, cat in enumerate(cats):
            r = hr + 1 + i
            ws.cell(r, 2, cat).font = F(color=INK)
            ws.cell(r, 2).alignment = Alignment(horizontal="left")
            if isinstance(cat, str) and cat.startswith("=EDATE"):
                ws.cell(r, 2).number_format = "mmm yy"
            for j, bc in enumerate(("$C$5", "$D$5", "$E$5")):
                c = ws.cell(r, 3 + j, f'=IF({bc}="","",{fn(bc, r)})')
                c.number_format = fmt
                c.font = F(color=INK)
            t = ws.cell(r, 6, f"=SUM(C{r}:E{r})")
            t.number_format = fmt
            t.font = F(bold=True, color=INK)
            band = FILL("F4F7F3") if i % 2 else FILL("FFFFFF")
            for j in range(5):
                ws.cell(r, 2 + j).border = Border(bottom=THIN)
                ws.cell(r, 2 + j).fill = band
        end = hr + len(cats)
        tr = end + 1
        ws.cell(tr, 2, total[0] if total else "Total").font = F(bold=True)
        for j in range(4):
            L = "CDEF"[j]
            c = ws.cell(tr, 3 + j, total[1].format(L=L, a=hr + 1, b=hr + 2, c=hr + 3) if total else f"=SUM({L}{hr + 1}:{L}{end})")
            c.number_format = fmt
            c.font = F(bold=True)
            c.fill = FILL("D5E8DC")
            c.border = Border(top=Side(style="thin", color=GREEN))
        ws.cell(tr, 2).fill = FILL("D5E8DC")
        ws.cell(tr, 2).border = Border(top=Side(style="thin", color=GREEN))
        ch = BarChart()
        ch.type = "bar"
        ch.add_data(Reference(ws, min_col=3, max_col=5, min_row=hr, max_row=end), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=2, min_row=hr + 1, max_row=end))
        ch.y_axis.numFmt = INR_AXIS if fmt == INR_FMT else KG_AXIS
        style_chart(ch, 3, reverse=True)
        ch.height = max(6.5, 0.5 * len(cats) + 2.5)
        ch.width = 16
        ws.add_chart(ch, f"I{start + 1}")
        row = max(tr + 3, start + int(ch.height * 2) + 4)
        return hr, end

    months = [f"=EDATE($D$6,{i})" for i in range(18)]
    mrange = lambda tb, dc, r: f'{tb}[{dc}],">="&$B${r},{tb}[{dc}],"<"&EDATE($B${r},1)'
    sup = list_ranges["Supplier"]
    nsup = int(sup.split("$")[-1]) - 1
    table("Raw material – quantity by supplier (kg)", "RM by supplier", "Supplier", [f"=IF(INDEX(L_Supplier,{i + 1})=\"\",\"\",INDEX(L_Supplier,{i + 1}))" for i in range(10)],
          lambda b, r: f"SUMIFS(RawMaterial[QtyKg],RawMaterial[Batch],{b},RawMaterial[Supplier],$B${r})", KG_FMT,
          "Suppliers are picked up automatically from the RawMaterial register (first 10).")
    table("Raw material – quantity by purchase month (kg)", "RM by month", "Month", months,
          lambda b, r: f"SUMIFS(RawMaterial[QtyKg],RawMaterial[Batch],{b},{mrange('RawMaterial', 'PurchaseDate', r)})", KG_FMT)
    table("Pre-compost – weight decomposed by start month (MT)", "Pre-compost", "Month", months,
          lambda b, r: f"SUMIFS(PreCompost[WeightMT],PreCompost[Batch],{b},{mrange('PreCompost', 'StartDate', r)})", "#,##0.0")
    blocks = [f"BED-{s:02d}–{s + 9:02d}" for s in range(1, 100, 10)]
    table("Beds – net yield by bed block (kg)", "Beds by block", "Bed block", blocks,
          lambda b, r: f'SUMIFS(Beds[NetYieldKg],Beds[Batch],{b},Beds[BedNum],">="&VALUE(MID($B${r},5,2)),Beds[BedNum],"<="&VALUE(MID($B${r},5,2))+9)', KG_FMT)
    table("Harvest – net yield by month (kg)", "Harvest by month", "Month", months,
          lambda b, r: f"SUMIFS(Harvest[NetKg],Harvest[Batch],{b},{mrange('Harvest', 'HarvestDate', r)})", KG_FMT)
    table("Sales – quantity by customer (kg)", "Sales by customer", "Customer", [f"=IF(INDEX(L_Customer,{i + 1})=\"\",\"\",INDEX(L_Customer,{i + 1}))" for i in range(10)],
          lambda b, r: f"SUMIFS(Sales[QtyKg],Sales[Batch],{b},Sales[Customer],$B${r})", KG_FMT, "Customers are picked up automatically from the Sales register (first 10).")
    table("Sales – revenue by month (₹)", "Revenue by month", "Month", months,
          lambda b, r: f"SUMIFS(Sales[Revenue],Sales[Batch],{b},{mrange('Sales', 'SaleDate', r)})", INR_FMT)
    table("Stock – movement (kg)", "Stock", "Movement", ["In (production)", "Out (sales)", "Loss / adjustment"],
          lambda b, r: f'SUMIFS(CHOOSE(MATCH($B${r},{{"In (production)","Out (sales)","Loss / adjustment"}},0),StockLedger[InKg],StockLedger[OutKg],StockLedger[LossKg]),StockLedger[Batch],{b})', KG_FMT,
          total=("Closing stock", "=N({L}{a})-N({L}{b})-N({L}{c})"))
    table("Expenses by category (₹)", "Expenses", "Category", [f"=INDEX(L_ExpenseCategory,{i + 1})" for i in range(len(LISTS['ExpenseCategory']))],
          lambda b, r: f"SUMIFS(Expenses[Amount],Expenses[Batch],{b},Expenses[Category],$B${r})", INR_FMT)
    table("Daily log – entries by activity", "Daily log", "Activity", [f"=INDEX(L_Activity,{i + 1})" for i in range(len(LISTS['Activity']))],
          lambda b, r: f"COUNTIFS(DailyLog[Batch],{b},DailyLog[Activity],$B${r})", "#,##0")

    # Quality: RM / FG / EX / EXRM / EXFG
    start = row
    for cc in range(2, 17):
        ws.cell(row, cc).fill = FILL("E3F1E8")
    ws.cell(row, 2, "Quality – RM · FG · EXRM · EXFG (average of lab reports)").font = F(bold=True, size=12, color=GREEN)
    ws.cell(row, 2).alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 22
    sections.append(("Quality comparison", row))
    top_button(ws, row, 9)
    ws.cell(row + 1, 2, "Batch chosen in C7 (All = every batch). FCO limits apply to FG and EXFG (finished material); "
                        "RM and EXRM are trials. Red = outside the limit.").font = F(color=MUTED, italic=True)
    hr = row + 2
    heads = ["Parameter", *[p[0] for p in PRODUCTS], "FCO min", "FCO max"]  # same names as the QualityControl sheet
    for j, h in enumerate(heads):
        c = ws.cell(hr, 2 + j, h)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = FILL([INK, *[p[2] for p in PRODUCTS], "5B6B62", "5B6B62"][j])
        c.alignment = Alignment(horizontal="left" if j == 0 else "right", wrap_text=True, vertical="center")
    ws.row_dimensions[hr].height = 30
    params = [(c, c) for c in ("pH", "EC", "Moisture", "OrganicCarbon", "OrganicMatter", "CNRatio", "Nitrogen", "Phosphorus",
                               "Potassium", "Ash", "Calcium", "Manganese", "Iron", "Manganese2", "Zinc", "Copper")]  # QualityControl order
    crit = 'IF($C$7="All","*",$C$7)'
    fco_idx = {c: i for i, (c, *_) in enumerate(FCO)}
    np_ = len(PRODUCTS)
    cmin, cmax = 3 + np_, 4 + np_
    for i, (colname, label) in enumerate(params):
        r = hr + 1 + i
        ws.cell(r, 2, label).font = F(color=INK)
        for j, (p, *_x) in enumerate(PRODUCTS):
            c = ws.cell(r, 3 + j, f'=IFERROR(AVERAGEIFS(QualityControl[{colname}],QualityControl[Batch],{crit},QualityControl[Product],"{p}"),"")')
            c.number_format = "0.00"
            c.font = F()
        if colname in fco_idx:
            k = fco_idx[colname] + 2
            ws.cell(r, cmin, f'=IF(Lists!$S${k}="","",Lists!$S${k})').font = F(color=MUTED)
            ws.cell(r, cmax, f'=IF(Lists!$T${k}="","",Lists!$T${k})').font = F(color=MUTED)
        for j in range(np_ + 3):
            ws.cell(r, 2 + j).border = Border(bottom=THIN)
            ws.cell(r, 2 + j).fill = FILL("F4F7F3") if i % 2 else FILL("FFFFFF")
    qend = hr + len(params)
    rr = qend + 1
    ws.cell(rr, 2, "Reports included").font = F(color=MUTED)
    for j, (p, *_x) in enumerate(PRODUCTS):
        ws.cell(rr, 3 + j, f'=COUNTIFS(QualityControl[Batch],{crit},QualityControl[Product],"{p}")').font = F(color=MUTED)
    Lmin, Lmax = ws.cell(1, cmin).column_letter, ws.cell(1, cmax).column_letter
    for j, (p, *_x) in enumerate(PRODUCTS):
        if p in FINISHED:
            L = ws.cell(1, 3 + j).column_letter
            ws.conditional_formatting.add(f"{L}{hr + 1}:{L}{qend}", FormulaRule(
                formula=[f'AND(ISNUMBER({L}{hr + 1}),OR(AND(ISNUMBER(${Lmin}{hr + 1}),{L}{hr + 1}<${Lmin}{hr + 1}),AND(ISNUMBER(${Lmax}{hr + 1}),{L}{hr + 1}>${Lmax}{hr + 1})))'],
                fill=BAD_FILL, font=F(bold=True, color="B3261E")))
    for L in "CDEFGHI":
        ws.column_dimensions[L].width = 15
    for title, a, b, anchor in (("pH, EC", hr + 1, hr + 2, start + 1),
                                ("Moisture, OrganicCarbon, OrganicMatter, CNRatio", hr + 3, hr + 6, start + 18),
                                ("Nitrogen, Phosphorus, Potassium", hr + 7, hr + 9, start + 35)):
        ws.cell(anchor, 11, title).font = F(bold=True, color=INK)
        ch = BarChart()
        ch.type = "col"
        for j, (p, lab, colr) in enumerate(PRODUCTS):
            ch.series.append(Series(Reference(ws, min_col=3 + j, min_row=a, max_row=b), title=p))
        ch.set_categories(Reference(ws, min_col=2, min_row=a, max_row=b))
        style_chart(ch, np_, colors=[p[2] for p in PRODUCTS])
        ch.height, ch.width = 7.5, 17
        ws.add_chart(ch, f"K{anchor + 1}")
    row = max(rr + 3, start + 52)

    # Latest FG report compared with the latest RM report
    start = row
    for cc in range(2, 17):
        ws.cell(row, cc).fill = FILL("E3F1E8")
    ws.cell(row, 2, "Quality – latest FG result vs latest RM result").font = F(bold=True, size=12, color=GREEN)
    ws.cell(row, 2).alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 22
    sections.append(("Latest FG vs RM", row))
    top_button(ws, row, 9)
    ws.cell(row + 1, 2, "Most recent lab report of each type for the batch chosen in C7 (not an average). "
                        "Change = FG − RM. Status checks the FG result against the FCO limits.").font = F(color=MUTED, italic=True)
    bc = '((($C$7="All")+(QualityControl[Batch]=$C$7))>0)*(QualityControl[Batch]<>"")'
    info = row + 2
    for k, (p, lab) in enumerate((("RM", "Latest RM report"), ("FG", "Latest FG report"))):
        r = info + k
        ws.cell(r, 2, lab).font = F(bold=True, color=INK)
        dcell = ws.cell(r, 3, f'=SUMPRODUCT(MAX({bc}*(QualityControl[Product]="{p}")*QualityControl[SamplingDate]))')
        dcell.number_format = 'dd-mm-yyyy;;"none yet"'
        dcell.font = F(bold=True, color=INK)
        ws.cell(r, 4, f'=IF(C{r}=0,"",IFERROR(LOOKUP(2,1/({bc}*(QualityControl[Product]="{p}")*(QualityControl[SamplingDate]=C{r})),QualityControl[ReportNo]),""))').font = F(color=MUTED)
        ws.cell(r, 2).fill = ws.cell(r, 3).fill = ws.cell(r, 4).fill = FILL("F4F7F3")
    rm_d, fg_d = f"$C${info}", f"$C${info + 1}"
    hr = info + 3
    heads = ["Parameter", "Latest RM", "Latest FG", "Change (FG − RM)", "FCO min", "FCO max", "FG status"]
    for j, h in enumerate(heads):
        c = ws.cell(hr, 2 + j, h)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = FILL([INK, PRODUCTS[0][2], PRODUCTS[1][2], "5B6B62", "5B6B62", "5B6B62", INK][j])
        c.alignment = Alignment(horizontal="left" if j == 0 else "right", wrap_text=True, vertical="center")
    ws.row_dimensions[hr].height = 30
    lcrit = 'IF($C$7="All","*",$C$7)'
    for i, (colname, _l) in enumerate(params):
        r = hr + 1 + i
        ws.cell(r, 2, colname).font = F(color=INK)
        for j, (p, dref) in enumerate((("RM", rm_d), ("FG", fg_d))):
            c = ws.cell(r, 3 + j, f'=IF({dref}=0,"",IFERROR(AVERAGEIFS(QualityControl[{colname}],QualityControl[Batch],{lcrit},'
                                  f'QualityControl[Product],"{p}",QualityControl[SamplingDate],{dref}),""))')
            c.number_format, c.font = "0.00", F(color=INK)
        ch_ = ws.cell(r, 5, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),D{r}-C{r},"")')
        ch_.number_format, ch_.font = '+0.00;-0.00;0.00', F(bold=True, color=INK)
        if colname in fco_idx:
            k = fco_idx[colname] + 2
            ws.cell(r, 6, f'=IF(Lists!$S${k}="","",Lists!$S${k})').font = F(color=MUTED)
            ws.cell(r, 7, f'=IF(Lists!$T${k}="","",Lists!$T${k})').font = F(color=MUTED)
            st = ws.cell(r, 8, f'=IF(NOT(ISNUMBER(D{r})),"",IF(OR(AND(ISNUMBER(F{r}),D{r}<F{r}),AND(ISNUMBER(G{r}),D{r}>G{r})),"✗ Outside FCO","✓ Within FCO"))')
            st.font = F(bold=True)
            st.alignment = Alignment(horizontal="right")
        for j in range(7):
            ws.cell(r, 2 + j).border = Border(bottom=THIN)
            if j != 6:
                ws.cell(r, 2 + j).fill = FILL("F4F7F3") if i % 2 else FILL("FFFFFF")
    lend = hr + len(params)
    ws.conditional_formatting.add(f"H{hr + 1}:H{lend}", CellIsRule(operator="equal", formula=['"✗ Outside FCO"'], fill=BAD_FILL, font=F(bold=True, color="B3261E")))
    ws.conditional_formatting.add(f"H{hr + 1}:H{lend}", CellIsRule(operator="equal", formula=['"✓ Within FCO"'], fill=OK_FILL, font=F(bold=True, color="1D7443")))
    idx = {c: i for i, (c, _l) in enumerate(params)}
    for title, first_p, last_p, anchor in (("Moisture, OrganicCarbon, OrganicMatter, CNRatio", "Moisture", "CNRatio", start + 1),
                                           ("Nitrogen, Phosphorus, Potassium", "Nitrogen", "Potassium", start + 18)):
        a, b = hr + 1 + idx[first_p], hr + 1 + idx[last_p]
        ws.cell(anchor, 11, title + " · latest RM vs FG").font = F(bold=True, color=INK)
        ch = BarChart()
        ch.type = "col"
        for j, (p, *_x) in enumerate(PRODUCTS[:2]):
            ch.series.append(Series(Reference(ws, min_col=3 + j, min_row=a, max_row=b), title=p))
        ch.set_categories(Reference(ws, min_col=2, min_row=a, max_row=b))
        style_chart(ch, 2, colors=[PRODUCTS[0][2], PRODUCTS[1][2]])
        ch.height, ch.width = 7.5, 17
        ws.add_chart(ch, f"K{anchor + 1}")
    row = max(lend + 3, start + 36)

    # index with links
    ws["B9"] = "Jump to  ▸"
    ws["B9"].font = F(bold=True, color=INK)
    col = 3
    r = 9
    for title, rr_ in sections:
        c = ws.cell(r, col, title)
        c.hyperlink = Hyperlink(ref=c.coordinate, location=f"Reports!B{rr_}", display=c.value)
        c.font = F(color="1C5FB8", underline="single", bold=True)
        col += 1
        if col > 6:
            col, r = 3, r + 1
    ws.freeze_panes = f"A{TOP_ROW}"
    protect(ws)


def first_date(dates, cond):
    """Earliest date in `dates` where `cond` holds, as a formula that works in every Excel version
    (100000 when there is none). Blank rows are ignored."""
    return f"(100000-SUMPRODUCT(MAX({cond}*({dates}>0)*(100000-{dates}))))"


def start_month(manual, first, first2):
    """Chosen start month, else the month of the first purchase, else of the first harvest, else 11 months ago."""
    return (f'=IF({manual}<>"",DATE(YEAR({manual}),MONTH({manual}),1),IF({first}<100000,DATE(YEAR({first}),MONTH({first}),1),'
            f'IF({first2}<100000,DATE(YEAR({first2}),MONTH({first2}),1),DATE(YEAR(TODAY()),MONTH(TODAY())-11,1))))')


TOP_ROW = 8  # first row below the frozen banner/selectors on Reports


def top_button(ws, row, col):
    """A '▲ TOP' button on a section heading that scrolls back to the top of the sheet.
    It targets the first row under the frozen panes: a link to a frozen cell (A1) would not scroll."""
    c = ws.cell(row, col, "▲ TOP")
    c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{ws.title}'!A{TOP_ROW}", display="▲ TOP")
    c.font = F(bold=True, color="FFFFFF", size=9)
    c.fill = FILL(GREEN)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = Border(left=Side(style="thin", color="FFFFFF"), right=Side(style="thin", color="FFFFFF"))


def protect(ws):
    """Lock the sheet: only cells marked unlocked (typing cells, selectors) can be changed."""
    ws.protection.sheet = True
    ws.protection.password = PASSWORD
    ws.protection.autoFilter = False      # filters still work
    ws.protection.sort = False
    ws.protection.formatColumns = False   # column widths can be adjusted
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False
    ws.protection.objects = True          # charts cannot be moved or deleted
    ws.protection.scenarios = True


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
        ("•", "Type new records in the first empty row of a table. Grey-header columns calculate by themselves."),
        ("•", "Always fill the Batch column (drop-down). Start Batch 5 records with B5."),
        ("•", "Use drop-downs where they appear. A new supplier or customer can simply be typed; it is added to the lists automatically."),
        ("•", "Dates: type as 14-02-2026 or pick from the date picker."),
        ("•", "Dashboard: choose a batch (or All) in cell C5. Reports: choose up to three batches in C5:E5 to compare, and the quality batch in C7."),
        ("•", "Every entry you add in any register (including Daily log and new batches) flows into the Dashboard and Reports automatically. New suppliers and customers appear in the reports by themselves; months follow the batch's first purchase unless you set a start month."),
        ("h2", "Locked cells"),
        ("•", "Headings, formula columns, the Dashboard, Reports and Lists are locked so nobody changes them by mistake. Only green-header columns and yellow selector cells accept typing."),
        ("•", f"Sheet password (owner only): {PASSWORD}. Review → Unprotect Sheet to change a layout, then Protect Sheet again."),
        ("•", "Each register has ready rows for new entries (1,500; Daily log 5,000). Type in the next empty row of the table."),
        ("h2", "Colour legend"),
        ("green", "Green header = you type here."),
        ("grey", "Grey header = formula column. Do not type in it; it calculates by itself."),
        ("yellow", "Yellow cell = a selector you can change (batch, month)."),
        ("h2", "Sheets"),
        ("Dashboard", "Key figures, alerts that need attention, batch comparison, monthly net yield vs sales."),
        ("Reports", "Bar charts and tables by supplier, month, bed block, customer, category and activity, for up to three batches side by side; Quality comparison of RM, FG, EXRM and EXFG against FCO limits, and the latest FG result against the latest RM result. Every section has a ▲ TOP button."),
        ("Batches", "One row per batch (B4, B5, …). Batch drop-downs everywhere read from here."),
        ("RawMaterial", "Excel sheet 1 – purchases, supplier, material type (drop-down), vehicle, qty, rate, moisture, lab acceptance. Amount is calculated."),
        ("PreCompost", "Excel sheet 2 – lots, culture dose, up to three turnings. Peak temperature and days to transfer are calculated."),
        ("Beds", "Excel sheet 3 – bed lifecycle. Net yield (from Harvest), live status (Harvested / Overdue / …) and bed number are calculated."),
        ("Harvest", "Excel sheet 4 – raw and net yield per bed, FG batch, packing. Recovery % is calculated."),
        ("QualityControl", "Excel sheet 5 – lab reports. Product: RM = raw material / trial, FG = finished goods, EXRM = export raw material / trial, EXFG = export finished goods. FCO check is calculated for FG and EXFG."),
        ("Sales", "Excel sheet 6 – invoices. Revenue is calculated."),
        ("StockLedger", "Excel sheet 7 – stock in / out / loss. Closing stock runs per batch in entry-number order."),
        ("Expenses", "Excel sheet 8 – expenses by category."),
        ("Earthworm", "One record sheet for earthworms, kept outside every batch. Entry = Earthworm purchase (supplier, species, kg, cost); Dung / feed added (cow dung used for breeding, kg and cost – not in RawMaterial or Expenses); Worms harvested (kg produced); Worms issued to beds (with ToBatch); Material out (kg, with Destination). When breeding material is mixed and ready, choose 'Stock for sale (after mixing)' and add the same kg as Production in on the StockLedger. The Dashboard shows these totals separately."),
        ("Go to links", "Grey 'GoTo…' columns at the end of the registers are links: click one to jump to and select the related record (RM lot ↔ pre-compost lot ↔ beds ↔ harvest ↔ sales, and each batch)."),
        ("DailyLog", "Watering, turning, temperature and moisture checks, feeding, inspections. Example: B5 · 01-10-2026 · Watering · BED-01 to BED-40 · 400 L · Ramesh."),
        ("Lists", "Drop-down values (suppliers and customers collect themselves from the registers) and the FCO reference limits (yellow cells, editable)."),
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
