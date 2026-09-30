"""Convert a Navjyoti vermicompost Excel register (Batch 4 layout) into JSON
documents for the cloud tracker's database.

Usage: python3 convert_register.py <register.xlsx> <batch-code> [out.json]

Dates: the register mixes typed text (dd-mm-yy, dd/mm/yyyy) with cells Excel
parsed as mm/dd. Parsed cells with day <= 12 are swapped back to dd/mm.
"""
import datetime as dt, json, re, sys
import openpyxl

def d(v):
    if v is None or v == "":
        return ""
    if isinstance(v, dt.datetime):
        if v.day <= 12:
            v = v.replace(month=v.day, day=v.month)
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    m = re.match(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$", s)
    if m:
        dd, mm, yy = map(int, m.groups())
        if yy < 100:
            yy += 2000
        return f"{yy:04d}-{mm:02d}-{dd:02d}"
    return s

def n(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return round(float(v), 4)
    m = re.search(r"-?\d+(\.\d+)?", str(v))
    return float(m.group()) if m else None

def t(v):
    return "" if v is None else str(v).strip()

def rows(ws, start):
    for r in ws.iter_rows(min_row=start, values_only=True):
        if any(c is not None for c in r):
            yield r

def expense_category(desc):
    s = desc.lower()
    for key, cat in [("sampl", "Lab testing"), ("smc cost", "Raw material"), ("dung", "Raw material"),
                     ("culcher", "Culture & inputs"), ("culture", "Culture & inputs"), ("dap", "Culture & inputs"),
                     ("jaggery", "Culture & inputs"), ("siev", "Sieving"), ("transport", "Transport"),
                     ("truck", "Transport"), ("jcb", "Machinery"), ("labour", "Labour"), ("bed", "Bed setup"),
                     ("bamboo", "Bed setup"), ("gunny", "Packaging"), ("oder", "Sales & dispatch")]:
        if key in s:
            return cat
    return "Other"

def convert(path, batch):
    wb = openpyxl.load_workbook(path)
    b = batch.lower()
    out = {"batches": [], "rm": [], "precompost": [], "beds": [], "harvest": [], "qc": [],
           "sales": [], "stock": [], "expenses": []}

    for i, r in enumerate(r for r in rows(wb["1. Raw Material Purchase"], 5) if r[1]):
        out["rm"].append({"_id": f"{b}-rm-{i+1:03d}", "batch": batch, "date": d(r[0]), "lot": t(r[1]),
            "supplier": t(r[2]), "material": t(r[3]), "vehicle": t(r[4]), "invoice": t(r[5]),
            "qtyKg": n(r[6]), "rate": n(r[7]), "transport": n(r[9]), "payment": t(r[10]),
            "moisture": n(r[11]), "temp": n(r[12]), "odour": t(r[13]), "contamination": t(r[14]),
            "sampleDate": d(r[15]), "labDate": d(r[16]), "labResult": t(r[17]), "remarks": t(r[18]),
            "seal": t(r[19])})

    for i, r in enumerate(r for r in rows(wb["2. Pre-Composting"], 5) if r[0]):
        turns = []
        for j in (7, 11, 15):
            if r[j] is not None:
                turns.append({"date": d(r[j]), "temp": n(r[j+1]), "moisture": n(r[j+2]), "watering": d(r[j+3])})
        out["precompost"].append({"_id": f"{b}-pc-{i+1:03d}", "batch": batch, "lot": t(r[0]), "rmLot": t(r[1]),
            "start": d(r[2]), "cultureDate": d(r[3]), "cultureDose": t(r[4]), "end": d(r[5]),
            "weightMT": n(r[6]), "turns": turns, "beds": t(r[19]), "remarks": t(r[20])})

    for r in rows(wb["3. Bed Production & Ops"], 5):
        if not (r[0] and str(r[0]).startswith("BED")):
            continue
        out["beds"].append({"_id": f"{b}-{t(r[0]).lower()}", "batch": batch, "bed": t(r[0]), "dims": t(r[1]),
            "fillDate": d(r[2]), "dungT": n(r[3]), "species": t(r[4]), "wormDate": d(r[5]), "wormsKg": n(r[6]),
            "feedDate": d(r[7]), "watering": t(r[8]), "expHarvest": d(r[9]), "actHarvest": d(r[10]),
            "prodCode": t(r[11]), "pcLot": t(r[12]), "shuffleDate": d(r[13]), "status": t(r[14]) or "Filled",
            "harvestDone": d(r[15]), "harvestKg": n(r[16]), "remark": t(r[17])})

    for r in rows(wb["4. Harvesting & Packing"], 5):
        if r[1] is None:
            continue
        bed = f"BED-{int(r[1]):02d}" if isinstance(r[1], (int, float)) else t(r[1])
        net = n(r[3])
        out["harvest"].append({"_id": f"{b}-hv-{bed.lower()}", "batch": batch, "date": d(r[0]), "bed": bed,
            "rawKg": n(r[2]), "netKg": net, "prodBatch": t(r[4]), "fgBatch": t(r[5]), "packaging": t(r[6]),
            "packets": round(net / 30, 1) if net else None, "moisture": n(r[8]), "tagging": t(r[9])})

    keys = ["ph", "ec", "moisture", "oc", "om", "cn", "n", "p", "k", "ash", "ca", "mn", "fe", "mn2", "zn", "cu"]
    for i, r in enumerate(r for r in rows(wb["5. Quality Control"], 5) if r[0]):
        doc = {"_id": f"{b}-qc-{i+1:03d}", "batch": batch, "date": d(r[0]), "seal": t(r[1]), "lab": t(r[2]),
               "product": t(r[3]), "batchNo": t(r[4]), "reportDate": d(r[5]), "reportNo": t(r[6]), "status": t(r[23])}
        for k, v in zip(keys, r[7:23]):
            doc[k] = round(v * 100, 3) if isinstance(v, (int, float)) else n(v)  # cells were formatted as %
        out["qc"].append(doc)

    for i, r in enumerate(r for r in rows(wb["6. Sales Register"], 5) if r[0]):
        out["sales"].append({"_id": f"{b}-sl-{i+1:03d}", "batch": batch, "invoice": t(r[0]), "date": d(r[1]),
            "customer": t(r[2]), "drc": t(r[3]), "fgBatch": t(r[4]), "qtyKg": n(r[5]), "price": n(r[6])})

    closing = 0.0
    for i, r in enumerate(r for r in rows(wb["7. Stock Inventory"], 5) if r[0]):
        prod, sold = n(r[3]) or 0, n(r[4]) or 0
        loss = n(r[5]) if not (isinstance(r[5], str) and r[5].startswith("=")) else closing - sold
        loss = loss or 0
        closing = closing + prod - sold - loss
        out["stock"].append({"_id": f"{b}-st-{i+1:03d}", "batch": batch, "seq": i + 1, "date": "",
            "product": t(r[0]), "pack": t(r[1]), "inKg": prod, "outKg": sold, "lossKg": round(loss, 2),
            "audit": t(r[7]), "note": "Stock adjusted to zero after bulk sale" if isinstance(r[5], str) else ""})

    for i, r in enumerate(r for r in rows(wb["8. Cost & Expenses"], 5) if r[2]):
        out["expenses"].append({"_id": f"{b}-ex-{i+1:03d}", "batch": batch, "date": d(r[0]),
            "category": t(r[1]) or expense_category(t(r[2])), "desc": t(r[2]), "mode": t(r[3]),
            "amount": n(r[4]), "note": t(r[5])})

    out["batches"].append({"_id": b, "code": batch, "name": f"Batch {batch.lstrip('Bb')}",
        "start": min(x["date"] for x in out["rm"] if x["date"]), "status": "Closing", "site": "Akola",
        "notes": "Imported from Batch 4 final register"})
    return out

if __name__ == "__main__":
    data = convert(sys.argv[1], sys.argv[2])
    dest = sys.argv[3] if len(sys.argv) > 3 else "seed.json"
    json.dump(data, open(dest, "w"), indent=1, ensure_ascii=False)
    print({k: len(v) for k, v in data.items()})
