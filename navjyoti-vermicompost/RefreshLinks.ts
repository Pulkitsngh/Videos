/**
 * Navjyoti Vermicompost Tracker – Refresh links (Office Script for Excel for the web)
 *
 * Turns every reference code (RM lot, pre-compost lot, bed number, production batch,
 * FG batch, batch code) into a link that jumps to the related record. Run it after adding
 * rows: Excel for the web → Automate → this script → Run, or the "Refresh links" button.
 * It unlocks each sheet with the workbook password, updates the links and locks it again.
 */
function main(workbook: ExcelScript.Workbook) {
  const PASSWORD = "Navjyoti2026";

  // [sheet, column shown as link, key columns, target sheet, target key columns, target column to select]
  // Generated from LINKS in build_cloud_workbook.py – keep the two in step.
  const LINKS: [string, string, string[], string, string[], string][] = [
    ["RawMaterial", "RMLot", ["RMLot"], "PreCompost", ["RMLot"], "RMLot"],
    ["RawMaterial", "Batch", ["Batch"], "Batches", ["BatchCode"], "BatchCode"],
    ["PreCompost", "RMLot", ["RMLot"], "RawMaterial", ["RMLot"], "RMLot"],
    ["PreCompost", "PCLot", ["PCLot"], "Beds", ["PCLot"], "PCLot"],
    ["PreCompost", "Batch", ["Batch"], "Batches", ["BatchCode"], "BatchCode"],
    ["Beds", "PCLot", ["PCLot"], "PreCompost", ["PCLot"], "PCLot"],
    ["Beds", "BedNo", ["Batch", "BedNo"], "Harvest", ["Batch", "BedNo"], "BedNo"],
    ["Beds", "ProdCode", ["Batch", "BedNo"], "Harvest", ["Batch", "BedNo"], "ProdBatch"],
    ["Beds", "Batch", ["Batch"], "Batches", ["BatchCode"], "BatchCode"],
    ["Harvest", "BedNo", ["Batch", "BedNo"], "Beds", ["Batch", "BedNo"], "BedNo"],
    ["Harvest", "ProdBatch", ["Batch", "BedNo"], "Beds", ["Batch", "BedNo"], "ProdCode"],
    ["Harvest", "FGBatch", ["FGBatch"], "Sales", ["FGBatch"], "FGBatch"],
    ["Harvest", "Batch", ["Batch"], "Batches", ["BatchCode"], "BatchCode"],
    ["QualityControl", "Batch", ["Batch"], "Batches", ["BatchCode"], "BatchCode"],
    ["Sales", "FGBatch", ["FGBatch"], "Harvest", ["FGBatch"], "FGBatch"],
    ["Sales", "Batch", ["Batch"], "Batches", ["BatchCode"], "BatchCode"],
    ["StockLedger", "Batch", ["Batch"], "Batches", ["BatchCode"], "BatchCode"],
    ["Expenses", "Batch", ["Batch"], "Batches", ["BatchCode"], "BatchCode"],
    ["Batches", "BatchCode", ["BatchCode"], "Beds", ["Batch"], "Batch"]
  ];

  const cache: { [sheet: string]: (string | number | boolean)[][] } = {};
  const rowsOf = (name: string) => {
    if (!cache[name]) {
      const used = workbook.getWorksheet(name).getUsedRange();
      cache[name] = used ? used.getValues() : [[]];
    }
    return cache[name];
  };
  const colIndex = (values: (string | number | boolean)[][], header: string) =>
    values[0].map(v => String(v)).indexOf(header);
  const letter = (i: number) => {
    let s = "", n = i + 1;
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  };

  const touched = new Set<string>(LINKS.map(l => l[0]));
  touched.forEach(name => {
    const p = workbook.getWorksheet(name).getProtection();
    if (p.getProtected()) p.unprotect(PASSWORD);
  });

  let linked = 0, unmatched = 0;
  for (const [sheet, col, keys, target, tkeys, tsel] of LINKS) {
    const src = rowsOf(sheet), tgt = rowsOf(target);
    const c = colIndex(src, col), ts = colIndex(tgt, tsel);
    const kc = keys.map(k => colIndex(src, k)), tkc = tkeys.map(k => colIndex(tgt, k));
    if (c < 0 || ts < 0 || kc.indexOf(-1) >= 0 || tkc.indexOf(-1) >= 0) continue;
    const keyOf = (row: (string | number | boolean)[], cols: number[]) => {
      const parts = cols.map(i => String(row[i]).trim());
      return parts.some(p => p === "") ? "" : parts.join("|");
    };
    const index: { [key: string]: number } = {};
    for (let r = 1; r < tgt.length; r++) {
      const k = keyOf(tgt[r], tkc);
      if (k && !(k in index)) index[k] = r + 1;
    }
    const ws = workbook.getWorksheet(sheet);
    for (let r = 1; r < src.length; r++) {
      const v = String(src[r][c]).trim();
      if (!v) continue;
      const row = index[keyOf(src[r], kc)];
      if (!row) { unmatched++; continue; }
      const cell = ws.getCell(r, c);
      cell.setHyperlink({ documentReference: `'${target}'!${letter(ts)}${row}`, textToDisplay: v, screenTip: `Open in ${target}` });
      cell.getFormat().getFont().setColor("#1C5FB8");
      cell.getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.single);
      linked++;
    }
  }

  touched.forEach(name => {
    workbook.getWorksheet(name).getProtection().protect({
      allowAutoFilter: true, allowSort: true, allowFormatColumns: true,
      allowEditObjects: false, allowInsertRows: false, allowDeleteRows: false,
      selectionMode: ExcelScript.ProtectionSelectionMode.normal,
    }, PASSWORD);
  });
  console.log(`Links refreshed: ${linked} reference cells linked, ${unmatched} without a matching record yet.`);
}
