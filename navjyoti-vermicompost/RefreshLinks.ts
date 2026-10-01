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

  // [sheet, column, target sheet, target column, extra column that must also match (or "")]
  const LINKS: string[][] = [
    ["RawMaterial", "RMLot", "PreCompost", "RMLot", ""],
    ["RawMaterial", "Batch", "Batches", "BatchCode", ""],
    ["PreCompost", "RMLot", "RawMaterial", "RMLot", ""],
    ["PreCompost", "PCLot", "Beds", "PCLot", ""],
    ["PreCompost", "Batch", "Batches", "BatchCode", ""],
    ["Beds", "PCLot", "PreCompost", "PCLot", ""],
    ["Beds", "BedNo", "Harvest", "BedNo", "Batch"],
    ["Beds", "ProdCode", "Harvest", "ProdBatch", ""],
    ["Beds", "Batch", "Batches", "BatchCode", ""],
    ["Harvest", "BedNo", "Beds", "BedNo", "Batch"],
    ["Harvest", "ProdBatch", "Beds", "ProdCode", ""],
    ["Harvest", "FGBatch", "Sales", "FGBatch", ""],
    ["Harvest", "Batch", "Batches", "BatchCode", ""],
    ["QualityControl", "Batch", "Batches", "BatchCode", ""],
    ["Sales", "FGBatch", "Harvest", "FGBatch", ""],
    ["Sales", "Batch", "Batches", "BatchCode", ""],
    ["StockLedger", "Batch", "Batches", "BatchCode", ""],
    ["Expenses", "Batch", "Batches", "BatchCode", ""],
    ["Earthworm", "ToBatch", "Batches", "BatchCode", ""],
    ["DailyLog", "Batch", "Batches", "BatchCode", ""],
    ["Batches", "BatchCode", "Beds", "Batch", ""],
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

  let linked = 0;
  for (const [sheet, col, target, tcol, extra] of LINKS) {
    const src = rowsOf(sheet), tgt = rowsOf(target);
    const c = colIndex(src, col), tc = colIndex(tgt, tcol);
    if (c < 0 || tc < 0) continue;
    const ce = extra ? colIndex(src, extra) : -1, tce = extra ? colIndex(tgt, extra) : -1;
    const index: { [key: string]: number } = {};
    for (let r = 1; r < tgt.length; r++) {
      const v = String(tgt[r][tc]).trim();
      if (!v) continue;
      const k = extra ? v + "|" + String(tgt[r][tce]).trim() : v;
      if (!(k in index)) index[k] = r + 1;
    }
    const ws = workbook.getWorksheet(sheet);
    for (let r = 1; r < src.length; r++) {
      const v = String(src[r][c]).trim();
      if (!v) continue;
      const k = extra ? v + "|" + String(src[r][ce]).trim() : v;
      const row = index[k];
      if (!row) continue;
      const cell = ws.getCell(r, c);
      cell.setHyperlink({ documentReference: `'${target}'!${letter(tc)}${row}`, textToDisplay: v, screenTip: `Open in ${target}` });
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
  console.log(`Links refreshed: ${linked} reference cells linked.`);
}
