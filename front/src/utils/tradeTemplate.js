import ExcelJS from "exceljs";
import { buildTradeDocFileName } from "./quote";

const TEMPLATE_URL = `${import.meta.env.BASE_URL}sheet/거래명세서 양식.xlsx`;

const toText = (v) => String(v ?? "").trim();
const toNum = (v) => { const n = Number(v); return Number.isFinite(n) ? n : 0; };
const deepClone = (v) => (v == null ? v : JSON.parse(JSON.stringify(v)));
const cloneCellValue = (v) => {
  if (v == null) return v;
  if (typeof v === "object") return deepClone(v);
  return v;
};

const sanitizeSheetName = (name) =>
  String(name || "").replace(/[\\/*?:[\]]/g, " ").trim().slice(0, 31) || "Sheet";

const makeUniqueSheetName = (used, base) => {
  const safeBase = sanitizeSheetName(base);
  if (!used.has(safeBase)) { used.add(safeBase); return safeBase; }
  let seq = 2;
  while (seq < 1000) {
    const suffix = `-${seq}`;
    const next = sanitizeSheetName(`${safeBase.slice(0, Math.max(1, 31 - suffix.length))}${suffix}`);
    if (!used.has(next)) { used.add(next); return next; }
    seq++;
  }
  const fb = `Sheet-${String(Date.now()).slice(-6)}`;
  used.add(fb); return fb;
};

const copyWorksheet = (srcWs, dstWb, name) => {
  const ws = dstWb.addWorksheet(name);
  ws.properties = deepClone(srcWs.properties || {});
  ws.pageSetup = deepClone(srcWs.pageSetup || {});
  ws.headerFooter = deepClone(srcWs.headerFooter || {});
  ws.views = deepClone(srcWs.views || []);
  ws.state = srcWs.state || "visible";
  for (let c = 1; c <= srcWs.columnCount; c++) {
    const sc = srcWs.getColumn(c), dc = ws.getColumn(c);
    dc.width = sc.width; dc.hidden = !!sc.hidden;
    dc.outlineLevel = sc.outlineLevel || 0; dc.style = deepClone(sc.style || {});
  }
  for (let r = 1; r <= srcWs.rowCount; r++) {
    const sr = srcWs.getRow(r), dr = ws.getRow(r);
    dr.height = sr.height; dr.hidden = !!sr.hidden; dr.outlineLevel = sr.outlineLevel || 0;
    for (let c = 1; c <= srcWs.columnCount; c++) {
      const sc = sr.getCell(c), dc = dr.getCell(c);
      dc.value = cloneCellValue(sc.value); dc.style = deepClone(sc.style || {}); dc.numFmt = sc.numFmt || undefined;
    }
  }
  const merges = Array.isArray(srcWs.model?.merges) ? srcWs.model.merges : [];
  merges.forEach((m) => { try { ws.mergeCells(m); } catch {} });
};

const downloadBlob = (blob, fileName) => {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = fileName; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
};

const workbookToPdfBlob = async (workbook) => {
  const xlsxBuffer = await workbook.xlsx.writeBuffer();
  const bytes = new Uint8Array(xlsxBuffer);
  let binary = "";
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
  const res = await fetch("/api/quote/convert-pdf", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ xlsxBase64: btoa(binary) })
  });
  const json = await res.json();
  if (!res.ok || !json?.ok || !json?.pdfBase64) throw new Error(json?.message || "PDF 변환 실패");
  const pdfBinary = atob(json.pdfBase64);
  const pdfBytes = new Uint8Array(pdfBinary.length);
  for (let i = 0; i < pdfBinary.length; i++) pdfBytes[i] = pdfBinary.charCodeAt(i);
  return new Blob([pdfBytes], { type: "application/pdf" });
};

const parseCellRef = (ref) => {
  const m = String(ref || "").match(/^([A-Z]+)(\d+)$/i);
  if (!m) return null;
  return { col: m[1].toUpperCase(), row: Number(m[2]) };
};
const colToIndex = (col) =>
  String(col || "").toUpperCase().split("").reduce((a, ch) => a * 26 + (ch.charCodeAt(0) - 64), 0);
const parseRange = (range) => {
  const [a, b] = String(range || "").split(":");
  const start = parseCellRef(a), end = parseCellRef(b || a);
  if (!start || !end) return null;
  return { start, end };
};

const copyRowStyle = (ws, fromRowNo, toRowNo, maxCol = 24) => {
  const fr = ws.getRow(fromRowNo), tr = ws.getRow(toRowNo);
  tr.height = fr.height;
  for (let c = 1; c <= maxCol; c++) {
    const sc = fr.getCell(c), dc = tr.getCell(c);
    dc.style = JSON.parse(JSON.stringify(sc.style || {})); dc.numFmt = sc.numFmt || undefined;
  }
};

const replicateSingleRowMerges = (ws, sourceRowNo, targetRowNo) => {
  const merges = Array.isArray(ws.model?.merges) ? ws.model.merges : [];
  merges.forEach((m) => {
    const p = parseRange(m);
    if (!p || p.start.row !== sourceRowNo || p.end.row !== sourceRowNo) return;
    try { ws.mergeCells(targetRowNo, colToIndex(p.start.col), targetRowNo, colToIndex(p.end.col)); } catch {}
  });
};

const normalizeMergeOnRow = (ws, rowNo, startCol, endCol) => {
  const merges = Array.isArray(ws.model?.merges) ? [...ws.model.merges] : [];
  merges.forEach((m) => {
    const p = parseRange(m);
    if (!p || !(p.start.row <= rowNo && rowNo <= p.end.row)) return;
    const ms = colToIndex(p.start.col), me = colToIndex(p.end.col);
    if (!(startCol <= me && ms <= endCol)) return;
    try { ws.unMergeCells(m); } catch {}
  });
  try { ws.mergeCells(rowNo, startCol, rowNo, endCol); } catch {}
};

const clearRange = (ws, rowNo, fromCol, toCol) => {
  for (let c = fromCol; c <= toCol; c++) ws.getCell(rowNo, c).value = "";
};

const snapshotRows = (ws, startRow, endRow, maxCol = 24) => {
  const rows = [];
  for (let r = startRow; r <= endRow; r++) {
    const row = ws.getRow(r);
    const cells = [];
    for (let c = 1; c <= maxCol; c++) {
      const cell = row.getCell(c);
      cells.push({ col: c, value: cell.value ?? null, style: JSON.parse(JSON.stringify(cell.style || {})), numFmt: cell.numFmt || undefined });
    }
    rows.push({ rowNo: r, height: row.height, cells });
  }
  return rows;
};

const snapshotSingleRowMerges = (ws, startRow, endRow) => {
  const merges = Array.isArray(ws.model?.merges) ? ws.model.merges : [];
  return merges
    .map((m) => ({ raw: m, parsed: parseRange(m) }))
    .filter((x) => x.parsed && x.parsed.start.row >= startRow && x.parsed.end.row <= endRow && x.parsed.start.row === x.parsed.end.row)
    .map((x) => ({ rowOffset: x.parsed.start.row - startRow, startCol: colToIndex(x.parsed.start.col), endCol: colToIndex(x.parsed.end.col) }));
};

const restoreRows = (ws, snapshot, targetStartRow, maxCol = 24) => {
  snapshot.forEach((saved, idx) => {
    const rowNo = targetStartRow + idx;
    const row = ws.getRow(rowNo);
    row.height = saved.height;
    for (let c = 1; c <= maxCol; c++) {
      const cell = row.getCell(c);
      cell.value = null; cell.style = {}; cell.numFmt = undefined;
    }
    saved.cells.forEach((sc) => {
      const cell = row.getCell(sc.col);
      cell.value = sc.value; cell.style = JSON.parse(JSON.stringify(sc.style || {})); cell.numFmt = sc.numFmt;
    });
  });
};

const restoreSingleRowMerges = (ws, mergesSnapshot, targetStartRow, rowCount = 4) => {
  const current = Array.isArray(ws.model?.merges) ? [...ws.model.merges] : [];
  current.forEach((m) => {
    const p = parseRange(m);
    if (!p || p.start.row !== p.end.row) return;
    if (p.start.row < targetStartRow || p.start.row > targetStartRow + rowCount - 1) return;
    try { ws.unMergeCells(m); } catch {}
  });
  mergesSnapshot.forEach((m) => {
    try { ws.mergeCells(targetStartRow + m.rowOffset, m.startCol, targetStartRow + m.rowOffset, m.endCol); } catch {}
  });
};

const normalizeText = (v) => toText(v).replace(/\s+/g, " ");
const withScrapSuffix = (desc, isScrap) => {
  const text = normalizeText(desc);
  if (!isScrap) return text;
  if (/\(scrap\)\s*$/i.test(text)) return text;
  return text ? `${text}(Scrap)` : "(Scrap)";
};

const getDetailRows = (doc) => {
  const src = Array.isArray(doc?.serialRows) ? doc.serialRows : [];
  return src.map((r) => ({
    description: withScrapSuffix(r.description, String(r.status || "").trim() === "스크랩"),
    partNo: toText(r.partNo),
    qty: 1,
    serial: toText(r.serialNo)
  }));
};

const applyWrapFromRow = (ws, startRow, endRow, startCol = 1, endCol = 22) => {
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      const cell = ws.getCell(r, c);
      cell.alignment = { ...(cell.alignment || {}), wrapText: true, shrinkToFit: false, vertical: (cell.alignment && cell.alignment.vertical) || "middle" };
    }
  }
};

const fillBlock = (ws, blockStartCol, doc, details, detailRowCount, footerStartRow) => {
  const c = (offset) => blockStartCol + offset;
  ws.getCell(11, c(0)).value = normalizeText(doc?.kitName || doc?.kitNo || "SERVICE KIT");
  const _kitSerial = toText(doc?.kitSerial);
  ws.getCell(11, c(1)).value = toText(doc?.kitNo);
  ws.getCell(11, c(4)).value = toNum(doc?.totalUSD);
  const detailStart = 12;
  const baseEnd = 20 + Math.max(0, detailRowCount - 9);
  for (let r = detailStart; r <= baseEnd; r++) {
    clearRange(ws, r, c(0), c(7));
    normalizeMergeOnRow(ws, r, c(1), c(2));
    normalizeMergeOnRow(ws, r, c(4), c(6));
  }
  for (let i = 0; i < detailRowCount; i++) {
    const rowNo = detailStart + i;
    const line = details[i] || { description: "", partNo: "", qty: "", serial: "" };
    ws.getCell(rowNo, c(0)).value = line.description;
    ws.getCell(rowNo, c(1)).value = line.partNo;
    ws.getCell(rowNo, c(3)).value = line.qty || "";
    const serialCell = ws.getCell(rowNo, c(4));
    serialCell.value = line.serial;
    serialCell.alignment = { ...(serialCell.alignment || {}), horizontal: "center", vertical: "middle" };
  }
  const poRow = footerStartRow + 3;
  normalizeMergeOnRow(ws, poRow, c(0), c(7));
  const poCell = ws.getCell(poRow, c(0));
  poCell.value = `PO ${toText(doc?.poNumber || "")}\n${_kitSerial}`.trim();
  poCell.alignment = { ...(poCell.alignment || {}), wrapText: true, vertical: "top", shrinkToFit: false };
};

export async function buildTradeWorkbookFromTemplate(doc, options = {}) {
  const wb = new ExcelJS.Workbook();
  const res = await fetch(encodeURI(options.templateUrl || TEMPLATE_URL));
  if (!res.ok) throw new Error("거래명세서 양식 파일을 불러오지 못했습니다.");
  const buf = await res.arrayBuffer();
  await wb.xlsx.load(buf);
  wb.calcProperties = { fullCalcOnLoad: true };
  const ws = wb.worksheets[0];
  if (!ws) throw new Error("거래명세서 양식 시트를 찾을 수 없습니다.");

  const details = getDetailRows(doc);
  const detailRowCount = Math.max(details.length, 1);
  const baseDetailRows = 9;
  const extraRows = Math.max(0, detailRowCount - baseDetailRows);
  const footerSnapshot = snapshotRows(ws, 21, 24, 24);
  const footerMergesSnapshot = snapshotSingleRowMerges(ws, 21, 24);

  if (extraRows > 0) {
    ws.spliceRows(21, 0, ...Array.from({ length: extraRows }, () => []));
    for (let i = 0; i < extraRows; i++) {
      const targetRow = 21 + i;
      copyRowStyle(ws, 20, targetRow, 24);
      replicateSingleRowMerges(ws, 20, targetRow);
    }
  }

  const footerStartRow = 21 + extraRows;
  restoreRows(ws, footerSnapshot, footerStartRow, 24);
  restoreSingleRowMerges(ws, footerMergesSnapshot, footerStartRow, 4);

  const docDate = toText(doc?.docDate || doc?.issueDate || "");
  ws.getCell("B5").value = docDate;
  ws.getCell("N5").value = docDate;

  fillBlock(ws, 2, doc, details, detailRowCount, footerStartRow);
  fillBlock(ws, 14, doc, details, detailRowCount, footerStartRow);
  applyWrapFromRow(ws, 11, footerStartRow + 3, 1, 22);

  ws.pageSetup = {
    ...ws.pageSetup,
    fitToPage: true, fitToWidth: 1, fitToHeight: 1,
    printArea: `A1:V${footerStartRow + 3}`
  };

  return { workbook: wb, rows: details };
}

export async function downloadTradeExcel(doc) {
  const { workbook } = await buildTradeWorkbookFromTemplate(doc);
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  downloadBlob(blob, buildTradeDocFileName({ docDate: doc?.docDate, poNumber: doc?.poNumber, kitNo: doc?.kitNo }) + ".xlsx");
}

export async function downloadTradePdf(doc) {
  const { workbook } = await buildTradeWorkbookFromTemplate(doc);
  const pdfBlob = await workbookToPdfBlob(workbook);
  downloadBlob(pdfBlob, buildTradeDocFileName({ docDate: doc?.docDate, poNumber: doc?.poNumber, kitNo: doc?.kitNo }) + ".pdf");
}

export async function downloadTradeExcelBatch(docs) {
  const list = (Array.isArray(docs) ? docs : []).filter(Boolean);
  if (!list.length) return;
  const wb = new ExcelJS.Workbook();
  const used = new Set();
  for (const doc of list) {
    const { workbook } = await buildTradeWorkbookFromTemplate(doc);
    const srcWs = workbook.worksheets[0];
    const base = `${String(doc?.kitNo || "").slice(-5)}-${String(doc?.poNumber || "").slice(-6)}` || "TDOC";
    copyWorksheet(srcWs, wb, makeUniqueSheetName(used, base));
  }
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const ymd = new Date().toISOString().slice(2, 10).replace(/-/g, "");
  downloadBlob(blob, `SYQ_거래명세서_${ymd}.xlsx`);
}
