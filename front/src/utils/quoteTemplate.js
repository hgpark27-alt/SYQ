import JSZip from "jszip";
import { calculateItem, buildQuoteFileName } from "./quote";

const TEMPLATE_URL = "/sheet/견적서 양식.xlsx";

const esc = (s) =>
  String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");

const inlineStr = (ref, style, v) =>
  `<c r="${ref}" s="${style}" t="inlineStr"><is><t>${esc(String(v || ""))}</t></is></c>`;

const numCell = (ref, style, v) =>
  v != null && v !== 0
    ? `<c r="${ref}" s="${style}"><v>${v}</v></c>`
    : `<c r="${ref}" s="${style}"/>`;

const numCellAlways = (ref, style, v) =>
  `<c r="${ref}" s="${style}"><v>${Number(v) || 0}</v></c>`;

function buildDataRow(rowNum, idx, item) {
  const r = rowNum;
  return [
    `<row r="${r}" spans="1:40" s="19" customFormat="1" ht="49.5" customHeight="1">`,
    numCellAlways(`A${r}`, "97", idx),
    `<c r="B${r}" s="98"/>`,
    `<c r="C${r}" s="98"/>`,
    inlineStr(`D${r}`, "99", item.partNo),
    `<c r="E${r}" s="100"/>`,
    `<c r="F${r}" s="100"/>`,
    `<c r="G${r}" s="101"/>`,
    inlineStr(`H${r}`, "99", item.description),
    `<c r="I${r}" s="100"/>`,
    `<c r="J${r}" s="100"/>`,
    `<c r="K${r}" s="101"/>`,
    numCellAlways(`L${r}`, "20", item.qty),
    numCell(`M${r}`, "106", item.unitPriceUSD),
    `<c r="N${r}" s="106"/>`,
    `<c r="O${r}" s="106"/>`,
    `<c r="P${r}" s="106"/>`,
    numCell(`Q${r}`, "71", item.reworkTotal),
    numCell(`R${r}`, "111", item.icpmsTotal),
    `<c r="S${r}" s="111"/>`,
    `<c r="T${r}" s="111"/>`,
    `<c r="U${r}" s="111"/>`,
    numCell(`V${r}`, "111", item.lpcTotal),
    `<c r="W${r}" s="111"/>`,
    `<c r="X${r}" s="111"/>`,
    `<c r="Y${r}" s="111"/>`,
    numCellAlways(`Z${r}`, "73", item.totalPriceUSD),
    item.remark
      ? inlineStr(`AA${r}`, "95", item.remark)
      : `<c r="AA${r}" s="95"/>`,
    `<c r="AB${r}" s="95"/>`,
    `<c r="AC${r}" s="95"/>`,
    `<c r="AD${r}" s="96"/>`,
    `</row>`
  ].join("");
}

function buildBlankRow(rowNum) {
  const r = rowNum;
  return [
    `<row r="${r}" spans="1:40" s="19" customFormat="1" ht="49.5" customHeight="1">`,
    `<c r="A${r}" s="97"/>`,
    `<c r="B${r}" s="98"/>`,
    `<c r="C${r}" s="98"/>`,
    `<c r="D${r}" s="99"/>`,
    `<c r="E${r}" s="100"/>`,
    `<c r="F${r}" s="100"/>`,
    `<c r="G${r}" s="101"/>`,
    `<c r="H${r}" s="99"/>`,
    `<c r="I${r}" s="100"/>`,
    `<c r="J${r}" s="100"/>`,
    `<c r="K${r}" s="101"/>`,
    `<c r="L${r}" s="20"/>`,
    `<c r="M${r}" s="106"/>`,
    `<c r="N${r}" s="106"/>`,
    `<c r="O${r}" s="106"/>`,
    `<c r="P${r}" s="106"/>`,
    `<c r="Q${r}" s="71"/>`,
    `<c r="R${r}" s="111"/>`,
    `<c r="S${r}" s="111"/>`,
    `<c r="T${r}" s="111"/>`,
    `<c r="U${r}" s="111"/>`,
    `<c r="V${r}" s="111"/>`,
    `<c r="W${r}" s="111"/>`,
    `<c r="X${r}" s="111"/>`,
    `<c r="Y${r}" s="111"/>`,
    `<c r="Z${r}" s="73"/>`,
    `<c r="AA${r}" s="95"/>`,
    `<c r="AB${r}" s="95"/>`,
    `<c r="AC${r}" s="95"/>`,
    `<c r="AD${r}" s="96"/>`,
    `</row>`
  ].join("");
}

function buildTotalRow(totalKRW) {
  return [
    `<row r="29" spans="1:40" s="19" customFormat="1" ht="49.5" customHeight="1">`,
    `<c r="A29" s="81" t="s"><v>60</v></c>`,
    `<c r="B29" s="82"/>`,
    `<c r="C29" s="82"/>`,
    `<c r="D29" s="82"/>`,
    `<c r="E29" s="82"/>`,
    `<c r="F29" s="82"/>`,
    `<c r="G29" s="82"/>`,
    `<c r="H29" s="82"/>`,
    `<c r="I29" s="82"/>`,
    `<c r="J29" s="82"/>`,
    `<c r="K29" s="82"/>`,
    `<c r="L29" s="82"/>`,
    `<c r="M29" s="82"/>`,
    `<c r="N29" s="82"/>`,
    `<c r="O29" s="82"/>`,
    `<c r="P29" s="82"/>`,
    `<c r="Q29" s="82"/>`,
    `<c r="R29" s="82"/>`,
    `<c r="S29" s="82"/>`,
    `<c r="T29" s="82"/>`,
    `<c r="U29" s="82"/>`,
    `<c r="V29" s="82"/>`,
    `<c r="W29" s="82"/>`,
    `<c r="X29" s="82"/>`,
    `<c r="Y29" s="83"/>`,
    `<c r="Z29" s="67"><v>${totalKRW}</v></c>`,
    `<c r="AA29" s="84"/>`,
    `<c r="AB29" s="84"/>`,
    `<c r="AC29" s="84"/>`,
    `<c r="AD29" s="85"/>`,
    `</row>`
  ].join("");
}

export async function generateQuoteXlsx(quote) {
  const res = await fetch(TEMPLATE_URL);
  if (!res.ok) throw new Error("견적서 양식 로드 실패");
  const buf = await res.arrayBuffer();

  const zip = await JSZip.loadAsync(buf);
  let xml = await zip.file("xl/worksheets/sheet2.xml").async("string");

  const { quotationNumber, issueDate, poNumber, kitSerial, kitNo, items } = quote;

  // Header: Quotation Number (AA8)
  xml = xml.replace(
    /<c r="AA8"\b[^>]*?>[\s\S]*?<\/c>/,
    `<c r="AA8" s="147" t="inlineStr"><is><t>${esc(quotationNumber || "")}</t></is></c>`
  );

  // Header: Date (AA9)
  xml = xml.replace(
    /<c r="AA9"\b[^>]*?>[\s\S]*?<\/c>/,
    `<c r="AA9" s="148" t="inlineStr"><is><t>${esc(issueDate || "")}</t></is></c>`
  );

  // Remove existing data rows 20-29
  xml = xml.replace(/<row\b[^>]*?\br="(2[0-9])"\b[^>]*?>[\s\S]*?<\/row>/g, "");

  // Compute prices and build rows
  let totalUSD = 0;
  let newRows = "";

  const activeItems = (items || []).slice(0, 9);
  activeItems.forEach((item, i) => {
    const calc = calculateItem(item);
    totalUSD += calc.totalPriceUSD;
    newRows += buildDataRow(20 + i, i + 1, { ...item, ...calc });
  });

  for (let i = activeItems.length; i < 9; i++) {
    newRows += buildBlankRow(20 + i);
  }

  newRows += buildTotalRow(totalUSD);

  // Insert before </sheetData>
  xml = xml.replace("</sheetData>", newRows + "</sheetData>");

  zip.file("xl/worksheets/sheet2.xml", xml);

  const blob = await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    compression: "DEFLATE"
  });

  return blob;
}

export function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

export async function downloadQuoteXlsx(quote) {
  const blob = await generateQuoteXlsx(quote);
  const name = buildQuoteFileName(quote) + ".xlsx";
  downloadBlob(blob, name);
}

export async function downloadQuotePdf(quote) {
  const blob = await generateQuoteXlsx(quote);
  const xlsxBuffer = await blob.arrayBuffer();
  const bytes = new Uint8Array(xlsxBuffer);
  let binary = "";
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
  const xlsxBase64 = btoa(binary);

  const res = await fetch("/api/quote/convert-pdf", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ xlsxBase64 })
  });
  const json = await res.json();
  if (!json?.ok || !json?.pdfBase64) throw new Error(json?.message || "PDF 변환 실패");

  const pdfBinary = atob(json.pdfBase64);
  const pdfBytes = new Uint8Array(pdfBinary.length);
  for (let i = 0; i < pdfBinary.length; i++) pdfBytes[i] = pdfBinary.charCodeAt(i);
  const pdfBlob = new Blob([pdfBytes], { type: "application/pdf" });
  downloadBlob(pdfBlob, buildQuoteFileName(quote) + ".pdf");
}
