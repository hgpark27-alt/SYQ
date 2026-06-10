import { getQuoteSeqMap, setQuoteSeqMap } from "./storage";

export const SCRAP_RATIO = 0.30;

const round2 = (n) => Math.round(Number(n || 0) * 100) / 100;

const yymmdd = (dateStr) => {
  const m = String(dateStr || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return `${m[1].slice(2)}${m[2]}${m[3]}`;
  const d = new Date();
  return `${String(d.getFullYear()).slice(2)}${String(d.getMonth()+1).padStart(2,"0")}${String(d.getDate()).padStart(2,"0")}`;
};

export function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}

export function makeQuotationNumber(issueDate) {
  const key = yymmdd(issueDate);
  const seqMap = getQuoteSeqMap();
  const nextSeq = (seqMap[key] || 0) + 1;
  seqMap[key] = nextSeq;
  setQuoteSeqMap(seqMap);
  return `SYQ-${key}-${String(nextSeq).padStart(2, "0")}`;
}

export function calculateItem(item) {
  const qty = Number(item.qty) || 1;
  const baseUSD = item.isScrap
    ? round2(Number(item.cleaningPriceUSD || 0) * SCRAP_RATIO)
    : Number(item.cleaningPriceUSD || 0);
  const yoUSD    = item.yoRecoating?.on   ? Number(item.yoRecoating.priceUSD   || 0) : 0;
  const bsUSD    = item.bsRecoating?.on   ? Number(item.bsRecoating.priceUSD   || 0) : 0;
  const lidUSD   = item.lidRecoating?.on  ? Number(item.lidRecoating.priceUSD  || 0) : 0;
  const icpmsUSD = item.icpms?.on         ? Number(item.icpms.priceUSD         || 0) : 0;
  const extraUSD = item.extraCleaning?.on ? Number(item.extraCleaning.priceUSD || 0) : 0;

  const unitPriceUSD        = baseUSD;
  const yoRecoatingTotal    = round2(yoUSD    * qty);
  const bsRecoatingTotal    = round2(bsUSD    * qty);
  const lidRecoatingTotal   = round2(lidUSD   * qty);
  const icpmsTotal          = round2(icpmsUSD * qty);
  const extraCleaningTotal  = round2(extraUSD * qty);
  const totalPriceUSD = round2(unitPriceUSD * qty)
    + yoRecoatingTotal + bsRecoatingTotal + lidRecoatingTotal
    + icpmsTotal + extraCleaningTotal;

  return { unitPriceUSD, yoRecoatingTotal, bsRecoatingTotal, lidRecoatingTotal, icpmsTotal, extraCleaningTotal, totalPriceUSD };
}

export function calculateQuote(items) {
  const mapped = items.map((item) => ({ ...item, ...calculateItem(item) }));
  return {
    items: mapped,
    grandTotalUSD: mapped.reduce((acc, i) => acc + (i.totalPriceUSD || 0), 0)
  };
}

export function formatUSD(v) {
  return Number(v || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

const safeSegment = (v, fb = "-") =>
  String(v || "").trim().replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, "") || fb;

export function buildQuoteFileName({ issueDate, poNumber, kitNo, quotationNumber }) {
  const ymd = yymmdd(issueDate);
  const po = safeSegment(poNumber, "NO-PO");
  const kit = safeSegment(kitNo, "0247-XXXXX");
  return quotationNumber
    ? `SYQ_${quotationNumber}_${po}`
    : `SYQ_${ymd}_${po}_${kit}`;
}

export function buildTradeDocFileName({ docDate, poNumber, kitNo }) {
  const ymd = yymmdd(docDate);
  const po = safeSegment(poNumber, "NO-PO");
  const kit = safeSegment(kitNo, "0247-XXXXX");
  return `SYQ_거래명세서_${ymd}_${po}_${kit}`;
}

export function normalizeKitNo(input) {
  const digits = String(input || "").replace(/\D/g, "").slice(-5);
  return digits.length === 5 ? `0247-${digits}` : "";
}
