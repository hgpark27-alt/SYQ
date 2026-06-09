const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbztYBORQgiJwOrws5nj04x5MgVXlTd1tNmojx5E7TlTUOb2MGm9vSFrT_ycGNxV3egA/exec";

const safeJson = (s, fallback) => {
  try { return s ? JSON.parse(s) : fallback; } catch { return fallback; }
};

const sheetGet = async (sheet) => {
  const res = await fetch(`${SCRIPT_URL}?sheet=${encodeURIComponent(sheet)}`, { redirect: "follow" });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "read failed");
  return (data.values || []).slice(1); // 헤더 행 제외
};

const sheetWrite = async (sheet, values) => {
  const res = await fetch(SCRIPT_URL, {
    method: "POST",
    redirect: "follow",
    body: JSON.stringify({ sheet, action: "write", values })
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "write failed");
};

// ── KITS ─────────────────────────────────────────────────────────

export const getKitsFromSheet = async () => {
  const rows = await sheetGet("KITS");
  return rows.filter(r => r[0]).map(r => ({
    id:        `kit-${String(r[0])}`,   // partNo 기반 안정적 id
    partNo:    String(r[0] || ""),
    name:      String(r[1] || ""),
    parts:     safeJson(r[2], []),
    updatedAt: String(r[3] || "")
  }));
};

export const saveKitsToSheet = async (kits) => {
  const now = new Date().toISOString();
  const values = kits.map(k => [
    k.partNo,           // 0247-XXXXX
    k.name || "",
    JSON.stringify(k.parts || []),
    now
  ]);
  await sheetWrite("KITS", values);
};

// ── QUOTES ───────────────────────────────────────────────────────

export const getQuotesFromSheet = async () => {
  const rows = await sheetGet("QUOTES");
  return rows.filter(r => r[0]).map(r => ({
    id:               String(r[0] || ""),
    quotationNumber:  String(r[1] || ""),
    issueDate:        String(r[2] || ""),
    poNumber:         String(r[3] || ""),
    kitNo:            String(r[4] || ""),
    kitSerial:        String(r[5] || ""),
    items:            safeJson(r[6], []),
    status:           String(r[7] || "draft"),
    createdAt:        String(r[8] || "")
  }));
};

export const saveQuotesToSheet = async (quotes) => {
  const values = quotes.map(q => [
    q.id, q.quotationNumber, q.issueDate,
    q.poNumber, q.kitNo, q.kitSerial,
    JSON.stringify(q.items || []),
    q.status || "draft",
    q.createdAt || new Date().toISOString()
  ]);
  await sheetWrite("QUOTES", values);
};

// ── TRADE_DOCS ───────────────────────────────────────────────────

export const getTradeDocsFromSheet = async () => {
  const rows = await sheetGet("TRADE_DOCS");
  return rows.filter(r => r[0]).map(r => ({
    id:         String(r[0] || ""),
    docDate:    String(r[1] || ""),
    poNumber:   String(r[2] || ""),
    kitNo:      String(r[3] || ""),
    kitSerial:  String(r[4] || ""),
    serialRows: safeJson(r[5], []),
    totalUSD:   Number(r[6]) || 0,
    kitName:    String(r[7] || ""),
    status:     String(r[8] || "draft"),
    createdAt:  String(r[9] || "")
  }));
};

export const saveTradeDocsToSheet = async (docs) => {
  const values = docs.map(d => [
    d.id, d.docDate, d.poNumber,
    d.kitNo, d.kitSerial,
    JSON.stringify(d.serialRows || []),
    d.totalUSD || 0, d.kitName || "",
    d.status || "draft",
    d.createdAt || new Date().toISOString()
  ]);
  await sheetWrite("TRADE_DOCS", values);
};

// ── CONFIG ───────────────────────────────────────────────────────

export const getConfigFromSheet = async () => {
  const rows = await sheetGet("CONFIG");
  const out = {};
  rows.filter(r => r[0]).forEach(r => { out[String(r[0])] = safeJson(r[1], r[1]); });
  return out;
};

export const saveConfigKey = async (key, value) => {
  const rows = await sheetGet("CONFIG");
  const existing = rows.filter(r => r[0]);
  const idx = existing.findIndex(r => String(r[0]) === key);
  if (idx >= 0) existing[idx] = [key, JSON.stringify(value)];
  else existing.push([key, JSON.stringify(value)]);
  await sheetWrite("CONFIG", existing);
};

// ── 전체 pull ────────────────────────────────────────────────────

export const pullAll = async () => {
  const [kits, quotes, tradeDocs, config] = await Promise.all([
    getKitsFromSheet(),
    getQuotesFromSheet(),
    getTradeDocsFromSheet(),
    getConfigFromSheet()
  ]);
  return { kits, quotes, tradeDocs, config };
};
