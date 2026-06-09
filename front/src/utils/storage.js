import * as gs from "./gsheets.js";

const KEYS = {
  kits:      "syq-kits",
  quotes:    "syq-quotes",
  tradedocs: "syq-tradedocs",
  settings:  "syq-settings",
  quoteSeq:  "syq-quote-seq"
};

// ── 로컬 캐시 헬퍼 ───────────────────────────────────────────────

const readLocal = (key, fallback) => {
  try { return JSON.parse(localStorage.getItem(key) ?? "null") ?? fallback; }
  catch { return fallback; }
};
const writeLocal = (key, value) => localStorage.setItem(key, JSON.stringify(value));

// ── 초기 부트스트랩 (구 로컬 서버 방식 — 호환 유지) ───────────────

export function bootstrapData() {
  const raw = window.__SYQ_STORAGE__;
  if (!raw || typeof raw !== "object") return;
  Object.entries(raw).forEach(([k, v]) => {
    try { localStorage.setItem(k, JSON.stringify(v)); } catch {}
  });
  try { delete window.__SYQ_STORAGE__; } catch {}
}

// ── Sheets → localStorage 전체 동기화 ───────────────────────────

export const pullFromSheets = async () => {
  const { kits, quotes, tradeDocs, config } = await gs.pullAll();
  writeLocal(KEYS.kits,      kits);
  writeLocal(KEYS.quotes,    quotes);
  writeLocal(KEYS.tradedocs, tradeDocs);
  if (config.settings)     writeLocal(KEYS.settings, config.settings);
  if (config.quoteSeqMap)  writeLocal(KEYS.quoteSeq, config.quoteSeqMap);
  return { kits, quotes, tradeDocs };
};

// ── sync 헬퍼 (fire-and-forget) ───────────────────────────────────

const sync = (fn) => { if (gs.isSignedIn()) fn().catch(console.warn); };

// ── KITS ─────────────────────────────────────────────────────────

export const getKits      = ()  => readLocal(KEYS.kits, []);
export const setKits      = (v) => {
  writeLocal(KEYS.kits, v);
  sync(() => gs.saveKitsToSheet(v));
};

// ── QUOTES ───────────────────────────────────────────────────────

export const getQuotes    = ()  => readLocal(KEYS.quotes, []);
export const setQuotes    = (v) => {
  writeLocal(KEYS.quotes, v);
  sync(() => gs.saveQuotesToSheet(v));
};

// ── TRADE DOCS ───────────────────────────────────────────────────

export const getTradeDocs = ()  => readLocal(KEYS.tradedocs, []);
export const setTradeDocs = (v) => {
  writeLocal(KEYS.tradedocs, v);
  sync(() => gs.saveTradeDocsToSheet(v));
};

// ── SETTINGS ─────────────────────────────────────────────────────

export const getSettings  = ()  => readLocal(KEYS.settings, { defaultExchangeRate: 1400, quoteAuthor: "SY Kim" });
export const setSettings  = (v) => {
  writeLocal(KEYS.settings, v);
  sync(() => gs.saveConfigKey("settings", v));
};

// ── QUOTE SEQ MAP ────────────────────────────────────────────────

export const getQuoteSeqMap = ()  => readLocal(KEYS.quoteSeq, {});
export const setQuoteSeqMap = (v) => {
  writeLocal(KEYS.quoteSeq, v);
  sync(() => gs.saveConfigKey("quoteSeqMap", v));
};
