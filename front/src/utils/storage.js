import * as gs from "./gsheets.js";

// localStorage 없음 — Google Sheets가 유일한 저장소

export function bootstrapData() {}  // 구 로컬 서버 방식 호환용 (no-op)

// ── 전체 pull (앱 로드 시) ────────────────────────────────────────

export const pullFromSheets = async () => {
  const { kits, quotes, tradeDocs, config } = await gs.pullAll();
  return {
    kits,
    quotes,
    tradeDocs,
    settings: config.settings  ?? { defaultExchangeRate: 1400, quoteAuthor: "SY Kim" },
    quoteSeqMap: config.quoteSeqMap ?? {}
  };
};

// ── KITS ─────────────────────────────────────────────────────────

export const setKits = (v) => gs.saveKitsToSheet(v).catch(console.error);

// ── QUOTES ───────────────────────────────────────────────────────

export const setQuotes = (v) => gs.saveQuotesToSheet(v).catch(console.error);

// ── TRADE DOCS ───────────────────────────────────────────────────

export const setTradeDocs = (v) => gs.saveTradeDocsToSheet(v).catch(console.error);

// ── SETTINGS ─────────────────────────────────────────────────────

export const setSettings = (v) => gs.saveConfigKey("settings", v).catch(console.error);

// ── QUOTE SEQ MAP ────────────────────────────────────────────────

export const getQuoteSeqMap = () => window.__syqSeqMap ?? {};

export const setQuoteSeqMap = (v) => {
  window.__syqSeqMap = v;  // 메모리에 유지 (세션 내 시퀀스 충돌 방지)
  gs.saveConfigKey("quoteSeqMap", v).catch(console.error);
};
