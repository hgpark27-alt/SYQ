const CLIENT_ID = "234290308020-4k4e69efbmbml7e0rf5k87u65fekcbh4.apps.googleusercontent.com";
const SPREADSHEET_ID = "16o_P4kUuTX88CcOLNuH0sSGpq5rqKoDW6Oc1pQLU1oc";
const SCOPE = "https://www.googleapis.com/auth/spreadsheets";

let _tokenClient = null;
let _accessToken = null;
let _gapiReady = false;

const waitFor = (fn) => new Promise((resolve) => {
  const t = setInterval(() => { if (fn()) { clearInterval(t); resolve(); } }, 80);
});

export const initGapi = async () => {
  if (_gapiReady) return;
  await waitFor(() => window.gapi);
  await new Promise((resolve, reject) => {
    window.gapi.load("client", async () => {
      try {
        await window.gapi.client.init({});
        await window.gapi.client.load("sheets", "v4");
        _gapiReady = true;
        resolve();
      } catch (e) { reject(e); }
    });
  });
};

export const initGsi = async () => {
  await waitFor(() => window.google?.accounts?.oauth2);
  _tokenClient = window.google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPE,
    callback: () => {}
  });
};

export const signIn = () => new Promise((resolve, reject) => {
  if (!_tokenClient) return reject(new Error("GSI not initialized"));
  _tokenClient.callback = (resp) => {
    if (resp.error) return reject(new Error(resp.error));
    _accessToken = resp.access_token;
    window.gapi.client.setToken({ access_token: _accessToken });
    resolve(_accessToken);
  };
  _tokenClient.requestAccessToken({ prompt: "" });
});

export const signOut = () => {
  if (_accessToken) {
    window.google?.accounts?.oauth2?.revoke(_accessToken, () => {});
    _accessToken = null;
    window.gapi?.client?.setToken(null);
  }
};

export const isSignedIn = () => !!_accessToken;

// ── 내부 Sheets 헬퍼 ─────────────────────────────────────────────

const sheetsGet = (range) =>
  window.gapi.client.sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range });

const sheetsUpdate = (range, values) =>
  window.gapi.client.sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID, range,
    valueInputOption: "RAW",
    resource: { values }
  });

const sheetsAppend = (range, values) =>
  window.gapi.client.sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID, range,
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS",
    resource: { values }
  });

const sheetsClear = (range) =>
  window.gapi.client.sheets.spreadsheets.values.clear({ spreadsheetId: SPREADSHEET_ID, range });

const safeJson = (s, fallback) => {
  try { return s ? JSON.parse(s) : fallback; } catch { return fallback; }
};

const now = () => new Date().toISOString();

// ── KITS ─────────────────────────────────────────────────────────

export const getKitsFromSheet = async () => {
  const res = await sheetsGet("KITS!A2:D");
  return (res.result.values || []).map(r => ({
    kitNo: r[0] || "",
    kitName: r[1] || "",
    parts: safeJson(r[2], []),
    updatedAt: r[3] || ""
  }));
};

export const saveKitsToSheet = async (kits) => {
  await sheetsClear("KITS!A2:D");
  if (!kits.length) return;
  const values = kits.map(k => [k.kitNo, k.kitName || "", JSON.stringify(k.parts || []), now()]);
  await sheetsAppend("KITS!A2", values);
};

// ── QUOTES ───────────────────────────────────────────────────────

export const getQuotesFromSheet = async () => {
  const res = await sheetsGet("QUOTES!A2:I");
  return (res.result.values || []).map(r => ({
    id: r[0], quotationNumber: r[1], issueDate: r[2],
    poNumber: r[3], kitNo: r[4], kitSerial: r[5],
    items: safeJson(r[6], []),
    status: r[7] || "draft",
    createdAt: r[8] || ""
  }));
};

export const saveQuotesToSheet = async (quotes) => {
  await sheetsClear("QUOTES!A2:I");
  if (!quotes.length) return;
  const values = quotes.map(q => [
    q.id, q.quotationNumber, q.issueDate,
    q.poNumber, q.kitNo, q.kitSerial,
    JSON.stringify(q.items || []),
    q.status || "draft",
    q.createdAt || now()
  ]);
  await sheetsAppend("QUOTES!A2", values);
};

// ── TRADE_DOCS ───────────────────────────────────────────────────

export const getTradeDocsFromSheet = async () => {
  const res = await sheetsGet("TRADE_DOCS!A2:J");
  return (res.result.values || []).map(r => ({
    id: r[0], docDate: r[1], poNumber: r[2],
    kitNo: r[3], kitSerial: r[4],
    serialRows: safeJson(r[5], []),
    totalUSD: Number(r[6]) || 0,
    kitName: r[7] || "",
    status: r[8] || "draft",
    createdAt: r[9] || ""
  }));
};

export const saveTradeDocsToSheet = async (docs) => {
  await sheetsClear("TRADE_DOCS!A2:J");
  if (!docs.length) return;
  const values = docs.map(d => [
    d.id, d.docDate, d.poNumber,
    d.kitNo, d.kitSerial,
    JSON.stringify(d.serialRows || []),
    d.totalUSD || 0,
    d.kitName || "",
    d.status || "draft",
    d.createdAt || now()
  ]);
  await sheetsAppend("TRADE_DOCS!A2", values);
};

// ── CONFIG (settings, quoteSeqMap) ───────────────────────────────

export const getConfigFromSheet = async () => {
  const res = await sheetsGet("CONFIG!A2:B");
  const out = {};
  (res.result.values || []).forEach(r => {
    if (r[0]) out[r[0]] = safeJson(r[1], r[1]);
  });
  return out;
};

export const saveConfigKey = async (key, value) => {
  const res = await sheetsGet("CONFIG!A2:B");
  const rows = res.result.values || [];
  const idx = rows.findIndex(r => r[0] === key);
  const cell = [[key, JSON.stringify(value)]];
  if (idx >= 0) {
    await sheetsUpdate(`CONFIG!A${idx + 2}:B${idx + 2}`, cell);
  } else {
    await sheetsAppend("CONFIG!A2", cell);
  }
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
