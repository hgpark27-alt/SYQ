import React, { useState, useEffect, useCallback } from "react";
import {
  Zap, FileText, BarChart2, FolderOpen,
  ClipboardList, Archive, Settings2
} from "lucide-react";
import {
  setKits as storeSetKits,
  setQuotes as storeSetQuotes,
  setTradeDocs as storeSetTradeDocs,
  pullFromSheets
} from "./utils/storage";
import { saveAddSettingsToSheet } from "./utils/gsheets";
import {
  calculateItem, calculateQuote,
  formatUSD,
  makeQuotationNumber, normalizeKitNo,
  buildQuoteFileName,
  todayISO
} from "./utils/quote";
import { downloadQuoteXlsx, downloadQuotePdf } from "./utils/quoteTemplate";
import { downloadTradeExcel, downloadTradePdf, downloadTradeExcelBatch } from "./utils/tradeTemplate";

// ─── Token detection ──────────────────────────────────────────────────────────
const isKitNoToken     = (t) => /^0247-\d{5}$/i.test(t);
const isPOToken        = (t) => /^45\d{8,}$/.test(t);
const isPartNoToken    = (t) => !isKitNoToken(t) && /^0\d{3}-\d{5}$/.test(t);
const isKitSerialToken = (t) => {
  if (!t || isKitNoToken(t) || isPOToken(t)) return false;
  if (/^\d+$/.test(t) || /^\d{4}-\d{2}-\d{2}$/.test(t)) return false;
  if (/^IMTK/i.test(t) || /^NONE[_-]/i.test(t)) return false;
  return /^[A-Z][A-Z0-9]*(?:-[A-Z0-9]+)*-\d{1,3}[A-Z]?$/.test(t);
};
const isSerialToken = (t) => {
  if (!t || isKitNoToken(t) || isPOToken(t)) return false;
  if (/^WIC\d+-\d+$/i.test(t)) return true;
  if (/^YK\d+-\d+-\d+$/i.test(t)) return true;
  if (isKitSerialToken(t)) return false;
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return false;
  if (/^\d{2,10}-\d{2,6}-[A-Z0-9]{2,6}$/i.test(t)) return true;
  if (/\/\//.test(t)) return true;
  if (/^NONE[_-][A-Z0-9-]{6,}$/i.test(t)) return true;
  if (/^\d{12,}$/.test(t)) return true;
  return false;
};
const parseStatusValue = (t) => {
  if (/^(정상|normal|ok)$/i.test(t)) return "정상";
  if (/^(스크랩|scrap)$/i.test(t)) return "스크랩";
  return "";
};

const isSerialCandidate = (t, kitNo, kitSerial) => {
  if (!t || t === kitNo || t === kitSerial) return false;
  if (isKitNoToken(t) || isPOToken(t)) return false;
  if (/^WIC\d+-\d+$/i.test(t)) return true;
  if (/^YK\d+-\d+-\d+$/i.test(t)) return true;
  if (isKitSerialToken(t)) return false;
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return false;
  if (/^\d+$/.test(t)) return false;
  if (t.includes("₩")) return false;
  if (/^[A-Z\s]+$/.test(t) && !/\d/.test(t)) return false;
  if (/^\d{2,10}-\d{2,6}-[A-Z0-9]{2,6}$/i.test(t)) return true;
  if (/\/\//.test(t)) return true;
  if (/^NONE[_-][A-Z0-9-]{6,}$/i.test(t)) return true;
  if (/^\d{12,}$/.test(t)) return true;
  if (/^[A-Z0-9]+-[A-Z0-9]+-[A-Z0-9]+$/.test(t) && /\d{3,}/.test(t)) return true;
  if (/^\d+\s+[A-Z][-A-Z0-9]+$/i.test(t)) return true;
  return false;
};

function parseCase1(lines, kitNo, kitSerial, poNumber) {
  const parts = [];
  for (const line of lines) {
    const tokens = line.split(/\t/).map(t => t.trim()).filter(Boolean);
    const partNo = tokens.find(t => isPartNoToken(t));
    if (!partNo) continue;
    const pidx = tokens.indexOf(partNo);
    const description = tokens[pidx + 1] || "";
    const candidate   = tokens[pidx + 2] || "";
    let serialNo = "", qty = 1;
    const mEA = candidate.match(/^(\d+)EA$/i);
    if (mEA) {
      qty = parseInt(mEA[1], 10); serialNo = "*EA";
    } else if (isSerialToken(candidate)) {
      serialNo = candidate;
      const qtyTok = tokens.slice(pidx + 3).find(t => /^\d+$/.test(t) && parseInt(t) < 100);
      if (qtyTok) qty = parseInt(qtyTok, 10);
    } else {
      const qtyTok = tokens.slice(pidx + 2).find(t => /^\d+$/.test(t) && parseInt(t) < 100);
      if (qtyTok) qty = parseInt(qtyTok, 10);
    }
    parts.push({ partNo, description, serialNo, qty });
  }
  return { kitNo, kitSerial, poNumber, caseType: 1, parts };
}

function parseCase23(lines, kitNo, kitSerial, poNumber, kits) {
  const kit = kits.find(k => k.partNo === kitNo);
  const mainLine = lines.reduce((a, b) =>
    b.split(/\t/).filter(t => t.trim()).length > a.split(/\t/).filter(t => t.trim()).length ? b : a
  , lines[0]);
  const tokens = mainLine.split(/\t/).map(t => t.trim()).filter(Boolean);
  let sideNozzleQty = 8;
  for (let i = tokens.length - 1; i >= 0; i--) {
    if (/^\d+$/.test(tokens[i])) { sideNozzleQty = parseInt(tokens[i], 10); break; }
  }
  const serials = tokens.filter(t => isSerialCandidate(t, kitNo, kitSerial));
  if (!kit) {
    return { kitNo, kitSerial, poNumber, caseType: 23, parts: [], rawSerials: serials, sideNozzleQty, kitNotFound: true };
  }
  const parts = [];
  let si = 0;
  for (const p of kit.parts) {
    const isSideNozzle = p.isSideNozzle === true || /side\s*nozzle/i.test(p.description);
    const isHub = /hub/i.test(p.description) || /hub/i.test(p.partNo);
    if (isSideNozzle) {
      parts.push({ partNo: p.partNo, description: p.description, serialNo: "*EA", qty: sideNozzleQty });
    } else {
      const matched = serials[si] || "";
      // HUB: 시리얼 미감지 시 "없음" 표기 → 견적 0원 처리
      const serialNo = matched ? matched : (isHub ? "없음" : "");
      parts.push({ partNo: p.partNo, description: p.description, serialNo, qty: p.qty || 1 });
      if (matched) si++;
    }
  }
  return { kitNo, kitSerial, poNumber, caseType: 23, parts, sideNozzleQty };
}

function parseSmart(text, kits = []) {
  const lines = String(text || "").split(/\r?\n/).filter(l => l.trim());
  if (!lines.length) return null;
  const allTokens = lines.flatMap(l => l.split(/\t/).map(t => t.trim()).filter(Boolean));
  const kitNo     = allTokens.find(t => isKitNoToken(t)) || "";
  const kitSerial = allTokens.find(t => isKitSerialToken(t)) || "";
  const poNumber  = allTokens.find(t => isPOToken(t)) || "";
  const hasPartNos = lines.some(l => l.split(/\t/).some(t => isPartNoToken(t.trim())));
  return hasPartNos
    ? parseCase1(lines, kitNo, kitSerial, poNumber)
    : parseCase23(lines, kitNo, kitSerial, poNumber, kits);
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
const uid = () => `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;

// sys = { icpmsPrice, lpcPrice } — 시스템 고정가 (ADD 시트)
const newQuoteItem = (part = {}, sys = {}) => ({
  id: uid(),
  partNo: part.partNo || "",
  description: part.description || "",
  qty: Number(part.qty) || 1,
  cleaningPriceUSD: Number(part.cleaningPriceUSD) || 0,
  isScrap: false,
  yoRecoating:   { on: false,              priceUSD: Number(part.yoRecoatingPriceUSD)   || 0 },
  bsRecoating:   { on: false,              priceUSD: Number(part.bsRecoatingPriceUSD)   || 0 },
  lidRecoating:  { on: false,              priceUSD: Number(part.lidRecoatingPriceUSD)  || 0 },
  ht:    { on: false, priceUSD: Number(part.htPriceUSD)    || 0 },
  lp:    { on: false, priceUSD: Number(part.lpPriceUSD)    || 0 },
  icpms: { on: false, priceUSD: Number(sys.icpmsPrice)     || 261 },
  lpc:   { on: false, priceUSD: Number(sys.lpcPrice)       || 225 },
});

const TABS = {
  SMART:      "스마트 생성",
  QUOTE:      "견적서 생성",
  SUMMARY:    "요약정리",
  QUOTE_MGMT: "견적서 관리",
  TRADE:      "거래명세서",
  TRADE_MGMT: "거래명세서 관리",
  KIT_MGMT:   "키트 관리"
};

// ─── SmartGenerator (파싱 + 인라인 리뷰 + 견적서 생성) ───────────────────────
function SmartGenerator({ kits, quotes, persistQuotes, onSendToQuote, settings }) {
  const [text, setText]   = useState("");
  const [parsed, setParsed] = useState(null);

  const handleParse = () => setParsed(parseSmart(text, kits));

  const handleCreate = () => {
    if (!parsed) return;
    const { kitNo, kitSerial, poNumber, parts, kitNotFound, sideNozzleQty } = parsed;
    if (kitNotFound) { alert(`미등록 키트: ${kitNo}\n0247 관리에서 먼저 등록하세요.`); return; }
    const kit = kits.find(k => k.partNo === kitNo);
    if (!kit) { alert(`미등록 키트: ${kitNo}`); return; }
    const today = todayISO();
    const items = kit.parts.map(p => {
      const pd = parts.find(pp => pp.partNo === p.partNo);
      const isSideNozzle = p.isSideNozzle === true || /side\s*nozzle/i.test(p.description);
      const isHub        = /hub/i.test(p.description);
      // 사이드 노즐: 파싱된 수량 우선 적용
      if (isSideNozzle) {
        const qty = pd?.qty || sideNozzleQty || p.qty;
        return newQuoteItem({ ...p, qty }, settings);
      }
            // HUB: 시리얼 없으면(없음 또는 빈값) 전 가격 0원
      if (isHub && (!pd?.serialNo || pd?.serialNo === "없음")) {
        return newQuoteItem({ ...p, cleaningPriceUSD: 0, yoRecoatingPriceUSD: 0, bsRecoatingPriceUSD: 0, lidRecoatingPriceUSD: 0, htPriceUSD: 0, lpPriceUSD: 0 }, settings);
      }
      return newQuoteItem(p, settings);
    });
    const smartSerialRows = parts
      .filter(p => p.serialNo !== "*EA")
      .map(p => ({ id: uid(), partNo: p.partNo, description: p.description, serialNo: p.serialNo || "", status: "", kitSerial }));
    const newDraft = {
      id: `syq-draft-${uid()}`,
      recordKey: `SYQ-${today.replace(/-/g,"")}-${uid().slice(0,4).toUpperCase()}`,
      status: "DRAFT", quotationNumber: "", issueDate: today,
      poNumber: poNumber || "", kitSerial: kitSerial || "",
      kitNo: kitNo || "", kitName: kit.name || "",
      grandTotalUSD: 0, items, smartSerialRows, smartDocDate: today,
      updatedAt: new Date().toISOString(), completedAt: ""
    };
    persistQuotes([...quotes, newDraft]);
    onSendToQuote(newDraft.id);
  };

  return (
    <div>
      <div className="card">
        <div style={{display:"flex", gap:8, alignItems:"flex-start"}}>
          <textarea
            style={{flex:1, height:72, resize:"vertical", fontSize:12, fontFamily:"monospace",
              border:"1px solid #e2e8f0", borderRadius:8, padding:"8px", outline:"none", color:"#1e293b"}}
            placeholder="고객 시스템 데이터 붙여넣기 (Ctrl+V)…"
            value={text}
            onChange={e => setText(e.target.value)}
          />
          <div style={{display:"flex", flexDirection:"column", gap:5, paddingTop:2}}>
            <button className="btn btn-primary" style={{fontSize:12}} onClick={handleParse} disabled={!text.trim()}>파싱</button>
            <button className="btn btn-ghost" style={{fontSize:11}} onClick={() => { setText(""); setParsed(null); }}>초기화</button>
          </div>
        </div>
      </div>

      {parsed && (
        <div className="card" style={{marginTop:8}}>
          <div style={{display:"flex", alignItems:"center", gap:8, marginBottom:12}}>
            <span style={{fontWeight:700, fontSize:14, color:"#1e293b"}}>파싱 결과</span>
            <span className="badge badge-green">Case {parsed.caseType}</span>
            {parsed.kitNotFound && <span style={{background:"#fef3c7",color:"#b45309",padding:"2px 8px",borderRadius:6,fontSize:11,fontWeight:600}}>키트 미등록</span>}
          </div>

          <div style={{display:"flex", gap:24, marginBottom:12, flexWrap:"wrap",
            padding:"10px 16px", background:"#f8fafc", borderRadius:8, border:"1px solid #e2e8f0"}}>
            {[["0247", parsed.kitNo], ["Kit Serial", parsed.kitSerial], ["PO", parsed.poNumber]].map(([lbl,val]) => (
              <div key={lbl}>
                <div style={{fontSize:10, color:"#94a3b8", fontWeight:600, textTransform:"uppercase", marginBottom:3}}>{lbl}</div>
                <div style={{fontSize:13, fontWeight:600, color:"#1e293b"}}>{val || "—"}</div>
              </div>
            ))}
          </div>

          {parsed.kitNotFound && (
            <div className="alert alert-warn" style={{marginBottom:10}}>
              미등록 키트 {parsed.kitNo} — 0247 관리에서 먼저 등록하세요.
            </div>
          )}

          <div className="table-wrap">
            <table>
              <thead>
                <tr><th>#</th><th>Part No</th><th>Description</th><th>Serial No</th><th>Qty</th></tr>
              </thead>
              <tbody>
                {parsed.parts.map((p, i) => (
                  <tr key={i}>
                    <td className="td-center" style={{color:"#6b7280"}}>{i+1}</td>
                    <td style={{fontFamily:"monospace",fontSize:11}}>{p.partNo || "—"}</td>
                    <td>{p.description || "—"}</td>
                    <td style={{fontFamily:"monospace",fontSize:11}}>
                      {p.serialNo === "*EA"
                        ? <span className="badge badge-green">*EA × {p.qty}</span>
                        : p.serialNo || <span style={{color:"#e07b2a"}}>없음</span>}
                    </td>
                    <td className="td-center">{p.qty}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {!parsed.kitNotFound && parsed.parts.length > 0 && (
            <div className="actions" style={{marginTop:12}}>
              <button className="btn btn-success" onClick={handleCreate}>견적서 생성 →</button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ─── AddonPill (컴포넌트 외부에 정의 — 내부 정의 시 렌더마다 언마운트/리마운트) ─────────
const PILL_COLOR = {
  yoRecoating:  { bg:"#dbeafe", text:"#1d4ed8", border:"#93c5fd" },
  bsRecoating:  { bg:"#f3e8ff", text:"#7c3aed", border:"#c4b5fd" },
  lidRecoating: { bg:"#dcfce7", text:"#15803d", border:"#86efac" },
  ht:           { bg:"#fff7ed", text:"#c2410c", border:"#fdba74" },
  lp:           { bg:"#f0fdfa", text:"#0f766e", border:"#5eead4" },
  icpms:        { bg:"#fef9c3", text:"#854d0e", border:"#fde047" },
  lpc:          { bg:"#fdf4ff", text:"#9333ea", border:"#e879f9" },
};
const PILL_LABEL = {
  yoRecoating:  "YO-RC",
  bsRecoating:  "BS-RC",
  lidRecoating: "LID",
  ht:           "HT",
  lp:           "LP",
  icpms:        "ICPMS",
  lpc:          "LPC",
};

// 가격은 0247에서만 설정 — 견적서에서는 클릭 토글만. 활성 시 라벨+가격 세로 스택으로 가로 확장 방지.
function AddonPill({ item, idx, addon, setAddon, isCompleted }) {
  const st    = item[addon];
  const col   = PILL_COLOR[addon];
  const label = PILL_LABEL[addon];

  if (!st.on) {
    if (isCompleted) return null;
    return (
      <span onClick={() => setAddon(idx, addon, { on: true })}
        style={{display:"inline-flex", alignItems:"center", justifyContent:"center",
          width:52, padding:"4px 0", borderRadius:6, fontSize:10, fontWeight:600,
          background:"#f8fafc", color:"#94a3b8", border:"1px dashed #d1d5db",
          cursor:"pointer", userSelect:"none"}}>
        {label}
      </span>
    );
  }
  return (
    <span onClick={() => !isCompleted && setAddon(idx, addon, { on: false })}
      style={{display:"inline-flex", flexDirection:"column", alignItems:"center",
        width:52, padding:"3px 0", borderRadius:6,
        background:col.bg, color:col.text, border:`1px solid ${col.border}`,
        cursor: isCompleted ? "default" : "pointer", userSelect:"none", lineHeight:1.3}}>
      <span style={{fontSize:10, fontWeight:700}}>{label}</span>
      <span style={{fontSize:10, fontWeight:600}}>${formatUSD(st.priceUSD)}</span>
    </span>
  );
}

// ─── QuoteBuilder ─────────────────────────────────────────────────────────────
function QuoteBuilder({ quote, kits, quotes, persistQuotes, onComplete }) {
  const [local, setLocal] = useState(null);
  const [busy, setBusy]   = useState(false);
  const [msg, setMsg]     = useState("");

  useEffect(() => {
    if (!quote) { setLocal(null); return; }
    setLocal({
      ...quote,
      issueDate: quote.issueDate || todayISO(),
      items: quote.items.map((it) => ({
        ...it,
        yoRecoating:  { on: false, priceUSD: 0, ...it.yoRecoating  },
        bsRecoating:  { on: false, priceUSD: 0, ...it.bsRecoating  },
        lidRecoating: { on: false, priceUSD: 0, ...it.lidRecoating },
        ht:           { on: false, priceUSD: 0, ...it.ht           },
        lp:           { on: false, priceUSD: 0, ...it.lp           },
        icpms:        { on: false, priceUSD: 0, ...it.icpms        },
        lpc:          { on: false, priceUSD: 0, ...it.lpc          },
      }))
    });
  }, [quote?.id]);

  if (!local) return (
    <div className="card">
      <div className="empty-state">
        
        <p>좌측에서 초안을 선택하거나 스마트 생성에서 초안을 만들어주세요.</p>
      </div>
    </div>
  );

  const isCompleted = local.status === "COMPLETED";
  const grandTotal  = local.items.reduce((s, it) => s + calculateItem(it).totalPriceUSD, 0);

  const setField = (key, val) => setLocal(p => ({ ...p, [key]: val }));
  const setItem  = (idx, patch) =>
    setLocal(p => ({ ...p, items: p.items.map((it, i) => i === idx ? { ...it, ...patch } : it) }));
  const setAddon = (idx, addon, patch) =>
    setLocal(p => ({ ...p, items: p.items.map((it, i) => i === idx ? { ...it, [addon]: { ...it[addon], ...patch } } : it) }));
  const setAllAddon = (addon, on) =>
    setLocal(p => ({ ...p, items: p.items.map(it => ({ ...it, [addon]: { ...it[addon], on } })) }));

  const saveLocal = (complete = false) => {
    if (complete && local.items.length === 0) { alert("항목이 없습니다."); return; }
    if (complete && !local.issueDate) { alert("발행일을 입력하세요."); return; }
    const { items } = calculateQuote(local.items);
    const grandTotalUSD = items.reduce((s, i) => s + i.totalPriceUSD, 0);
    let updated = { ...local, items, grandTotalUSD, updatedAt: new Date().toISOString() };
    if (complete) {
      updated.quotationNumber = makeQuotationNumber(updated.issueDate);
      updated.status = "COMPLETED";
      updated.completedAt = new Date().toISOString();
    }
    persistQuotes(quotes.map(q => q.id === updated.id ? updated : q));
    setLocal(updated);
    setMsg(complete ? "완료 저장됨" : "저장됨");
    setTimeout(() => setMsg(""), 2000);
    if (complete && onComplete) onComplete(updated);
  };

  const getCalcQuote = () => {
    const { items } = calculateQuote(local.items);
    return { ...local, items, grandTotalUSD: items.reduce((s, i) => s + i.totalPriceUSD, 0) };
  };

  const handleXlsx = async () => {
    setBusy(true);
    try { await downloadQuoteXlsx(getCalcQuote()); } catch (e) { alert("xlsx 오류: " + e.message); }
    finally { setBusy(false); }
  };
  const handlePdf = async () => {
    setBusy(true);
    try { await downloadQuotePdf(getCalcQuote()); } catch (e) { alert("PDF 오류: " + e.message); }
    finally { setBusy(false); }
  };

  // AddonPill은 컴포넌트 외부에 정의됨 (위 참조)

  return (
    <div className="card">
      {/* ── 헤더 ── */}
      <div style={{display:"flex", alignItems:"center", gap:10, marginBottom:16, flexWrap:"wrap"}}>
        <span style={{fontSize:15, fontWeight:700, color:"#1e293b"}}>견적서</span>
        {isCompleted && <span className="badge badge-green">완료</span>}
        {local.quotationNumber && (
          <span style={{fontSize:12, color:"#2d73ba", fontWeight:600, background:"#eff6ff",
            padding:"2px 10px", borderRadius:6, border:"1px solid #bfdbfe"}}>
            {local.quotationNumber}
          </span>
        )}
        <div style={{flex:1}} />
      </div>

      {/* ── 메타 정보 ── */}
      <div style={{display:"flex", gap:20, flexWrap:"wrap", marginBottom:18,
        padding:"12px 18px", background:"#f8fafc", borderRadius:10, border:"1px solid #e2e8f0"}}>
        {[["Kit No", local.kitNo], ["PO", local.poNumber], ["Kit Serial", local.kitSerial]].map(([lbl, val]) => (
          <div key={lbl}>
            <div style={{fontSize:10, color:"#94a3b8", fontWeight:600, letterSpacing:.5, textTransform:"uppercase", marginBottom:3}}>{lbl}</div>
            <div style={{fontSize:13, fontWeight:600, color:"#1e293b"}}>{val || "—"}</div>
          </div>
        ))}
        <div>
          <div style={{fontSize:10, color:"#94a3b8", fontWeight:600, letterSpacing:.5, textTransform:"uppercase", marginBottom:3}}>발행일</div>
          {isCompleted
            ? <div style={{fontSize:13, fontWeight:600, color:"#1e293b"}}>{local.issueDate || "—"}</div>
            : <input type="date" value={local.issueDate || ""}
                style={{fontSize:12, border:"1px solid #e2e8f0", borderRadius:6, padding:"3px 8px", color:"#1e293b"}}
                onChange={e => setField("issueDate", e.target.value)} />
          }
        </div>
      </div>

      {local.items.length > 9 && (
        <div className="alert alert-warn" style={{marginBottom:12}}>항목이 9개 초과 — 견적서 양식은 최대 9개까지 출력됩니다.</div>
      )}

      {/* ── 파트 목록 ── */}
      <div style={{display:"flex", flexDirection:"column", gap:4}}>
        {local.items.map((item, idx) => {
          const calc = calculateItem(item);
          return (
            <div key={item.id} style={{
              display:"flex", alignItems:"center", gap:8,
              padding:"8px 12px",
              background: item.isScrap ? "#fffbf7" : "#fff",
              border: `1px solid ${item.isScrap ? "#fcd34d" : "#e2e8f0"}`,
              borderRadius:8
            }}>
              {/* 번호 */}
              <div style={{width:18, flexShrink:0, color:"#cbd5e1", fontSize:11, fontWeight:700, textAlign:"center"}}>{idx+1}</div>

              {/* 파트 정보 */}
              <div style={{flex:1, minWidth:0}}>
                <div style={{fontSize:12, fontWeight:700, color:"#0f172a"}}>{item.partNo || "—"}</div>
                <div style={{fontSize:11, color:"#64748b"}}>{item.description || "—"}<span style={{color:"#94a3b8", marginLeft:5}}>× {item.qty}</span></div>
              </div>

              {/* 세정가 */}
              <div style={{textAlign:"right", flexShrink:0, minWidth:56}}>
                <div style={{fontSize:9, color:"#94a3b8", textTransform:"uppercase", letterSpacing:.4}}>세정가</div>
                <div style={{fontSize:11, fontWeight:600, color:"#94a3b8"}}>${formatUSD(item.cleaningPriceUSD)}</div>
              </div>

              {/* 조건 pills — 고정 너비 칸, 가로 확장 없음 */}
              <div style={{display:"flex", gap:4, alignItems:"center", flexShrink:0}}>
                {/* SCRAP */}
                {!isCompleted
                  ? <span onClick={() => setItem(idx, { isScrap: !item.isScrap })}
                      style={{display:"inline-flex", alignItems:"center", justifyContent:"center",
                        width:52, padding:"4px 0", borderRadius:6, fontSize:10, fontWeight:600,
                        cursor:"pointer", userSelect:"none",
                        background: item.isScrap ? "#fef3c7" : "#f8fafc",
                        color: item.isScrap ? "#b45309" : "#94a3b8",
                        border: item.isScrap ? "1px solid #fcd34d" : "1px dashed #d1d5db"}}>
                      SCRAP
                    </span>
                  : item.isScrap && (
                    <span style={{display:"inline-flex", alignItems:"center", justifyContent:"center",
                      width:52, padding:"4px 0", borderRadius:6, fontSize:10, fontWeight:600,
                      background:"#fef3c7", color:"#b45309", border:"1px solid #fcd34d"}}>
                      SCRAP
                    </span>
                  )
                }
                <AddonPill key={`${idx}-yo`}    item={item} idx={idx} addon="yoRecoating"   setAddon={setAddon} isCompleted={isCompleted} />
                <AddonPill key={`${idx}-bs`}    item={item} idx={idx} addon="bsRecoating"   setAddon={setAddon} isCompleted={isCompleted} />
                <AddonPill key={`${idx}-lid`}   item={item} idx={idx} addon="lidRecoating"  setAddon={setAddon} isCompleted={isCompleted} />
                <AddonPill key={`${idx}-ht`}    item={item} idx={idx} addon="ht"    setAddon={setAddon} isCompleted={isCompleted} />
                <AddonPill key={`${idx}-lp`}    item={item} idx={idx} addon="lp"    setAddon={setAddon} isCompleted={isCompleted} />
                <AddonPill key={`${idx}-icpms`} item={item} idx={idx} addon="icpms" setAddon={setAddon} isCompleted={isCompleted} />
                <AddonPill key={`${idx}-lpc`}   item={item} idx={idx} addon="lpc"   setAddon={setAddon} isCompleted={isCompleted} />
              </div>

              {/* 합계 */}
              <div style={{textAlign:"right", flexShrink:0, minWidth:60}}>
                <div style={{fontSize:9, color:"#94a3b8", textTransform:"uppercase", letterSpacing:.4}}>합계</div>
                <div style={{fontSize:13, fontWeight:700, color:"#1e293b"}}>${formatUSD(calc.totalPriceUSD)}</div>
              </div>
            </div>
          );
        })}
      </div>

      {/* ── 총합 ── */}
      <div style={{display:"flex", justifyContent:"flex-end", alignItems:"baseline", gap:8,
        marginTop:12, padding:"12px 18px", background:"#f8fafc", borderRadius:10, border:"1px solid #e2e8f0"}}>
        <span style={{fontSize:12, color:"#94a3b8"}}>총 합계</span>
        <span style={{fontSize:20, fontWeight:800, color:"#1e293b", letterSpacing:-0.5}}>$ {formatUSD(grandTotal)}</span>
      </div>

      {/* ── 액션 ── */}
      <div className="actions" style={{marginTop:14}}>
        {!isCompleted && (
          <>
            <button className="btn btn-outline" onClick={() => saveLocal(false)}>임시저장</button>
            <button className="btn btn-success" onClick={() => saveLocal(true)}>✓ 완료 저장</button>
          </>
        )}
        <button className="btn btn-primary" onClick={handleXlsx} disabled={busy}>
          {busy ? <span className="spinner"/> : "xlsx"}
        </button>
        <button className="btn btn-outline" onClick={handlePdf} disabled={busy}>
          {busy ? <span className="spinner"/> : "PDF"}
        </button>
        {msg && <span style={{fontSize:12, color:"#3ab06b", fontWeight:600}}>✓ {msg}</span>}
      </div>
    </div>
  );
}

// ─── SummaryTab ────────────────────────────────────────────────────────────────
function SummaryTab({ quotes, persistQuotes, onSelectQuote }) {
  const drafts    = quotes.filter((q) => q.status === "DRAFT");
  const completed = quotes.filter((q) => q.status === "COMPLETED");
  const total     = completed.reduce((s, q) => s + (q.grandTotalUSD || 0), 0);

  const del = (id) => {
    if (!confirm("삭제하시겠습니까?")) return;
    persistQuotes(quotes.filter((q) => q.id !== id));
  };

  return (
    <div>
      <div className="summary-bar">
        <div className="summary-item">
          <span className="summary-label">전체</span>
          <span className="summary-val">{quotes.length}</span>
        </div>
        <div className="summary-sep"/>
        <div className="summary-item">
          <span className="summary-label">초안</span>
          <span className="summary-val">{drafts.length}</span>
        </div>
        <div className="summary-sep"/>
        <div className="summary-item">
          <span className="summary-label">완료</span>
          <span className="summary-val">{completed.length}</span>
        </div>
        <div className="summary-sep"/>
        <div className="summary-item">
          <span className="summary-label">완료 합계</span>
          <span className="summary-val">$ {formatUSD(total)}</span>
        </div>
      </div>

      <div className="card" style={{padding:0}}>
        {quotes.length === 0 ? (
          <div className="empty-state"><p>견적서가 없습니다. 스마트 생성에서 만들어보세요.</p></div>
        ) : (
          <table>
            <thead>
              <tr>
                <th>상태</th><th>PO</th><th>Kit No</th><th>발행일</th>
                <th>항목</th><th style={{textAlign:"right"}}>금액(USD)</th><th></th>
              </tr>
            </thead>
            <tbody>
              {quotes.map((q) => (
                <tr key={q.id}>
                  <td>
                    {q.status === "COMPLETED"
                      ? <span className="badge badge-green">완료</span>
                      : <span className="badge badge-gray">초안</span>}
                  </td>
                  <td style={{fontFamily:"monospace",fontSize:11,color:"var(--text-sec)"}}>{q.poNumber || "—"}</td>
                  <td style={{fontWeight:500}}>{q.kitNo || "—"}</td>
                  <td style={{color:"var(--text-sec)"}}>{q.issueDate}</td>
                  <td className="td-center" style={{color:"var(--text-sec)"}}>{q.items.length}</td>
                  <td className="td-num" style={{fontWeight:600}}>
                    {q.grandTotalUSD ? `$${formatUSD(q.grandTotalUSD)}` : "—"}
                  </td>
                  <td>
                    <div style={{display:"flex",gap:4}}>
                      {q.status === "DRAFT" && (
                        <button className="btn btn-ghost" style={{padding:"3px 10px",fontSize:11}} onClick={() => onSelectQuote(q.id)}>편집</button>
                      )}
                      <button className="btn btn-ghost" style={{padding:"3px 8px",fontSize:11,color:"var(--danger)"}} onClick={() => del(q.id)}>삭제</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

// ─── QuoteManager ─────────────────────────────────────────────────────────────
function QuoteManager({ quotes, persistQuotes }) {
  const [search, setSearch] = useState("");
  const [busy, setBusy] = useState({});

  const completed = quotes.filter((q) => q.status === "COMPLETED");
  const filtered  = completed.filter((q) => {
    const s = search.toLowerCase();
    return !s || [q.quotationNumber, q.poNumber, q.kitNo, q.kitSerial].some((v) => (v || "").toLowerCase().includes(s));
  });

  const del = (id) => {
    if (!confirm("삭제하시겠습니까?")) return;
    persistQuotes(quotes.filter((q) => q.id !== id));
  };

  const dlXlsx = async (q) => {
    setBusy((p) => ({ ...p, [q.id]: "xlsx" }));
    try { await downloadQuoteXlsx(q); } catch (e) { alert(e.message); }
    finally { setBusy((p) => ({ ...p, [q.id]: null })); }
  };

  const dlPdf = async (q) => {
    setBusy((p) => ({ ...p, [q.id]: "pdf" }));
    try { await downloadQuotePdf(q); } catch (e) { alert(e.message); }
    finally { setBusy((p) => ({ ...p, [q.id]: null })); }
  };

  return (
    <div className="card">
      <div className="card-title">완료 견적서 관리</div>
      <div className="search-bar">
        <input placeholder="견적번호 / PO / Kit No 검색..." value={search} onChange={(e) => setSearch(e.target.value)} />
        <span style={{fontSize:12,color:"#6b7280"}}>{filtered.length}건</span>
      </div>
      {filtered.length === 0 ? (
        <div className="empty-state"><p>완료된 견적서가 없습니다.</p></div>
      ) : (
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>견적번호</th><th>PO</th><th>Kit No</th><th>Kit Serial</th>
                <th>발행일</th><th>합계(USD)</th><th>항목</th><th></th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((q) => (
                <tr key={q.id}>
                  <td style={{fontWeight:700,color:"#2d73ba",fontSize:12}}>{q.quotationNumber}</td>
                  <td>{q.poNumber}</td>
                  <td>{q.kitNo}</td>
                  <td style={{fontFamily:"monospace",fontSize:11}}>{q.kitSerial}</td>
                  <td>{q.issueDate}</td>
                  <td className="td-num">${formatUSD(q.grandTotalUSD)}</td>
                  <td className="td-center">{q.items.length}개</td>
                  <td>
                    <div style={{display:"flex",gap:4}}>
                      <button className="btn btn-primary" style={{padding:"3px 10px",fontSize:11}} disabled={!!busy[q.id]} onClick={() => dlXlsx(q)}>
                        {busy[q.id] === "xlsx" ? <span className="spinner"/> : "xlsx"}
                      </button>
                      <button className="btn btn-outline" style={{padding:"3px 10px",fontSize:11}} disabled={!!busy[q.id]} onClick={() => dlPdf(q)}>
                        {busy[q.id] === "pdf" ? <span className="spinner"/> : "PDF"}
                      </button>
                      <button className="btn btn-danger" style={{padding:"3px 8px",fontSize:11}} onClick={() => del(q.id)}>✕</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ─── TradeDocBuilder ──────────────────────────────────────────────────────────
function TradeDocBuilder({ quotes, tradeDocs, persistTradeDocs }) {
  const [quoteId, setQuoteId] = useState("");
  const [local, setLocal] = useState(null);
  const [busy, setBusy] = useState(false);

  const completedQuotes = quotes.filter((q) => q.status === "COMPLETED");

  const pickQuote = (id) => {
    setQuoteId(id);
    if (!id) { setLocal(null); return; }
    const q = completedQuotes.find((q) => q.id === id);
    if (!q) return;

    const existing = tradeDocs.find((d) => d.quoteId === id);
    if (existing) { setLocal({ ...existing, serialRows: existing.serialRows.map((r) => ({ ...r })) }); return; }

    const serialRows = [];
    (q.items || []).forEach((item) => {
      const n = Number(item.qty) || 1;
      for (let i = 0; i < n; i++) {
        const smart = (q.smartSerialRows || []).find(
          (sr) => sr.partNo === item.partNo && !serialRows.some((r) => r.serialNo && r.serialNo === sr.serialNo)
        );
        serialRows.push({
          id: uid(),
          partNo: item.partNo,
          description: item.description,
          serialNo: smart?.serialNo || "",
          status: smart?.status || (item.isScrap ? "스크랩" : ""),
          itemIndex: i
        });
      }
    });

    setLocal({
      id: `syq-tdoc-draft-${uid()}`,
      quoteId: id,
      status: "DRAFT",
      docDate: q.smartDocDate || todayISO(),
      poNumber: q.poNumber,
      kitSerial: q.kitSerial,
      kitNo: q.kitNo,
      kitName: q.kitName,
      quotationNumber: q.quotationNumber,
      totalUSD: q.grandTotalUSD,
      items: (q.items || []).map((it) => ({
        partNo: it.partNo, description: it.description, qty: it.qty, totalPriceUSD: it.totalPriceUSD
      })),
      serialRows,
      updatedAt: new Date().toISOString(),
      completedAt: ""
    });
  };

  const setSerial = (idx, val) =>
    setLocal((p) => ({ ...p, serialRows: p.serialRows.map((r, i) => i === idx ? { ...r, serialNo: val } : r) }));
  const setStatus = (idx, val) =>
    setLocal((p) => ({ ...p, serialRows: p.serialRows.map((r, i) => i === idx ? { ...r, status: val } : r) }));

  const save = (complete = false) => {
    if (!local) return;
    const emptyCount = local.serialRows.filter((r) => !r.serialNo.trim()).length;
    if (complete && emptyCount > 0 && !confirm(`시리얼 번호가 ${emptyCount}개 비어있습니다. 그래도 완료 저장하시겠습니까?`)) return;
    const updated = {
      ...local,
      status: complete ? "COMPLETED" : "DRAFT",
      completedAt: complete ? new Date().toISOString() : local.completedAt,
      updatedAt: new Date().toISOString()
    };
    const exists = tradeDocs.find((d) => d.id === updated.id);
    persistTradeDocs(exists ? tradeDocs.map((d) => (d.id === updated.id ? updated : d)) : [...tradeDocs, updated]);
    setLocal(updated);
  };

  const dlXlsx = async () => {
    setBusy(true);
    try { await downloadTradeExcel(local); } catch (e) { alert(e.message); }
    finally { setBusy(false); }
  };

  const dlPdf = async () => {
    setBusy(true);
    try { await downloadTradePdf(local); } catch (e) { alert(e.message); }
    finally { setBusy(false); }
  };

  const filledCount = local ? local.serialRows.filter((r) => r.serialNo.trim()).length : 0;
  const totalCount  = local ? local.serialRows.length : 0;

  return (
    <div>
      <div className="card">
        <div className="card-title">거래명세서 생성</div>
        <div className="form-row">
          <div className="form-group" style={{flex:1}}>
            <label>완료 견적서 선택</label>
            <select value={quoteId} onChange={(e) => pickQuote(e.target.value)}>
              <option value="">-- 견적서 선택 --</option>
              {completedQuotes.map((q) => (
                <option key={q.id} value={q.id}>
                  {q.quotationNumber} | {q.poNumber} | {q.kitNo} | {q.issueDate}
                </option>
              ))}
            </select>
          </div>
          {local && (
            <div className="form-group">
              <label>출하일</label>
              <input type="date" value={local.docDate || ""}
                onChange={(e) => setLocal((p) => ({ ...p, docDate: e.target.value }))} />
            </div>
          )}
        </div>

        {!local && completedQuotes.length === 0 && (
          <div className="alert alert-info">완료된 견적서가 없습니다. 견적서 생성 탭에서 먼저 완료 저장하세요.</div>
        )}

        {local && (
          <>
            <div className="header-info">
              <div className="header-info-item"><label>견적번호</label><span>{local.quotationNumber}</span></div>
              <div className="header-info-item"><label>PO</label><span>{local.poNumber}</span></div>
              <div className="header-info-item"><label>Kit No</label><span>{local.kitNo}</span></div>
              <div className="header-info-item"><label>Kit Serial</label><span style={{fontFamily:"monospace"}}>{local.kitSerial}</span></div>
              <div className="header-info-item"><label>합계(USD)</label><span style={{fontWeight:700}}>${formatUSD(local.totalUSD)}</span></div>
              <div className="header-info-item">
                <label>시리얼 입력</label>
                <span style={{color: filledCount === totalCount ? "#3ab06b" : "#e07b2a", fontWeight:700}}>
                  {filledCount}/{totalCount}
                </span>
              </div>
            </div>

            <div className="table-wrap">
              <table className="serial-rows-table">
                <thead>
                  <tr><th>#</th><th>Part No</th><th>Description</th><th>Serial No</th><th>Status</th></tr>
                </thead>
                <tbody>
                  {local.serialRows.map((row, idx) => (
                    <tr key={row.id}>
                      <td className="td-center" style={{color:"#6b7280"}}>{idx + 1}</td>
                      <td style={{fontFamily:"monospace",fontSize:11}}>{row.partNo}</td>
                      <td>{row.description}</td>
                      <td>
                        {local.status === "COMPLETED"
                          ? <span className={row.serialNo ? "serial-filled" : "serial-empty"}>{row.serialNo || "(없음)"}</span>
                          : <input value={row.serialNo} onChange={(e) => setSerial(idx, e.target.value)} placeholder="시리얼 입력..." />
                        }
                      </td>
                      <td>
                        {local.status === "COMPLETED"
                          ? (row.status === "스크랩" ? <span className="badge badge-red">스크랩</span>
                            : row.status === "정상" ? <span className="badge badge-green">정상</span> : "-")
                          : (
                            <select value={row.status} onChange={(e) => setStatus(idx, e.target.value)}
                              style={{fontSize:11,padding:"3px 6px",border:"1.5px solid #e5e7eb",borderRadius:4}}>
                              <option value="">-</option>
                              <option value="정상">정상</option>
                              <option value="스크랩">스크랩</option>
                            </select>
                          )
                        }
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="actions">
              {local.status !== "COMPLETED" && (
                <>
                  <button className="btn btn-outline" onClick={() => save(false)}>임시저장</button>
                  <button className="btn btn-success" onClick={() => save(true)}>✓ 완료 저장</button>
                </>
              )}
              <button className="btn btn-primary" onClick={dlXlsx} disabled={busy}>
                {busy ? <span className="spinner"/> : null} xlsx
              </button>
              <button className="btn btn-outline" onClick={dlPdf} disabled={busy}>
                {busy ? <span className="spinner"/> : null} PDF
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ─── TradeDocManager ──────────────────────────────────────────────────────────
function TradeDocManager({ tradeDocs, persistTradeDocs }) {
  const [search, setSearch] = useState("");
  const [busy, setBusy] = useState(false);
  const [sel, setSel] = useState(new Set());

  const filtered = tradeDocs.filter((d) => {
    const s = search.toLowerCase();
    return !s || [d.poNumber, d.kitNo, d.quotationNumber, d.kitSerial].some((v) => (v || "").toLowerCase().includes(s));
  });

  const toggleSel = (id) =>
    setSel((p) => { const n = new Set(p); n.has(id) ? n.delete(id) : n.add(id); return n; });

  const del = (id) => {
    if (!confirm("삭제?")) return;
    persistTradeDocs(tradeDocs.filter((d) => d.id !== id));
  };

  const dlBatch = async () => {
    const docs = filtered.filter((d) => sel.has(d.id));
    if (!docs.length) { alert("선택된 항목이 없습니다."); return; }
    setBusy(true);
    try { await downloadTradeExcelBatch(docs); } catch (e) { alert(e.message); }
    finally { setBusy(false); }
  };

  const emptyCount = (d) => d.serialRows.filter((r) => !r.serialNo?.trim()).length;

  return (
    <div className="card">
      <div className="card-title">거래명세서 관리</div>
      <div className="search-bar">
        <input placeholder="PO / Kit No / 견적번호 검색..." value={search} onChange={(e) => setSearch(e.target.value)} />
        <span style={{fontSize:12,color:"#6b7280"}}>{filtered.length}건</span>
        {sel.size > 0 && (
          <button className="btn btn-primary" onClick={dlBatch} disabled={busy}>
            {busy ? <span className="spinner"/> : null} 배치 xlsx ({sel.size}건)
          </button>
        )}
      </div>
      {filtered.length === 0 ? (
        <div className="empty-state"><p>거래명세서가 없습니다.</p></div>
      ) : (
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>
                  <input type="checkbox"
                    checked={sel.size === filtered.length && filtered.length > 0}
                    onChange={() => setSel(sel.size === filtered.length ? new Set() : new Set(filtered.map((d) => d.id)))} />
                </th>
                <th>견적번호</th><th>PO</th><th>Kit No</th><th>출하일</th>
                <th>시리얼</th><th>합계(USD)</th><th>상태</th><th></th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((d) => {
                const empty = emptyCount(d);
                return (
                  <tr key={d.id}>
                    <td className="td-center">
                      <input type="checkbox" checked={sel.has(d.id)} onChange={() => toggleSel(d.id)} />
                    </td>
                    <td style={{fontSize:12,color:"#2d73ba",fontWeight:700}}>{d.quotationNumber}</td>
                    <td>{d.poNumber}</td>
                    <td>{d.kitNo}</td>
                    <td>{d.docDate}</td>
                    <td className="td-center">
                      <span style={{color: empty === 0 ? "#3ab06b" : "#e07b2a", fontWeight:600, fontSize:11}}>
                        {d.serialRows.length - empty}/{d.serialRows.length}
                      </span>
                    </td>
                    <td className="td-num">${formatUSD(d.totalUSD)}</td>
                    <td>
                      {d.status === "COMPLETED"
                        ? <span className="status-completed">완료</span>
                        : <span className="status-draft">초안</span>}
                    </td>
                    <td>
                      <div style={{display:"flex",gap:4}}>
                        <button className="btn btn-primary" style={{padding:"3px 10px",fontSize:11}} onClick={async () => {
                          try { await downloadTradeExcel(d); } catch (e) { alert(e.message); }
                        }}>xlsx</button>
                        <button className="btn btn-outline" style={{padding:"3px 10px",fontSize:11}} onClick={async () => {
                          try { await downloadTradePdf(d); } catch (e) { alert(e.message); }
                        }}>PDF</button>
                        <button className="btn btn-danger" style={{padding:"3px 8px",fontSize:11}} onClick={() => del(d.id)}>✕</button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ─── KitManager ───────────────────────────────────────────────────────────────
function KitManager({ kits, persistKits }) {
  const [editKit, setEditKit] = useState(null);
  const [kitForm, setKitForm] = useState({ partNo: "", name: "" });
  const [partForm, setPartForm] = useState({
    partNo: "", description: "", qty: 1, isSideNozzle: false,
    cleaningPriceUSD: 0, yoRecoatingPriceUSD: 0, bsRecoatingPriceUSD: 0, lidRecoatingPriceUSD: 0, htPriceUSD: 0, lpPriceUSD: 0
  });

  const openNew  = () => { setKitForm({ partNo: "", name: "" }); setEditKit({ id: uid(), parts: [] }); };
  const openEdit = (kit) => { setKitForm({ partNo: kit.partNo, name: kit.name }); setEditKit({ ...kit, parts: kit.parts.map((p) => ({ ...p })) }); };

  const saveKit = () => {
    if (!kitForm.partNo.trim()) { alert("Kit No를 입력하세요 (예: 0247-06765)"); return; }
    const partNo = kitForm.partNo.trim();
    const kit = { ...editKit, id: `kit-${partNo}`, partNo, name: kitForm.name.trim() };
    const exists = kits.find((k) => k.id === kit.id || k.partNo === partNo);
    persistKits(exists ? kits.map((k) => (k.id === kit.id || k.partNo === partNo ? kit : k)) : [...kits, kit]);
    setEditKit(null);
  };

  const delKit    = (id) => { if (confirm("삭제?")) persistKits(kits.filter((k) => k.id !== id)); };
  const toggleDis = (id) => persistKits(kits.map((k) => k.id === id ? { ...k, disabled: !k.disabled } : k));

  const emptyPart = () => ({ id: uid(), partNo: "", description: "", qty: 1, isSideNozzle: false, cleaningPriceUSD: 0, yoRecoatingPriceUSD: 0, bsRecoatingPriceUSD: 0, lidRecoatingPriceUSD: 0, htPriceUSD: 0, lpPriceUSD: 0 });

  const applyColumnPaste = (startIdx, key, rawText) => {
    const lines = String(rawText || "").split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
    if (lines.length <= 1) return false;
    setEditKit((prev) => {
      const parts = [...prev.parts];
      while (parts.length < startIdx + lines.length) parts.push(emptyPart());
      lines.forEach((line, offset) => {
        const i = startIdx + offset;
        if (key === "qty") parts[i] = { ...parts[i], qty: Math.max(1, Number(line) || 1) };
        else if (["cleaningPriceUSD","yoRecoatingPriceUSD","bsRecoatingPriceUSD","lidRecoatingPriceUSD","htPriceUSD","lpPriceUSD"].includes(key))
          parts[i] = { ...parts[i], [key]: Number(line) || 0 };
        else parts[i] = { ...parts[i], [key]: line };
      });
      return { ...prev, parts };
    });
    return true;
  };

  const onPasteColumn = (e, idx, key) => {
    if (applyColumnPaste(idx, key, e.clipboardData?.getData("text") || "")) e.preventDefault();
  };

  const addPart = () => {
    if (!partForm.partNo.trim()) { alert("Part No를 입력하세요"); return; }
    setEditKit((p) => ({ ...p, parts: [...p.parts, { id: uid(), ...partForm }] }));
    setPartForm({ partNo: "", description: "", qty: 1, isSideNozzle: false, cleaningPriceUSD: 0, yoRecoatingPriceUSD: 0, bsRecoatingPriceUSD: 0, lidRecoatingPriceUSD: 0, htPriceUSD: 0, lpPriceUSD: 0 });
  };

  const removePart = (idx) => setEditKit((p) => ({ ...p, parts: p.parts.filter((_, i) => i !== idx) }));
  const updatePart = (idx, patch) =>
    setEditKit((p) => ({ ...p, parts: p.parts.map((pt, i) => i === idx ? { ...pt, ...patch } : pt) }));

  if (editKit) return (
    <div className="card">
      <div className="card-title">{editKit.partNo ? "키트 편집" : "새 키트 추가"}</div>
      <div className="form-row">
        <div className="form-group">
          <label>Kit No (0247-XXXXX)</label>
          <input style={{width:140}} value={kitForm.partNo} placeholder="0247-06765"
            onChange={(e) => setKitForm((p) => ({ ...p, partNo: e.target.value }))} />
        </div>
        <div className="form-group" style={{flex:1}}>
          <label>키트 이름</label>
          <input value={kitForm.name} placeholder="SERVICE KIT, SYM3Y..."
            onChange={(e) => setKitForm((p) => ({ ...p, name: e.target.value }))} />
        </div>
      </div>

      <div className="divider" />
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
        <div className="card-title" style={{marginBottom:0}}>부품 목록</div>
        <span style={{fontSize:11,color:"var(--text-light)"}}>각 컬럼 칸에서 Ctrl+V — Part No 칸에 Part No 열, Description 칸에 Description 열 붙여넣기 (행 자동 추가)</span>
      </div>

      <div className="table-wrap">
        <table className="items-table">
          <thead>
            <tr>
              <th>Part No</th><th>Description</th><th>Qty</th><th>Side</th>
              <th>세정가($)</th><th>YO-RC($)</th><th>BS-RC($)</th><th>LID($)</th><th>HT($)</th><th>LP($)</th><th></th>
            </tr>
          </thead>
          <tbody>
            {editKit.parts.map((pt, idx) => (
              <tr key={pt.id || idx} style={pt.isSideNozzle ? {background:"rgba(74,144,217,0.08)"} : {}}>
                <td><input value={pt.partNo} style={{width:108}} onChange={(e) => updatePart(idx, { partNo: e.target.value })} onPaste={(e) => onPasteColumn(e, idx, "partNo")} /></td>
                <td><input value={pt.description} style={{width:"100%"}} onChange={(e) => updatePart(idx, { description: e.target.value })} onPaste={(e) => onPasteColumn(e, idx, "description")} /></td>
                <td><input type="number" min="1" value={pt.qty} style={{width:40}} onChange={(e) => updatePart(idx, { qty: Number(e.target.value) })} onPaste={(e) => onPasteColumn(e, idx, "qty")} /></td>
                <td className="td-center">
                  <input type="checkbox" checked={!!pt.isSideNozzle} onChange={(e) => updatePart(idx, { isSideNozzle: e.target.checked })} title="Side Nozzle 파트 지정" />
                </td>
                <td><input type="number" min="0" step="0.01" value={pt.cleaningPriceUSD       ?? 0} style={{width:68}} onChange={(e) => updatePart(idx, { cleaningPriceUSD:       Number(e.target.value) })} onPaste={(e) => onPasteColumn(e, idx, "cleaningPriceUSD")} /></td>
                <td><input type="number" min="0" step="0.01" value={pt.yoRecoatingPriceUSD   ?? 0} style={{width:68}} onChange={(e) => updatePart(idx, { yoRecoatingPriceUSD:  Number(e.target.value) })} onPaste={(e) => onPasteColumn(e, idx, "yoRecoatingPriceUSD")} /></td>
                <td><input type="number" min="0" step="0.01" value={pt.bsRecoatingPriceUSD   ?? 0} style={{width:68}} onChange={(e) => updatePart(idx, { bsRecoatingPriceUSD:  Number(e.target.value) })} onPaste={(e) => onPasteColumn(e, idx, "bsRecoatingPriceUSD")} /></td>
                <td><input type="number" min="0" step="0.01" value={pt.lidRecoatingPriceUSD  ?? 0} style={{width:68}} onChange={(e) => updatePart(idx, { lidRecoatingPriceUSD: Number(e.target.value) })} onPaste={(e) => onPasteColumn(e, idx, "lidRecoatingPriceUSD")} /></td>
                <td><input type="number" min="0" step="0.01" value={pt.htPriceUSD ?? 0} style={{width:68}} onChange={(e) => updatePart(idx, { htPriceUSD: Number(e.target.value) })} onPaste={(e) => onPasteColumn(e, idx, "htPriceUSD")} /></td>
                <td><input type="number" min="0" step="0.01" value={pt.lpPriceUSD ?? 0} style={{width:68}} onChange={(e) => updatePart(idx, { lpPriceUSD: Number(e.target.value) })} onPaste={(e) => onPasteColumn(e, idx, "lpPriceUSD")} /></td>
                <td><button className="btn btn-danger" style={{padding:"3px 8px",fontSize:11}} onClick={() => removePart(idx)}>✕</button></td>
              </tr>
            ))}
            <tr style={{background:"rgba(74,144,217,0.04)"}}>
              <td><input value={partForm.partNo} style={{width:108}} placeholder="0041-90314" onChange={(e) => setPartForm((p) => ({ ...p, partNo: e.target.value }))} onPaste={(e) => onPasteColumn(e, editKit.parts.length, "partNo")} /></td>
              <td><input value={partForm.description} style={{width:"100%"}} placeholder="TSGD" onChange={(e) => setPartForm((p) => ({ ...p, description: e.target.value }))} onPaste={(e) => onPasteColumn(e, editKit.parts.length, "description")} /></td>
              <td><input type="number" min="1" value={partForm.qty} style={{width:40}} onChange={(e) => setPartForm((p) => ({ ...p, qty: Number(e.target.value) }))} onPaste={(e) => onPasteColumn(e, editKit.parts.length, "qty")} /></td>
              <td className="td-center">
                <input type="checkbox" checked={!!partForm.isSideNozzle} onChange={(e) => setPartForm((p) => ({ ...p, isSideNozzle: e.target.checked }))} title="Side Nozzle 파트 지정" />
              </td>
              <td><input type="number" min="0" step="0.01" value={partForm.cleaningPriceUSD}      style={{width:68}} onChange={(e) => setPartForm((p) => ({ ...p, cleaningPriceUSD:       Number(e.target.value) }))} onPaste={(e) => onPasteColumn(e, editKit.parts.length, "cleaningPriceUSD")} /></td>
              <td><input type="number" min="0" step="0.01" value={partForm.yoRecoatingPriceUSD}  style={{width:68}} onChange={(e) => setPartForm((p) => ({ ...p, yoRecoatingPriceUSD:  Number(e.target.value) }))} onPaste={(e) => onPasteColumn(e, editKit.parts.length, "yoRecoatingPriceUSD")} /></td>
              <td><input type="number" min="0" step="0.01" value={partForm.bsRecoatingPriceUSD}  style={{width:68}} onChange={(e) => setPartForm((p) => ({ ...p, bsRecoatingPriceUSD:  Number(e.target.value) }))} onPaste={(e) => onPasteColumn(e, editKit.parts.length, "bsRecoatingPriceUSD")} /></td>
              <td><input type="number" min="0" step="0.01" value={partForm.lidRecoatingPriceUSD} style={{width:68}} onChange={(e) => setPartForm((p) => ({ ...p, lidRecoatingPriceUSD: Number(e.target.value) }))} onPaste={(e) => onPasteColumn(e, editKit.parts.length, "lidRecoatingPriceUSD")} /></td>
              <td><input type="number" min="0" step="0.01" value={partForm.htPriceUSD} style={{width:68}} onChange={(e) => setPartForm((p) => ({ ...p, htPriceUSD: Number(e.target.value) }))} onPaste={(e) => onPasteColumn(e, editKit.parts.length, "htPriceUSD")} /></td>
              <td><input type="number" min="0" step="0.01" value={partForm.lpPriceUSD} style={{width:68}} onChange={(e) => setPartForm((p) => ({ ...p, lpPriceUSD: Number(e.target.value) }))} onPaste={(e) => onPasteColumn(e, editKit.parts.length, "lpPriceUSD")} /></td>
              <td><button className="btn btn-success" style={{padding:"3px 10px",fontSize:11}} onClick={addPart}>+ 추가</button></td>
            </tr>
          </tbody>
        </table>
      </div>

      <div className="actions">
        <button className="btn btn-primary" onClick={saveKit}>저장</button>
        <button className="btn btn-ghost" onClick={() => setEditKit(null)}>취소</button>
      </div>
    </div>
  );

  return (
    <div className="card">
      <div className="card-title">키트 프리셋 관리</div>
      <div className="actions" style={{marginBottom:12,marginTop:0}}>
        <button className="btn btn-primary" onClick={openNew}>+ 새 키트 추가</button>
      </div>
      {kits.length === 0 ? (
        <div className="empty-state"><p>등록된 키트가 없습니다.</p></div>
      ) : (
        <div className="table-wrap">
          <table>
            <thead><tr><th>Kit No</th><th>이름</th><th>부품 수</th><th>상태</th><th></th></tr></thead>
            <tbody>
              {kits.map((k) => (
                <tr key={k.id} style={k.disabled ? {opacity:0.5} : {}}>
                  <td style={{fontWeight:700,color:"#2d73ba"}}>{k.partNo}</td>
                  <td>{k.name}</td>
                  <td className="td-center">{(k.parts || []).length}개</td>
                  <td>{k.disabled ? <span className="badge badge-gray">비활성</span> : <span className="badge badge-green">활성</span>}</td>
                  <td>
                    <div style={{display:"flex",gap:4}}>
                      <button className="btn btn-outline" style={{padding:"3px 10px",fontSize:11}} onClick={() => openEdit(k)}>편집</button>
                      <button className="btn btn-ghost" style={{padding:"3px 10px",fontSize:11}} onClick={() => toggleDis(k.id)}>
                        {k.disabled ? "활성화" : "비활성"}
                      </button>
                      <button className="btn btn-danger" style={{padding:"3px 8px",fontSize:11}} onClick={() => delKit(k.id)}>✕</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ─── App Root ─────────────────────────────────────────────────────────────────
const DEFAULT_SETTINGS = { exchangeRate: 1350, icpmsPrice: 261, lpcPrice: 225 };

export default function App() {
  const [tab, setTab]             = useState(TABS.SMART);
  const [kits, setKits]           = useState([]);
  const [quotes, setQuotes]       = useState([]);
  const [tradeDocs, setTradeDocs] = useState([]);
  const [activeQuoteId, setActiveQuoteId] = useState(null);
  const [loading, setLoading] = useState(true);
  const [settings, setSettings]   = useState(DEFAULT_SETTINGS);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [settingsForm, setSettingsForm] = useState(DEFAULT_SETTINGS);
  const [settingsSaving, setSettingsSaving] = useState(false);

  const persistKits      = useCallback((v) => { setKits(v);      storeSetKits(v);      }, []);
  const persistQuotes    = useCallback((v) => { setQuotes(v);    storeSetQuotes(v);    }, []);
  const persistTradeDocs = useCallback((v) => { setTradeDocs(v); storeSetTradeDocs(v); }, []);

  const saveSettings = async () => {
    setSettingsSaving(true);
    try {
      await saveAddSettingsToSheet(settingsForm);
      setSettings(settingsForm);
    } catch (e) { alert("설정 저장 실패: " + e.message); }
    finally { setSettingsSaving(false); }
  };

  const loadFromSheets = useCallback(async () => {
    setLoading(true);
    try {
      const { kits: k, quotes: q, tradeDocs: t, quoteSeqMap, addSettings } = await pullFromSheets();
      setKits(k); setQuotes(q); setTradeDocs(t);
      if (quoteSeqMap) window.__syqSeqMap = quoteSeqMap;
      if (addSettings && Object.keys(addSettings).length > 0) {
        const merged = { ...DEFAULT_SETTINGS, ...addSettings };
        setSettings(merged);
        setSettingsForm(merged);
      }
    } catch (e) {
      alert("데이터 로드 실패: " + e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { loadFromSheets(); }, []);

  const goToQuote = (id) => {
    setActiveQuoteId(id);
    setTab(TABS.QUOTE);
  };

  const handleParsed = () => {}; // unused — SmartGenerator manages state internally

  const activeQuote = activeQuoteId ? (quotes.find((q) => q.id === activeQuoteId) || null) : null;
  const draftQuotes = quotes.filter((q) => q.status === "DRAFT");

  const nav = [
    { section: "견적",    tab: TABS.SMART,      Icon: Zap,           label: "스마트 생성" },
    { section: null,      tab: TABS.QUOTE,      Icon: FileText,      label: "견적서 작성" },
    { section: null,      tab: TABS.QUOTE_MGMT, Icon: FolderOpen,    label: "견적 리스트" },
    { section: null,      tab: TABS.SUMMARY,    Icon: BarChart2,     label: "요약" },
    { section: "거래명세", tab: TABS.TRADE,      Icon: ClipboardList, label: "거래명세서" },
    { section: null,      tab: TABS.TRADE_MGMT, Icon: Archive,       label: "명세서 관리" },
    { section: "설정",    tab: TABS.KIT_MGMT,   Icon: Settings2,     label: "0247 관리" }
  ];

  const tabInfo = {
    [TABS.SMART]:      { title: "스마트 생성",   sub: "고객 데이터 붙여넣기 → 자동 파싱 → 견적서 생성" },
    [TABS.QUOTE]:      { title: "견적서 작성",   sub: "SYM3 제품 견적서 작성 · 스크랩/부가조건 · xlsx·PDF 출력" },
    [TABS.QUOTE_MGMT]: { title: "견적 리스트",   sub: "완료 견적서 목록 · 거래명세서 생성 · 다운로드" },
    [TABS.SUMMARY]:    { title: "요약",          sub: "전체 견적 현황 요약" },
    [TABS.TRADE]:      { title: "거래명세서",    sub: "시리얼 입력 · 거래명세서 출력" },
    [TABS.TRADE_MGMT]: { title: "명세서 관리",   sub: "거래명세서 목록 · 배치 다운로드" },
    [TABS.KIT_MGMT]:   { title: "0247 관리",     sub: "0247 키트 부품 목록 · 단가 등록" }
  };

  const renderContent = () => {
    switch (tab) {
      case TABS.SMART:
        return <SmartGenerator kits={kits} quotes={quotes} persistQuotes={persistQuotes} onSendToQuote={goToQuote} settings={settings} />;

      case TABS.QUOTE:
        return (
          <QuoteBuilder
            key={activeQuoteId}
            quote={activeQuote}
            kits={kits}
            quotes={quotes}
            persistQuotes={persistQuotes}
            onComplete={() => setTab(TABS.QUOTE_MGMT)}
          />
        );

      case TABS.SUMMARY:
        return <SummaryTab quotes={quotes} persistQuotes={persistQuotes} onSelectQuote={goToQuote} />;

      case TABS.QUOTE_MGMT:
        return <QuoteManager quotes={quotes} persistQuotes={persistQuotes} />;

      case TABS.TRADE:
        return <TradeDocBuilder quotes={quotes} tradeDocs={tradeDocs} persistTradeDocs={persistTradeDocs} />;

      case TABS.TRADE_MGMT:
        return <TradeDocManager tradeDocs={tradeDocs} persistTradeDocs={persistTradeDocs} />;

      case TABS.KIT_MGMT:
        return <KitManager kits={kits} persistKits={persistKits} />;

      default:
        return null;
    }
  };

  return (
    <div id="root">
      <div className="sidebar" style={{display:"flex", flexDirection:"column"}}>
        <div className="sidebar-logo">
          <div className="sidebar-logo-mark">SQ</div>
          <div>
            <div className="sidebar-logo-text">SYQ</div>
            <div className="sidebar-logo-sub">SYM3 Quotation</div>
          </div>
        </div>
        <div style={{flex:1}}>
          {nav.map((item, idx) => {
            const prevSection = idx > 0 ? nav[idx - 1].section : null;
            const { Icon } = item;
            return (
              <React.Fragment key={item.tab}>
                {item.section && item.section !== prevSection && (
                  <div className="sidebar-section">{item.section}</div>
                )}
                <div
                  className={`sidebar-item ${tab === item.tab ? "active" : ""}`}
                  onClick={() => setTab(item.tab)}
                >
                  <Icon size={15} className="sidebar-icon" />
                  <span>{item.label}</span>
                </div>
              </React.Fragment>
            );
          })}
        </div>

        {/* ── 환경설정 패널 ── */}
        <div style={{borderTop:"1px solid rgba(255,255,255,0.08)"}}>
          <div
            onClick={() => { setSettingsOpen(o => !o); setSettingsForm(settings); }}
            style={{display:"flex", alignItems:"center", gap:8, padding:"10px 16px",
              cursor:"pointer", color:"#64748b", fontSize:12, userSelect:"none"}}
          >
            <Settings2 size={13}/>
            <span style={{flex:1}}>환경설정</span>
            <span style={{fontSize:10}}>{settingsOpen ? "▲" : "▼"}</span>
          </div>
          {settingsOpen && (
            <div style={{padding:"6px 14px 12px", borderTop:"1px solid rgba(255,255,255,0.06)"}}>
              {[
                { key:"exchangeRate", label:"환율 (KRW/USD)" },
                { key:"icpmsPrice",   label:"ICPMS 분석가 ($)" },
                { key:"lpcPrice",     label:"LPC 분석가 ($)" },
              ].map(({ key, label }) => (
                <div key={key} style={{marginBottom:7}}>
                  <div style={{fontSize:10, color:"#475569", marginBottom:3}}>{label}</div>
                  <input
                    type="number" step="0.01"
                    value={settingsForm[key] ?? ""}
                    onChange={e => setSettingsForm(p => ({ ...p, [key]: Number(e.target.value) }))}
                    style={{width:"100%", fontSize:12, padding:"4px 7px",
                      border:"1px solid rgba(255,255,255,0.12)", borderRadius:6,
                      background:"rgba(255,255,255,0.06)", color:"#e2e8f0", outline:"none"}}
                  />
                </div>
              ))}
              <button
                onClick={saveSettings}
                disabled={settingsSaving}
                style={{width:"100%", padding:"5px 0", fontSize:11, fontWeight:700,
                  borderRadius:6, background:"#2563eb", color:"white", border:"none",
                  cursor: settingsSaving ? "not-allowed" : "pointer", opacity: settingsSaving ? 0.7 : 1}}
              >
                {settingsSaving ? "저장 중…" : "저장"}
              </button>
            </div>
          )}
        </div>
      </div>

      <div className="main">
        <div className="topbar">
          <div className="topbar-left">
            <div className="topbar-title">{tabInfo[tab]?.title || tab}</div>
            <span className="topbar-sep">·</span>
            <div className="topbar-sub">{tabInfo[tab]?.sub}</div>
          </div>
          <div className="topbar-right">
            <div className="topbar-stat">
              완료 견적 <strong>{quotes.filter(q => q.status === "COMPLETED").length}</strong>
            </div>
            <div className="topbar-stat">
              완료 명세서 <strong>{tradeDocs.filter(d => d.status === "COMPLETED").length}</strong>
            </div>
            {loading
              ? <span style={{fontSize:11,color:"#94a3b8",display:"flex",alignItems:"center",gap:5}}>
                  <span style={{width:8,height:8,borderRadius:"50%",background:"#94a3b8",display:"inline-block",animation:"pulse 1s infinite"}}/>
                  불러오는 중…
                </span>
              : <button className="btn btn-ghost" style={{padding:"3px 10px",fontSize:11}}
                  onClick={loadFromSheets}>↻ 새로고침</button>
            }
          </div>
        </div>
        <div className="content">
          {renderContent()}
        </div>
      </div>
    </div>
  );
}

