import { useState, useMemo, useRef, useEffect } from "react";
import { PieChart, Pie, Cell, Sector, ResponsiveContainer, Treemap } from "recharts";
import { Upload, LayoutDashboard, Boxes, PieChart as PieIcon, TrendingUp, Search, ChevronLeft, ChevronRight, ArrowUpRight, ArrowDownRight, Clock, LayoutGrid, List as ListIcon, Sparkles, Wifi } from "lucide-react";
import * as XLSX from "xlsx";

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:opsz,wght@9..40,400;9..40,500;9..40,600;9..40,700&family=DM+Mono:wght@400;500&display=swap');
.pf{--bg:#eef1f6;--surface:#ffffff;--ink:#0a1830;--line:rgba(12,26,48,.09);--line2:rgba(12,26,48,.15);--txt:#101f33;--muted:#5a6b82;--faint:#94a3b8;--acc:#2b5fac;--acc-deep:#1a3c6e;--acc-bg:#eaf1fb;--gain:#16834a;--loss:#c8362a;display:flex;min-height:620px;height:100%;background:var(--bg);color:var(--txt);font-family:'DM Sans',system-ui,sans-serif;font-size:14px;position:relative;overflow:hidden;-webkit-font-smoothing:antialiased;}
.pf::before{content:"";position:absolute;top:-260px;left:220px;width:560px;height:560px;border-radius:50%;background:radial-gradient(circle,rgba(43,95,172,.10),transparent 62%);filter:blur(60px);pointer-events:none;}
.pf *{box-sizing:border-box;}
.disp{font-family:'DM Serif Display',Georgia,serif;letter-spacing:.005em;}
.mono{font-family:'DM Mono',ui-monospace,monospace;font-variant-numeric:tabular-nums;}
.lbl{text-transform:uppercase;letter-spacing:.14em;font-size:.57rem;font-weight:500;color:var(--muted);font-family:'DM Mono',monospace;}
.gain{color:var(--gain);}.loss{color:var(--loss);}
.dot{width:8px;height:8px;border-radius:50%;display:inline-block;flex:none;}
.chrome{background:linear-gradient(135deg,#3f72c0 0%,#2b5fac 50%,#1a3c6e 100%);-webkit-background-clip:text;background-clip:text;color:transparent;}
.recharts-sector{cursor:pointer;}
.glass{background:var(--surface);border:1px solid var(--line);border-radius:14px;box-shadow:0 1px 2px rgba(12,26,48,.05),0 14px 34px -24px rgba(12,26,48,.30);}
.side{width:222px;flex:none;background:linear-gradient(180deg,#12244a,#0a1830);border-right:1px solid rgba(255,255,255,.07);display:flex;flex-direction:column;padding:22px 0;z-index:2;}
.brand{padding:0 22px 22px;}.brand .lg{display:flex;align-items:center;gap:11px;}
.mark{width:32px;height:32px;border-radius:9px;background:linear-gradient(140deg,#3f72c0 0%,#2b5fac 52%,#1a3c6e 100%);display:flex;align-items:center;justify-content:center;color:#fff;font-weight:400;font-family:'DM Serif Display',serif;font-size:18px;box-shadow:0 3px 12px -2px rgba(43,95,172,.7);}
.nav{flex:1;display:flex;flex-direction:column;gap:3px;padding:0 12px;}
.navi{display:flex;align-items:center;gap:11px;padding:10px 13px;border-radius:9px;cursor:pointer;color:#93a6c0;font-size:.85rem;font-weight:500;border:none;background:none;text-align:left;width:100%;transition:.15s;font-family:'DM Sans';}
.navi:hover{color:#fff;background:rgba(255,255,255,.07);}
.navi.on{color:#fff;background:rgba(63,114,192,.26);box-shadow:inset 2px 0 0 #5a86c4;font-weight:600;}
.main{flex:1;display:flex;flex-direction:column;z-index:2;min-width:0;}
.topbar{display:flex;align-items:center;justify-content:space-between;gap:16px;padding:15px 28px;border-bottom:1px solid var(--line);background:rgba(238,241,246,.82);backdrop-filter:blur(12px);position:sticky;top:0;z-index:5;}
.content{overflow:auto;padding:24px 28px 50px;flex:1;}
.content::-webkit-scrollbar{width:9px;}.content::-webkit-scrollbar-thumb{background:var(--line2);border-radius:9px;}
.fade{animation:fade .45s ease;}@keyframes fade{from{opacity:0;transform:translateY(8px);}to{opacity:1;transform:none;}}
.btn-acc{background:linear-gradient(180deg,#3f72c0,#2b5fac);color:#fff;border:1px solid #27569b;font-weight:600;font-size:.82rem;padding:9px 16px;border-radius:9px;cursor:pointer;display:inline-flex;align-items:center;gap:8px;box-shadow:inset 0 1px 0 rgba(255,255,255,.3),0 6px 16px -8px rgba(43,95,172,.7);transition:.15s;font-family:'DM Sans';}
.btn-acc:hover{transform:translateY(-1px);box-shadow:inset 0 1px 0 rgba(255,255,255,.3),0 9px 22px -9px rgba(43,95,172,.85);}
.btn-ghost{background:var(--surface);color:var(--acc-deep);border:1px solid rgba(43,95,172,.4);box-shadow:0 1px 2px rgba(12,26,48,.05);font-weight:600;font-size:.82rem;padding:9px 16px;border-radius:9px;cursor:pointer;display:inline-flex;align-items:center;gap:8px;transition:.15s;font-family:'DM Sans';}
.btn-ghost:hover{background:var(--acc-bg);border-color:var(--acc);}
.kpi{padding:17px 19px;position:relative;overflow:hidden;background:linear-gradient(180deg,#ffffff,#f7fafe);}
.kpi-rise{opacity:0;transform:translateY(8px);animation:rise .6s cubic-bezier(.2,.7,.3,1) forwards;}@keyframes rise{to{opacity:1;transform:none;}}
.barbg{background:rgba(12,26,48,.09);border-radius:4px;overflow:hidden;height:6px;}
.pill{display:inline-flex;align-items:center;gap:4px;font-size:.74rem;font-weight:500;padding:3px 9px;border-radius:20px;font-family:'DM Mono';}
.tag{font-size:.58rem;font-weight:500;letter-spacing:.03em;padding:3px 8px;border-radius:6px;white-space:nowrap;font-family:'DM Mono';}
.search{display:flex;align-items:center;gap:8px;background:var(--surface);border:1px solid var(--line2);border-radius:9px;padding:8px 12px;width:240px;}
.search input{border:none;outline:none;background:none;font-family:'DM Sans';font-size:.84rem;width:100%;color:var(--txt);}
.search input::placeholder{color:var(--faint);}
.chip{font-size:.74rem;font-weight:500;border:1px solid var(--line2);background:var(--surface);padding:6px 13px;border-radius:20px;cursor:pointer;color:var(--muted);transition:.15s;font-family:'DM Sans';}
.chip:hover{border-color:var(--acc);color:var(--txt);}
.chip.on{background:var(--acc);color:#fff;border-color:var(--acc);font-weight:600;}
.seg{display:flex;background:#e2e8f1;border:1px solid var(--line);border-radius:9px;padding:3px;gap:2px;}
.seg button{border:none;background:none;color:var(--muted);padding:6px 11px;border-radius:7px;cursor:pointer;display:flex;align-items:center;gap:6px;font-size:.78rem;font-weight:500;font-family:'DM Sans';}
.seg button.on{background:var(--surface);color:var(--txt);box-shadow:0 1px 3px rgba(12,26,48,.16);}
.tile{padding:15px 16px;cursor:pointer;position:relative;overflow:hidden;transition:transform .18s,box-shadow .18s,border-color .18s;}
.tile::before{content:"";position:absolute;top:0;left:0;right:0;height:3px;background:var(--sw);transform:scaleX(0);transform-origin:left;transition:transform .25s;}
.tile:hover{transform:translateY(-4px);box-shadow:0 20px 38px -22px rgba(12,26,48,.36),0 0 0 1px var(--sw);border-color:transparent;}
.tile:hover::before{transform:scaleX(1);}
.expand{max-height:0;overflow:hidden;transition:max-height .3s ease;}
.tile.open .expand{max-height:150px;}
.row{display:grid;grid-template-columns:1.7fr 1.1fr .9fr .9fr .8fr;align-items:center;gap:10px;padding:11px 16px;border-bottom:1px solid var(--line);font-size:.83rem;transition:.12s;}
.row:hover{background:var(--acc-bg);}
.rowhead{font-size:.57rem;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);font-weight:500;font-family:'DM Mono';}
.rowhead span{cursor:pointer;}.rowhead span:hover{color:var(--acc);}
.pgbtn{border:1px solid var(--line2);background:var(--surface);border-radius:9px;padding:7px 9px;cursor:pointer;display:flex;color:var(--txt);}
.pgbtn:disabled{opacity:.3;cursor:default;}
.drop{position:absolute;inset:0;z-index:40;background:rgba(238,241,246,.86);backdrop-filter:blur(6px);display:flex;align-items:center;justify-content:center;border:2px dashed var(--acc);}
`;

const BUCKETS = ["Market-priced", "Accrual / Fixed", "Illiquid / Lagged", "Cash & Insurance"];
const BUCKET_COLOR = { "Market-priced": "#2b5fac", "Accrual / Fixed": "#1a3c6e", "Illiquid / Lagged": "#7aa0d4", "Cash & Insurance": "#9aa7b8", Other: "#c4cedb" };
const BUCKET_NOTE = { "Market-priced": "Daily-priced — expect day-to-day swings.", "Accrual / Fixed": "Accrues toward maturity — low volatility.", "Illiquid / Lagged": "Valued periodically — lock-ins likely.", "Cash & Insurance": "Liquidity buffer / protection.", Other: "Unmapped class." };
const TO_BUCKET = {
  "Mutual Fund": "Market-priced", Stock: "Market-priced", Equity: "Market-priced", ETF: "Market-priced", REIT: "Market-priced", InvIT: "Market-priced", Gold: "Market-priced", "International Equity": "Market-priced", "Fund of Funds": "Market-priced", "Arbitrage Fund": "Market-priced",
  Bond: "Accrual / Fixed", NCD: "Accrual / Fixed", "G-Sec": "Accrual / Fixed", "T-Bill": "Accrual / Fixed", FD: "Accrual / Fixed", RD: "Accrual / Fixed", PPF: "Accrual / Fixed", EPF: "Accrual / Fixed", NPS: "Accrual / Fixed", "Tax-free Bond": "Accrual / Fixed",
  AIF: "Illiquid / Lagged", PMS: "Illiquid / Lagged", "Unlisted Equity": "Illiquid / Lagged", "Private Credit": "Illiquid / Lagged", "Real Estate": "Illiquid / Lagged", "Structured Product": "Illiquid / Lagged",
  "Liquid Fund": "Cash & Insurance", Savings: "Cash & Insurance", ULIP: "Cash & Insurance",
};
const bucketOf = (c) => TO_BUCKET[c] || "Other";
const PALETTE = ["#1a3c6e","#234e8a","#2b5fac","#3f72c0","#5a86c4","#7aa0d4","#9bb6dc","#33445e","#4b5d77","#5a6b82","#8595a8","#9aa7b8","#16315f","#c4cedb"];
const catColor = (c, all) => PALETTE[Math.max(0, all.indexOf(c)) % PALETTE.length];

/* ---------- currency-aware formatting ---------- */
let CUR = { sym: "₹", code: "INR", locale: "en-IN", indian: true };
const money = (n) => CUR.sym + Math.round(n || 0).toLocaleString(CUR.locale);
const signed = (n) => (n >= 0 ? "+" : "−") + CUR.sym + Math.round(Math.abs(n || 0)).toLocaleString(CUR.locale);
const compact = (n) => { const a = Math.abs(n || 0);
  if (CUR.indian) { if (a >= 1e7) return CUR.sym + (n / 1e7).toFixed(2) + " Cr"; if (a >= 1e5) return CUR.sym + (n / 1e5).toFixed(2) + " L"; return money(n); }
  if (a >= 1e9) return CUR.sym + (n / 1e9).toFixed(2) + "B"; if (a >= 1e6) return CUR.sym + (n / 1e6).toFixed(2) + "M"; if (a >= 1e3) return CUR.sym + (n / 1e3).toFixed(1) + "K"; return money(n); };
const pct = (n) => (n >= 0 ? "+" : "−") + Math.abs(n).toFixed(1) + "%";

/* currency map for the backend's currency code (Live mode) */
const CMAP = {
  USD: { sym: "$", code: "USD", locale: "en-US", indian: false },
  EUR: { sym: "€", code: "EUR", locale: "en-IE", indian: false },
  GBP: { sym: "£", code: "GBP", locale: "en-GB", indian: false },
  INR: { sym: "₹", code: "INR", locale: "en-IN", indian: true },
  JPY: { sym: "¥", code: "JPY", locale: "ja-JP", indian: false },
  AUD: { sym: "A$", code: "AUD", locale: "en-AU", indian: false },
  CAD: { sym: "C$", code: "CAD", locale: "en-CA", indian: false },
};

/* ---------- SMART PARSER (offline / browser) ---------- */
const toNum = (v) => { if (v == null || v === "") return null; const n = Number(String(v).replace(/[₹$€£¥,%\s]/g, "")); return isNaN(n) ? null : n; };
const H = {
  name: ["name","company","instrument","scheme","security","fund","stock","holding","asset","scheme name","fund name","security name","instrument name"],
  symbol: ["ticker","symbol","code","isin"],
  category: ["industry","sector","category","type","asset class","class","instrument type"],
  qty: ["quantity","units","shares","qty"],
  price: ["current price","price","nav","ltp","market price","avg nav","current nav","average nav"],
  current: ["market value","current value","present value","value","current","mkt value"],
  invested: ["invested","cost","cost value","invested amount","buy value","principal","amount invested","book value","total cost","total transaction amount","cost of inv","cost of investment","invested value","purchase cost"],
  gain: ["gain (loss)","gain(loss)","gain / loss","gain/loss","gain loss","gain","p&l","pnl","profit","unrealized","unrealised","notional p/l","total p/l"],
  day: ["$ day change","day change","day chg","today","chg"],
};
const ALL_KEYS = Object.values(H).flat();
/* rows that are section headers / totals, and a debt-fund name sniffer for files with no asset-class column */
const SKIP_NAMES = new Set(["total","grand total","subtotal","sub total","redeemed funds","redeemed","total:"]);
const DEBT_RE = /low dur|ultra short|money mark|floating|float|medium dur|liquid|gilt|bond|debt|overnight|short term|corporate bond|psu/i;
const CCY = { USD:{sym:"$",code:"USD",locale:"en-US",indian:false}, EUR:{sym:"€",code:"EUR",locale:"en-IE",indian:false}, GBP:{sym:"£",code:"GBP",locale:"en-GB",indian:false}, INR:{sym:"₹",code:"INR",locale:"en-IN",indian:true}, JPY:{sym:"¥",code:"JPY",locale:"ja-JP",indian:false}, AUD:{sym:"A$",code:"AUD",locale:"en-AU",indian:false}, CAD:{sym:"C$",code:"CAD",locale:"en-CA",indian:false} };

function detectCurrency(grids) {
  for (const g of grids) for (const r of g.slice(0, 50)) for (const cell of r) { const v = String(cell).trim().toUpperCase(); if (CCY[v]) return CCY[v]; }
  return CCY.INR;
}
function parseSheet(grid, sheetName = "") {
  let hRow = -1, hScore = 1;
  grid.slice(0, 30).forEach((r, i) => { const score = r.map((c) => String(c).trim().toLowerCase()).filter((c) => ALL_KEYS.includes(c)).length; if (score > hScore) { hScore = score; hRow = i; } });
  if (hRow < 0) return null;
  const header = grid[hRow].map((c) => String(c).trim().toLowerCase());
  const col = (keys) => { for (let j = 0; j < header.length; j++) if (keys.includes(header[j])) return j; return -1; };
  const ci = { name: col(H.name), symbol: col(H.symbol), category: col(H.category), qty: col(H.qty), price: col(H.price), current: col(H.current), invested: col(H.invested), gain: col(H.gain), day: col(H.day) };
  if (ci.name < 0 && ci.symbol < 0) return null;
  const sectorCat = ci.category >= 0 && ["industry", "sector"].includes(header[ci.category]);
  let score = (ci.current >= 0 ? 2 : 0) + (ci.gain >= 0 ? 1 : 0) + (ci.invested >= 0 ? 1 : 0) + (ci.qty >= 0 && ci.price >= 0 ? 1 : 0);
  const sn = sheetName.toLowerCase();
  if (/current|holding|portfolio/.test(sn)) score += 2;          // real holdings sheet
  if (/transaction|historical|return|comparison|active sip|wise|principal|category|amc|holder|sheet\d|^nav$/.test(sn)) score -= 2;  // summary / ledger tabs
  const out = [];
  for (let i = hRow + 1; i < grid.length; i++) {
    const r = grid[i]; if (!r) continue;
    let name = ci.name >= 0 ? String(r[ci.name] ?? "").trim() : "";
    const sym = ci.symbol >= 0 ? String(r[ci.symbol] ?? "").trim() : "";
    if (!name || name.startsWith("#")) name = sym;          // ticker fallback when Company is #VALUE!
    if (!name || name.startsWith("#")) continue;             // drop error / section / blank rows
    if (SKIP_NAMES.has(name.toLowerCase())) continue;        // drop Total / Redeemed-Funds section rows
    const qty = ci.qty >= 0 ? toNum(r[ci.qty]) : null, price = ci.price >= 0 ? toNum(r[ci.price]) : null, gain = ci.gain >= 0 ? toNum(r[ci.gain]) : null, day = ci.day >= 0 ? toNum(r[ci.day]) : null;
    let current = ci.current >= 0 ? toNum(r[ci.current]) : null, invested = ci.invested >= 0 ? toNum(r[ci.invested]) : null;
    if (current == null && qty != null && price != null) current = qty * price;
    if (invested == null && current != null && gain != null) invested = current - gain;     // derive cost basis
    if (invested == null && qty != null && price != null && gain != null) invested = qty * price - gain;
    if (current == null && invested != null && gain != null) current = invested + gain;
    if (current == null && invested == null) continue;       // no money on this row → skip
    if (current == null) current = invested; if (invested == null) invested = current;
    if (Math.abs(current) < 1 && Math.abs(invested) < 1) continue;   // closed / redeemed (zero value) → skip
    let category, bucket;
    if (ci.category >= 0) { category = String(r[ci.category] ?? "").trim() || "Other"; bucket = sectorCat ? "Market-priced" : undefined; }
    else { const isDebt = DEBT_RE.test(name); category = isDebt ? "Debt Fund" : "Equity Fund"; bucket = isDebt ? "Accrual / Fixed" : "Market-priced"; }  // infer class from fund name
    out.push({ name, symbol: sym, category, invested, current, qty, day, bucket });
  }
  return out.length ? { rows: out, score } : null;
}
function parseWorkbook(wb) {
  const grids = wb.SheetNames.map((sn) => XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "" }));
  let best = null;
  wb.SheetNames.forEach((sn, gi) => { const r = parseSheet(grids[gi], sn); if (r && (!best || r.score > best.score || (r.score === best.score && r.rows.length > best.rows.length))) best = r; });
  if (!best) throw new Error("no-data");
  return { rows: best.rows.map((p, i) => ({ ...p, id: i })), currency: detectCurrency(grids) };
}

/* ---------- sample book (INR) ---------- */
const RAW = [
  ["Parag Parikh Flexi Cap","PPFCF","Mutual Fund",320000,418000],["Quant Small Cap","QSC","Mutual Fund",180000,251000],["Mirae Asset Large Cap","MALC","Mutual Fund",240000,268000],["HDFC Balanced Advantage","HBA","Mutual Fund",150000,161000],["Axis Midcap","AXMC","Mutual Fund",120000,109000],["SBI Bluechip","SBIBC","Mutual Fund",160000,178000],["Kotak Emerging Equity","KEE","Mutual Fund",95000,121000],["ICICI Pru Technology","IPTECH","Mutual Fund",80000,71000],
  ["Reliance Industries","RELIANCE","Stock",200000,236000],["Tata Consultancy Svcs","TCS","Stock",175000,168000],["HDFC Bank","HDFCBANK","Stock",160000,187000],["Infosys","INFY","Stock",140000,151000],["Tata Motors","TATAMOTORS","Stock",90000,132000],["Zomato","ZOMATO","Stock",70000,58000],["Bajaj Finance","BAJFINANCE","Stock",130000,142000],["ITC","ITC","Stock",110000,128000],
  ["Nippon Nifty 50 BeES","NIFTYBEES","ETF",220000,247000],["Motilal Nasdaq 100","MON100","ETF",130000,169000],["ICICI Gold ETF","GOLDIETF","ETF",90000,113000],["Embassy Office Parks","EMBASSY","REIT",110000,118000],["IndiGrid InvIT","INDIGRID","InvIT",95000,101000],["Apple Inc (US)","AAPL","International Equity",140000,172000],
  ["Sovereign Gold Bond 2031","SGB31","Gold",150000,192000],["GOI 7.18% GS 2033","GS2033","G-Sec",200000,209000],["L&T Corporate Bond","LTBOND","Bond",100000,103500],["Muthoot Finance NCD","MUTHOOTNCD","NCD",120000,127000],
  ["HDFC Bank Fixed Deposit","HDFCFD","FD",300000,318000],["Bajaj Finance FD","BAJFD","FD",150000,160500],["Public Provident Fund","PPF","PPF",250000,271000],["NPS Tier-1","NPS","NPS",180000,205000],
  ["Avendus Absolute Return","AVENDUS","AIF",500000,575000],["Marcellus CCP","MARCELLUS","PMS",750000,861000],["NSE Unlisted Shares","NSEUL","Unlisted Equity",200000,245000],["Nippon Liquid Fund","NIPLIQ","Liquid Fund",180000,184500],["HDFC Life ProGrowth","HDFCULIP","ULIP",120000,131000],["Savings — Operating","CASH","Savings",90000,90000],
].map(([name, symbol, category, invested, current], i) => ({ id: i, name, symbol, category, invested, current }));

function useCountUp(value, ms = 800) {
  const [v, setV] = useState(0); const ref = useRef(0);
  useEffect(() => { const s = ref.current, t0 = performance.now(); let raf;
    const tick = (t) => { const p = Math.min(1, (t - t0) / ms); const e = 1 - Math.pow(1 - p, 3); setV(s + (value - s) * e); if (p < 1) raf = requestAnimationFrame(tick); else ref.current = value; };
    raf = requestAnimationFrame(tick); return () => cancelAnimationFrame(raf); }, [value, ms]); return v;
}
const ActiveShape = (p) => { const { cx, cy, innerRadius, outerRadius, startAngle, endAngle, fill } = p; return <g><Sector cx={cx} cy={cy} innerRadius={innerRadius} outerRadius={outerRadius + 9} startAngle={startAngle} endAngle={endAngle} fill={fill} /></g>; };
const PAGE = 12;

export default function App() {
  const [holdings, setHoldings] = useState(RAW);
  const [tab, setTab] = useState("overview");
  const [bucket, setBucket] = useState("All");
  const [q, setQ] = useState("");
  const [sort, setSort] = useState({ key: "current", dir: -1 });
  const [page, setPage] = useState(0);
  const [active, setActive] = useState(null);
  const [hview, setHview] = useState("tiles");
  const [openId, setOpenId] = useState(null);
  const [err, setErr] = useState("");
  const [drag, setDrag] = useState(false);
  const [fileName, setFileName] = useState("Sample book · 36 positions");
  const [, force] = useState(0);
  const fileRef = useRef();
  const serverRef = useRef();

  const rows = useMemo(() => holdings.map((h) => {
    const bucket = h.bucket || bucketOf(h.category); const pl = h.current - h.invested; const size = Math.abs(h.current);
    const ret = h.invested ? (pl / Math.abs(h.invested)) * 100 : 0; const isShort = (h.qty != null && h.qty < 0) || h.current < 0;
    return { ...h, bucket, pl, size, ret, isShort }; }), [holdings]);
  const cats = useMemo(() => [...new Set(rows.map((r) => r.category))], [rows]);
  const totals = useMemo(() => { const inv = rows.reduce((s, r) => s + r.invested, 0), cur = rows.reduce((s, r) => s + r.current, 0), abs = rows.reduce((s, r) => s + r.size, 0); return { inv, cur, abs, pl: cur - inv, ret: inv ? ((cur - inv) / Math.abs(inv)) * 100 : 0 }; }, [rows]);
  const A = totals.abs || 1;

  const byBucket = useMemo(() => { const m = {};
    rows.forEach((r) => { m[r.bucket] = m[r.bucket] || { bucket: r.bucket, invested: 0, current: 0, size: 0, n: 0, cats: {} };
      m[r.bucket].invested += r.invested; m[r.bucket].current += r.current; m[r.bucket].size += r.size; m[r.bucket].n++;
      m[r.bucket].cats[r.category] = m[r.bucket].cats[r.category] || { category: r.category, invested: 0, current: 0, size: 0, n: 0 };
      m[r.bucket].cats[r.category].invested += r.invested; m[r.bucket].cats[r.category].current += r.current; m[r.bucket].cats[r.category].size += r.size; m[r.bucket].cats[r.category].n++; });
    return Object.values(m).map((b) => ({ ...b, pl: b.current - b.invested, ret: b.invested ? (b.pl / Math.abs(b.invested)) * 100 : 0, share: (b.size / A) * 100,
      cats: Object.values(b.cats).map((c) => ({ ...c, pl: c.current - c.invested, ret: c.invested ? (c.pl / Math.abs(c.invested)) * 100 : 0, share: (c.size / A) * 100, plShare: totals.pl ? (c.pl / totals.pl) * 100 : 0 })).sort((a, z) => z.size - a.size),
    })).sort((a, z) => z.size - a.size); }, [rows, totals, A]);

  const donutData = byBucket.map((b) => ({ name: b.bucket, value: b.size }));
  const activeIdx = active ? donutData.findIndex((d) => d.name === active) : -1;
  const centerB = active ? byBucket.find((b) => b.bucket === active) : null;
  const movers = useMemo(() => [...rows].sort((a, z) => z.ret - a.ret), [rows]);
  const spread = [...movers.slice(0, 5), ...movers.slice(-4)].filter((v, i, a) => a.findIndex((x) => x.id === v.id) === i);
  const maxAbs = Math.max(1, ...spread.map((m) => Math.abs(m.ret)));
  const nameColor = useMemo(() => { const m = {}; rows.forEach((r) => (m[r.name] = BUCKET_COLOR[r.bucket])); return m; }, [rows]);
  const treeData = useMemo(() => { const s = [...rows].sort((a, z) => z.size - a.size); const top = s.slice(0, 14).map((r) => ({ name: r.name, value: r.size })); const rest = s.slice(14).reduce((a, r) => a + r.size, 0); if (rest > 0) top.push({ name: "Other", value: rest });
    const max = Math.max(...top.map((t) => t.value), 1), tot = top.reduce((a, t) => a + t.value, 0) || 1;
    const ramp = ["#0e2347","#16315f","#1a3c6e","#234e8a","#2b5fac","#4d7cc0","#7aa0d4"];
    return top.map((t) => { const idx = Math.round((1 - t.value / max) * (ramp.length - 1)); return { ...t, pctw: (t.value / tot) * 100, tone: t.name === "Other" ? "#94a3b8" : ramp[idx] }; }); }, [rows]);
  const brief = useMemo(() => { if (!rows.length) return null; const best = movers[0], worst = movers[movers.length - 1]; const driver = [...byBucket].sort((a, z) => z.pl - a.pl)[0]; const top = [...rows].sort((a, z) => z.size - a.size)[0]; return { best, worst, driver, top, topW: (top.size / A) * 100, lead: byBucket[0] }; }, [rows, byBucket, movers, A]);
  const analytics = useMemo(() => {
    const winners = rows.filter((r) => r.pl > 0), losers = rows.filter((r) => r.pl < 0);
    const avg = (a) => (a.length ? a.reduce((s, r) => s + r.ret, 0) / a.length : 0);
    const byCur = [...rows].sort((a, z) => z.size - a.size); const top1 = byCur[0];
    const hhi = rows.reduce((s, r) => { const w = r.size / A; return s + w * w; }, 0) * 10000;
    return { winners, losers, winRate: rows.length ? (winners.length / rows.length) * 100 : 0, avgWin: avg(winners), avgLoss: avg(losers),
      top1, top1w: (top1.size / A) * 100, top5w: byCur.slice(0, 5).reduce((s, r) => s + r.size, 0) / A * 100,
      hhi, conc: hhi < 1200 ? "Well diversified" : hhi < 2000 ? "Moderately concentrated" : "Highly concentrated",
      maxPL: Math.max(1, ...byBucket.map((b) => Math.abs(b.pl))), gainers: movers.slice(0, 6), losersList: [...movers].slice(-6).reverse() };
  }, [rows, byBucket, movers, A]);

  const filtered = useMemo(() => { let a = bucket === "All" ? rows : rows.filter((r) => r.bucket === bucket); if (q.trim()) { const s = q.toLowerCase(); a = a.filter((r) => r.name.toLowerCase().includes(s) || r.symbol.toLowerCase().includes(s) || r.category.toLowerCase().includes(s)); } return a; }, [rows, bucket, q]);
  const sorted = useMemo(() => { const a = [...filtered]; a.sort((x, y) => { const xv = x[sort.key], yv = y[sort.key]; return typeof xv === "string" ? xv.localeCompare(yv) * sort.dir : (xv - yv) * sort.dir; }); return a; }, [filtered, sort]);
  const pages = Math.max(1, Math.ceil(sorted.length / PAGE)); const pg = Math.min(page, pages - 1); const slice = sorted.slice(pg * PAGE, pg * PAGE + PAGE);
  useEffect(() => setPage(0), [bucket, q, sort]);
  const cv = useCountUp(totals.cur), iv = useCountUp(totals.inv), pv = useCountUp(totals.pl);

  function handleFile(f) { if (!f) return; const r = new FileReader();
    r.onload = (ev) => { try { const wb = XLSX.read(ev.target.result, { type: "array" }); const { rows: parsed, currency } = parseWorkbook(wb); CUR = currency; force((n) => n + 1);
      setHoldings(parsed); setFileName(f.name + " · " + parsed.length + " positions" + (currency.code !== "INR" ? " · " + currency.code : "")); setErr(""); setBucket("All"); setTab("overview"); }
      catch (e2) { setErr(e2.message === "no-data" ? "Scanned every sheet but couldn't find a holdings table. It needs a header row with a Name or Ticker, plus a value column (Market Value / Current, or Quantity + Price)." : "Couldn't read that file — use a valid .xlsx / .csv."); } };
    r.readAsArrayBuffer(f); }

  /* ---- Live mode: send the file to the FastAPI backend for reconciled, live-priced numbers + XIRR ---- */
  async function loadFromServer(f) {
    if (!f) return;
    try {
      const fd = new FormData();
      fd.append("file", f);
      const res = await fetch("https://ledger-backend-dspk.onrender.com/api/portfolio?live=true", { method: "POST", body: fd });
      if (!res.ok) throw new Error((await res.json()).detail || "server error");
      const data = await res.json();
      CUR = CMAP[data.meta.currency] || CUR;
      force((n) => n + 1);
      setHoldings(data.holdings.map((h, i) => ({ ...h, id: i })));
      setFileName(`${f.name} · live${data.meta.xirr != null ? ` · XIRR ${data.meta.xirr}%` : ""}`);
      setErr(""); setBucket("All"); setTab("overview");
    } catch (e) {
      setErr("Server: " + e.message + " — is the backend running on :8000?");
    }
  }

  const NAV = [["overview", "Overview", LayoutDashboard], ["holdings", "Holdings", Boxes], ["classes", "Allocation", PieIcon], ["performance", "Performance", TrendingUp]];
  const arrow = (k) => (sort.key === k ? (sort.dir === -1 ? " ↓" : " ↑") : "");
  const TreeCell = (p) => { const { x, y, width, height, name, depth, pctw, tone, value } = p; if (depth === 0 || width <= 0 || height <= 0) return null; const label = name == null ? "" : String(name); const showName = width > 60 && height > 28; const showPct = width > 50 && height > 44; const showVal = width > 116 && height > 66;
    const ix = x + 1.5, iy = y + 1.5, iw = Math.max(0, width - 3), ih = Math.max(0, height - 3), r = Math.min(9, iw / 2, ih / 2);
    return (<g>
      <rect x={ix} y={iy} width={iw} height={ih} rx={r} ry={r} style={{ fill: tone || "#9aa7b8" }} />
      <path d={`M${ix} ${iy + r} Q${ix} ${iy} ${ix + r} ${iy} L${ix + iw - r} ${iy} Q${ix + iw} ${iy} ${ix + iw} ${iy + r} L${ix + iw} ${iy + ih * 0.46} L${ix} ${iy + ih * 0.46} Z`} fill="#ffffff" fillOpacity={0.09} />
      <rect x={ix + 0.5} y={iy + 0.5} width={Math.max(0, iw - 1)} height={Math.max(0, ih - 1)} rx={r} ry={r} fill="none" stroke="rgba(255,255,255,.16)" strokeWidth={1} />
      {showName && <text x={x + 12} y={y + 21} fill="#ffffff" fontSize={11.5} fontWeight={600} fontFamily="DM Sans">{label.length > 18 ? label.slice(0, 17) + "…" : label}</text>}
      {showPct && <text x={x + 12} y={y + 38} fill="rgba(255,255,255,.82)" fontSize={11} fontFamily="DM Mono">{pctw != null ? pctw.toFixed(1) + "%" : ""}</text>}
      {showVal && <text x={x + 12} y={y + 56} fill="rgba(255,255,255,.62)" fontSize={10} fontFamily="DM Mono">{compact(value)}</text>}
    </g>); };
  const moverRow = (m, neg) => (<div key={m.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 0", borderBottom: "1px solid var(--line)" }}>
    <span style={{ fontSize: ".8rem", fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 150 }}>{m.name}</span>
    <span style={{ display: "flex", gap: 12, alignItems: "baseline" }}><span className={`mono ${neg ? "loss" : "gain"}`} style={{ fontSize: ".72rem" }}>{signed(m.pl)}</span><span className={`mono ${neg ? "loss" : "gain"}`} style={{ fontSize: ".78rem", width: 52, textAlign: "right" }}>{pct(m.ret)}</span></span></div>);

  return (
    <div className="pf" onDragOver={(e) => { e.preventDefault(); setDrag(true); }} onDragLeave={(e) => { e.preventDefault(); setDrag(false); }} onDrop={(e) => { e.preventDefault(); setDrag(false); handleFile(e.dataTransfer.files?.[0]); }}>
      <style>{CSS}</style>
      {drag && <div className="drop"><div style={{ textAlign: "center" }}><Upload size={30} style={{ color: "var(--acc)" }} /><div className="disp" style={{ fontSize: 18, marginTop: 10 }}>Drop your Excel / CSV</div></div></div>}
      <aside className="side">
        <div className="brand"><div className="lg"><span className="mark">L</span><div><div className="disp" style={{ fontSize: 19, color: "#eef3fb", lineHeight: 1, letterSpacing: ".02em" }}>LEDGER</div><div className="lbl" style={{ color: "#7aa0d4", marginTop: 5, fontSize: ".5rem" }}>Wealth Intelligence</div></div></div></div>
        <nav className="nav">{NAV.map(([id, label, Icon]) => (<button key={id} className={`navi ${tab === id ? "on" : ""}`} onClick={() => setTab(id)}><Icon size={16} /> {label}</button>))}</nav>
        <div className="mono" style={{ color: "var(--faint)", fontSize: ".6rem", padding: "12px 24px 0", lineHeight: 1.5 }}>{fileName}</div>
      </aside>
      <div className="main">
        <div className="topbar">
          <div><div className="disp" style={{ fontSize: 15, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".05em" }}>{NAV.find((n) => n[0] === tab)[1]}</div><div className="lbl" style={{ marginTop: 3, fontSize: ".52rem" }}>{rows.length} positions · {cats.length} classes</div></div>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <button className="btn-acc" onClick={() => fileRef.current.click()}><Upload size={15} /> Upload Excel / CSV</button>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={(e) => handleFile(e.target.files?.[0])} />
            <button className="btn-ghost" onClick={() => serverRef.current.click()}><Wifi size={15} /> Live (server)</button>
            <input ref={serverRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={(e) => loadFromServer(e.target.files?.[0])} />
          </div>
        </div>
        <div className="content">
          {err && <div className="glass" style={{ borderColor: "var(--loss)", color: "var(--loss)", padding: "11px 15px", marginBottom: 16, fontSize: ".84rem" }}>{err}</div>}

          {tab === "overview" && (<div className="fade" key="ov">
            <section style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr 1fr 1fr", gap: 12 }}>
              {[{ l: "Net Portfolio Value", v: money(cv), big: 1, chrome: 1 }, { l: "Invested", v: money(iv) }, { l: "Net Gain / Loss", v: signed(pv), c: totals.pl >= 0 ? "gain" : "loss" }, { l: "Total Return", v: pct(totals.ret), c: totals.ret >= 0 ? "gain" : "loss" }].map((k, i) => (
                <div key={i} className="glass kpi kpi-rise" style={{ animationDelay: `${i * 80}ms` }}>{k.chrome && <div style={{ position: "absolute", right: -30, top: -30, width: 110, height: 110, borderRadius: "50%", background: "radial-gradient(circle,rgba(43,95,172,.12),transparent 65%)" }} />}
                  <div className="lbl">{k.l}</div><div className={`disp ${k.chrome ? "chrome" : ""} ${k.c || ""}`} style={{ fontWeight: 700, fontSize: k.big ? 27 : 22, marginTop: 7, lineHeight: 1 }}>{k.v}</div></div>))}
            </section>
            {brief && (<div className="glass" style={{ marginTop: 14, padding: "14px 18px", display: "flex", gap: 13, alignItems: "center" }}><Sparkles size={17} style={{ color: "var(--acc)", flex: "none" }} />
              <p style={{ fontSize: ".9rem", lineHeight: 1.55, margin: 0 }}><b>{brief.lead.bucket}</b> leads at <b style={{ color: BUCKET_COLOR[brief.lead.bucket] }}>{brief.lead.share.toFixed(0)}%</b>; <b>{brief.driver.bucket}</b> added the most value <b className={brief.driver.pl >= 0 ? "gain" : "loss"}>({compact(brief.driver.pl)})</b>. Strongest <b>{brief.best.name}</b> <b className="gain">{pct(brief.best.ret)}</b>, weakest <b>{brief.worst.name}</b> <b className="loss">{pct(brief.worst.ret)}</b>. Largest <b>{brief.top.name}</b> is <b>{brief.topW.toFixed(0)}%</b> of exposure.</p></div>)}
            <section style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginTop: 14 }}>
              <div className="glass" style={{ padding: "16px 18px" }}><div className="lbl" style={{ marginBottom: 4 }}>Allocation by Behaviour · click to filter</div>
                <div style={{ position: "relative", height: 252 }}><ResponsiveContainer width="100%" height={252} minWidth={0}><PieChart>
                  <Pie data={donutData} dataKey="value" innerRadius={82} outerRadius={114} paddingAngle={2} stroke="none" isAnimationActive={false} activeIndex={activeIdx >= 0 ? activeIdx : undefined} activeShape={ActiveShape} onMouseEnter={(_, i) => setActive(donutData[i].name)} onMouseLeave={() => setActive(null)} onClick={(_, i) => { setBucket(donutData[i].name); setTab("holdings"); }}>{donutData.map((d) => <Cell key={d.name} fill={BUCKET_COLOR[d.name]} />)}</Pie>
                </PieChart></ResponsiveContainer>
                  <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", pointerEvents: "none" }}><div className="lbl" style={{ fontSize: ".52rem", textAlign: "center", maxWidth: 110 }}>{centerB ? centerB.bucket : "Exposure"}</div><div className="disp chrome" style={{ fontSize: 23, fontWeight: 700 }}>{compact(centerB ? centerB.size : totals.abs)}</div><div className="mono" style={{ fontSize: ".7rem", color: "var(--muted)" }}>{centerB ? centerB.share.toFixed(0) + "%" : rows.length + " held"}</div></div></div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: "5px 16px", marginTop: 8 }}>{byBucket.map((b) => (<div key={b.bucket} onMouseEnter={() => setActive(b.bucket)} onMouseLeave={() => setActive(null)} onClick={() => { setBucket(b.bucket); setTab("holdings"); }} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: ".74rem", cursor: "pointer", opacity: active && active !== b.bucket ? .4 : 1, transition: ".15s" }}><span className="dot" style={{ background: BUCKET_COLOR[b.bucket] }} />{b.bucket}<span className="mono" style={{ color: "var(--muted)" }}>{b.share.toFixed(0)}%</span></div>))}</div>
              </div>
              <div className="glass" style={{ padding: "16px 18px" }}><div className="lbl" style={{ marginBottom: 10 }}>Return Spread — Best &amp; Worst</div><div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {spread.map((m) => { const w = (Math.abs(m.ret) / maxAbs) * 50; return (<div key={m.id} style={{ display: "grid", gridTemplateColumns: "140px 1fr 52px", alignItems: "center", gap: 10 }}><span style={{ fontSize: ".78rem", fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{m.name}</span><div style={{ position: "relative", height: 15 }}><div style={{ position: "absolute", left: "50%", top: 0, bottom: 0, width: 1, background: "var(--line2)" }} /><div style={{ position: "absolute", top: 4, height: 7, borderRadius: 3, background: m.ret >= 0 ? "var(--gain)" : "var(--loss)", left: m.ret >= 0 ? "50%" : `${50 - w}%`, width: w + "%" }} /></div><span className={`mono ${m.ret >= 0 ? "gain" : "loss"}`} style={{ fontSize: ".78rem", textAlign: "right" }}>{pct(m.ret)}</span></div>); })}
              </div></div>
            </section>
            <div className="glass" style={{ marginTop: 14, padding: "16px 18px" }}><div className="lbl" style={{ marginBottom: 10 }}>Holdings Map — sized &amp; shaded by exposure</div><ResponsiveContainer width="100%" height={232} minWidth={0}><Treemap data={treeData} dataKey="value" content={<TreeCell />} isAnimationActive={false} /></ResponsiveContainer></div>
          </div>)}

          {tab === "holdings" && (<div className="fade" key="hd">
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12, marginBottom: 14 }}><div className="search"><Search size={15} style={{ color: "var(--muted)" }} /><input placeholder="Search name, ticker, class…" value={q} onChange={(e) => setQ(e.target.value)} /></div><div className="seg"><button className={hview === "tiles" ? "on" : ""} onClick={() => setHview("tiles")}><LayoutGrid size={14} /> Tiles</button><button className={hview === "list" ? "on" : ""} onClick={() => setHview("list")}><ListIcon size={14} /> List</button></div></div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 16 }}><button className={`chip ${bucket === "All" ? "on" : ""}`} onClick={() => setBucket("All")}>All · {rows.length}</button>{BUCKETS.filter((b) => rows.some((r) => r.bucket === b)).map((b) => (<button key={b} className={`chip ${bucket === b ? "on" : ""}`} onClick={() => setBucket(b)}>{b} · {rows.filter((r) => r.bucket === b).length}</button>))}</div>
            {hview === "tiles" && (<div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(255px,1fr))", gap: 13 }}>
              {slice.map((h) => { const w = (h.size / A) * 100; const ps = totals.pl ? (h.pl / totals.pl) * 100 : 0; const open = openId === h.id; return (
                <div key={h.id} className={`glass tile ${open ? "open" : ""}`} style={{ "--sw": BUCKET_COLOR[h.bucket] }} onClick={() => setOpenId(open ? null : h.id)}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 6 }}><span style={{ display: "flex", gap: 5, alignItems: "center", minWidth: 0 }}><span className="tag" style={{ background: BUCKET_COLOR[h.bucket] + "26", color: BUCKET_COLOR[h.bucket], overflow: "hidden", textOverflow: "ellipsis" }}>{h.category}</span>{h.isShort && <span className="tag" style={{ background: "rgba(200,54,42,.14)", color: "var(--loss)" }}>SHORT</span>}</span>{h.symbol && <span className="mono" style={{ fontSize: ".64rem", color: "var(--faint)" }}>{h.symbol}</span>}</div>
                  <div className="disp" style={{ fontSize: 14.5, fontWeight: 600, margin: "9px 0 9px", lineHeight: 1.25 }}>{h.name}</div>
                  <div className="mono" style={{ fontSize: 19, fontWeight: 700 }}>{money(h.current)}</div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", margin: "9px 0 8px" }}><span className="mono" style={{ fontSize: ".7rem", color: "var(--muted)" }}>cost {compact(h.invested)}</span><span className={`pill ${h.pl >= 0 ? "gain" : "loss"}`} style={{ background: h.pl >= 0 ? "rgba(22,131,74,.12)" : "rgba(200,54,42,.12)" }}>{h.pl >= 0 ? <ArrowUpRight size={12} /> : <ArrowDownRight size={12} />}{pct(h.ret)}</span></div>
                  <div className="barbg"><div style={{ width: Math.max(2, w) + "%", height: "100%", background: BUCKET_COLOR[h.bucket] }} /></div>
                  <div className="mono" style={{ fontSize: ".6rem", color: "var(--faint)", marginTop: 4 }}>{w.toFixed(1)}% of exposure</div>
                  <div className="expand"><div style={{ borderTop: "1px solid var(--line)", marginTop: 10, paddingTop: 10, fontSize: ".74rem", color: "var(--muted)", lineHeight: 1.7 }}>
                    <div style={{ display: "flex", justifyContent: "space-between" }}><span>Gain / loss</span><span className={`mono ${h.pl >= 0 ? "gain" : "loss"}`}>{signed(h.pl)}</span></div>
                    {h.day != null && <div style={{ display: "flex", justifyContent: "space-between" }}><span>Today</span><span className={`mono ${h.day >= 0 ? "gain" : "loss"}`}>{signed(h.day)}</span></div>}
                    <div style={{ display: "flex", justifyContent: "space-between" }}><span>Share of net gains</span><span className="mono">{ps >= 0 ? ps.toFixed(0) + "%" : "drag"}</span></div>
                    <div style={{ marginTop: 5, color: BUCKET_COLOR[h.bucket] }}>{BUCKET_NOTE[h.bucket]}</div>
                  </div></div>
                </div>); })}
            </div>)}
            {hview === "list" && (<div className="glass" style={{ overflow: "hidden" }}><div className="row rowhead">{[["name", "Instrument"], ["bucket", "Behaviour"], ["invested", "Cost"], ["current", "Value"], ["ret", "Return"]].map(([k, l]) => (<span key={k} onClick={() => setSort((s) => ({ key: k, dir: s.key === k ? -s.dir : -1 }))} style={{ textAlign: ["invested", "current", "ret"].includes(k) ? "right" : "left" }}>{l}{arrow(k)}</span>))}</div>
              {slice.map((h) => (<div className="row" key={h.id}><span><span style={{ fontWeight: 500 }}>{h.name}</span>{h.symbol && <span className="mono" style={{ color: "var(--faint)", fontSize: ".68rem", marginLeft: 7 }}>{h.symbol}</span>}{h.isShort && <span className="tag" style={{ background: "rgba(200,54,42,.14)", color: "var(--loss)", marginLeft: 6 }}>SHORT</span>}</span><span><span className="tag" style={{ background: BUCKET_COLOR[h.bucket] + "26", color: BUCKET_COLOR[h.bucket] }}>{h.bucket}</span></span><span className="mono" style={{ textAlign: "right" }}>{money(h.invested)}</span><span className="mono" style={{ textAlign: "right", fontWeight: 600 }}>{money(h.current)}</span><span className={`mono ${h.ret >= 0 ? "gain" : "loss"}`} style={{ textAlign: "right" }}>{pct(h.ret)}</span></div>))}</div>)}
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 14 }}><span className="mono" style={{ fontSize: ".72rem", color: "var(--muted)" }}>{sorted.length === 0 ? "No matches" : `${pg * PAGE + 1}–${Math.min((pg + 1) * PAGE, sorted.length)} of ${sorted.length}`}</span><div style={{ display: "flex", alignItems: "center", gap: 10 }}><button className="pgbtn" disabled={pg === 0} onClick={() => setPage(pg - 1)}><ChevronLeft size={15} /></button><span className="mono" style={{ fontSize: ".74rem" }}>{pg + 1} / {pages}</span><button className="pgbtn" disabled={pg >= pages - 1} onClick={() => setPage(pg + 1)}><ChevronRight size={15} /></button></div></div>
          </div>)}

          {tab === "classes" && (<div className="fade" key="al" style={{ display: "flex", flexDirection: "column", gap: 14 }}>{byBucket.map((b) => (
            <div key={b.bucket} className="glass" style={{ padding: "16px 20px" }}><div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", borderBottom: "1px solid var(--line)", paddingBottom: 11, marginBottom: 13 }}><span style={{ display: "flex", alignItems: "center", gap: 9, fontSize: ".94rem", fontWeight: 600 }}><span className="dot" style={{ width: 11, height: 11, background: BUCKET_COLOR[b.bucket] }} />{b.bucket}<span className="mono" style={{ color: "var(--muted)", fontSize: ".68rem", fontWeight: 400 }}>{b.n} · {b.share.toFixed(0)}%</span></span><span style={{ display: "flex", gap: 18, alignItems: "baseline" }}><span className="mono" style={{ fontWeight: 600 }}>{compact(b.current)}</span><span className={`mono ${b.pl >= 0 ? "gain" : "loss"}`} style={{ width: 58, textAlign: "right" }}>{pct(b.ret)}</span></span></div>
              {b.cats.map((c) => (<div key={c.category} style={{ marginBottom: 11 }}><div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 4 }}><span style={{ display: "flex", alignItems: "center", gap: 7, fontSize: ".84rem" }}><span className="dot" style={{ background: catColor(c.category, cats) }} />{c.category}<span className="mono" style={{ color: "var(--faint)", fontSize: ".66rem" }}>×{c.n}</span></span><span style={{ display: "flex", gap: 14, alignItems: "baseline" }}><span className="mono" style={{ fontSize: ".8rem" }}>{compact(c.current)}</span><span className={`mono ${c.ret >= 0 ? "gain" : "loss"}`} style={{ fontSize: ".76rem", width: 52, textAlign: "right" }}>{pct(c.ret)}</span></span></div><div className="barbg"><div style={{ width: Math.max(2, c.share) + "%", height: "100%", background: catColor(c.category, cats) }} /></div><div className="mono" style={{ fontSize: ".6rem", color: "var(--faint)", marginTop: 3 }}>{c.share.toFixed(1)}% of exposure · {c.plShare >= 0 ? c.plShare.toFixed(0) + "% of net gains" : "drag on returns"}</div></div>))}
            </div>))}</div>)}

          {tab === "performance" && (<div className="fade" key="pf">
            <section style={{ display: "grid", gridTemplateColumns: "repeat(6,1fr)", gap: 10 }}>{[{ l: "Net Return", v: pct(totals.ret), c: totals.ret >= 0 ? "gain" : "loss" }, { l: "Win Rate", v: analytics.winRate.toFixed(0) + "%", c: "" }, { l: "Winners", v: analytics.winners.length, c: "gain" }, { l: "Losers", v: analytics.losers.length, c: "loss" }, { l: "Avg Winner", v: pct(analytics.avgWin), c: "gain" }, { l: "Avg Loser", v: pct(analytics.avgLoss), c: "loss" }].map((k, i) => (<div key={i} className="glass" style={{ padding: "13px 14px" }}><div className="lbl" style={{ fontSize: ".5rem" }}>{k.l}</div><div className={`disp ${k.c}`} style={{ fontSize: 18, fontWeight: 700, marginTop: 5 }}>{k.v}</div></div>))}</section>
            <div className="glass" style={{ marginTop: 14, padding: "16px 18px" }}><div className="lbl" style={{ marginBottom: 12 }}>Contribution to Net P&amp;L — by behaviour</div>{byBucket.map((b) => { const w = (Math.abs(b.pl) / analytics.maxPL) * 50; return (<div key={b.bucket} style={{ display: "grid", gridTemplateColumns: "160px 1fr 96px", alignItems: "center", gap: 12, marginBottom: 9 }}><span style={{ display: "flex", alignItems: "center", gap: 7, fontSize: ".8rem" }}><span className="dot" style={{ background: BUCKET_COLOR[b.bucket] }} />{b.bucket}</span><div style={{ position: "relative", height: 16 }}><div style={{ position: "absolute", left: "50%", top: 0, bottom: 0, width: 1, background: "var(--line2)" }} /><div style={{ position: "absolute", top: 4, height: 8, borderRadius: 3, background: b.pl >= 0 ? "var(--gain)" : "var(--loss)", left: b.pl >= 0 ? "50%" : `${50 - w}%`, width: w + "%" }} /></div><span className={`mono ${b.pl >= 0 ? "gain" : "loss"}`} style={{ fontSize: ".78rem", textAlign: "right" }}>{signed(b.pl)}</span></div>); })}</div>
            <section style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginTop: 14 }}><div className="glass" style={{ padding: "16px 18px" }}><div className="lbl gain" style={{ marginBottom: 8 }}>Top Gainers</div>{analytics.gainers.map((m) => moverRow(m, false))}</div><div className="glass" style={{ padding: "16px 18px" }}><div className="lbl loss" style={{ marginBottom: 8 }}>Top Losers</div>{analytics.losersList.map((m) => moverRow(m, true))}</div></section>
            <section style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginTop: 14 }}>
              <div className="glass" style={{ padding: "16px 18px" }}><div className="lbl" style={{ marginBottom: 10 }}>Concentration &amp; Diversification</div><div className="disp" style={{ fontSize: 17, fontWeight: 700, color: analytics.hhi < 1200 ? "var(--gain)" : analytics.hhi < 2000 ? "var(--acc)" : "var(--loss)" }}>{analytics.conc}</div><div className="mono" style={{ fontSize: ".68rem", color: "var(--muted)", marginTop: 2, marginBottom: 12 }}>HHI {analytics.hhi.toFixed(0)} · {rows.length} holdings · {cats.length} classes</div><div style={{ fontSize: ".76rem", color: "var(--muted)", display: "flex", justifyContent: "space-between", marginBottom: 5 }}><span>Largest — {analytics.top1?.name}</span><span className="mono">{analytics.top1w.toFixed(0)}%</span></div><div className="barbg" style={{ marginBottom: 10 }}><div style={{ width: analytics.top1w + "%", height: "100%", background: "var(--acc)" }} /></div><div style={{ fontSize: ".76rem", color: "var(--muted)", display: "flex", justifyContent: "space-between", marginBottom: 5 }}><span>Top 5 holdings</span><span className="mono">{analytics.top5w.toFixed(0)}%</span></div><div className="barbg"><div style={{ width: analytics.top5w + "%", height: "100%", background: "var(--acc-deep)" }} /></div></div>
              <div className="glass" style={{ padding: "16px 18px", borderStyle: "dashed" }}><div style={{ display: "flex", alignItems: "center", gap: 9, marginBottom: 8 }}><Clock size={16} style={{ color: "var(--acc)" }} /><span className="lbl">Live Layer · Phase 2</span></div><p style={{ fontSize: ".82rem", color: "var(--muted)", lineHeight: 1.6, margin: 0 }}>Use the <b>Live (server)</b> button to pull reconciled, live-priced numbers and XIRR from the backend. Value-over-time curve and benchmark overlay come next.</p></div>
            </section>
          </div>)}
        </div>
      </div>
    </div>
  );
}