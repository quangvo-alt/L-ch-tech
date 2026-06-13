// ============================================================
//  TECH SHIFT DASHBOARD v7
//  Layout tự co theo nội dung · Chữ to · Không cắt
//  Tab: "Filter" | Cột G = Tên | H = Dept | I = Shift
//  Shift: C1 C2 C3 ME NP HO OFF
//  Deploy: Execute as Me · Who has access: Anyone
// ============================================================

function doGet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Filter");

  const dateCell = sheet.getRange("I1").getValue();
  const tz       = ss.getSpreadsheetTimeZone();
  const today    = dateCell ? new Date(dateCell) : new Date();
  const dateStr  = Utilities.formatDate(today, tz, "MMMM dd, yyyy");
  const dayName  = Utilities.formatDate(today, tz, "EEEE");

  // Convert VN fixed times (UTC+7) to any local timezone dynamically
  const now         = new Date();
  const tzOffsetStr = Utilities.formatDate(now, tz, "Z"); // e.g. "-0600", "-0500", "-0400"
  const tzAbbr      = Utilities.formatDate(now, tz, "z"); // e.g. "CST", "CDT", "EST", "EDT"
  const sign        = tzOffsetStr[0] === "-" ? -1 : 1;
  const hh          = parseInt(tzOffsetStr.slice(1,3), 10);
  const mm          = parseInt(tzOffsetStr.slice(3,5), 10);
  const localOffset = sign * (hh + mm / 60); // e.g. -6, -5, -4
  const utcPart     = "UTC" + (localOffset >= 0 ? "+" : "") + localOffset;
  const tzLabel     = tzAbbr + " (" + utcPart + ")";

  function vnToLocal(vnHour) {
    // VN = UTC+7, so UTC = VN - 7; local = UTC + localOffset
    return ((vnHour - 7 + localOffset) % 24 + 24) % 24;
  }
  function fmtHour(h) {
    const period = h < 12 ? "AM" : "PM";
    const h12    = h % 12 === 0 ? 12 : h % 12;
    return h12 + ":00 " + period;
  }

  // VN fixed: Shift1 20→28(=4), Shift2 4→16(noon), Shift3 22→29(=5)
  const shiftTimesRaw = {
    C1: { s: vnToLocal(20),   e: vnToLocal(28 % 24) },
    C2: { s: vnToLocal(4),    e: vnToLocal(16) },
    C3: { s: vnToLocal(22),   e: vnToLocal(29 % 24) },
  };

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return HtmlService.createHtmlOutput("<p>Không có dữ liệu</p>");
  const raw = sheet.getRange(3, 7, lastRow - 2, 3).getValues();

  const shiftOrder = ["C1","C2","C3","ME","NP","HO","OFF"];
  const grouped = {};
  shiftOrder.forEach(s => grouped[s] = []);

  raw.forEach(row => {
    const name  = (row[0] || "").toString().trim();
    const dept  = (row[1] || "").toString().trim();
    const shift = (row[2] || "").toString().trim().toUpperCase();
    if (!name) return;
    if (grouped[shift] !== undefined) grouped[shift].push({ name, dept });
  });

  return HtmlService
    .createHtmlOutput(buildHtml(grouped, shiftOrder, dateStr, dayName, shiftTimesRaw, tzLabel))
    .setTitle("Tech Shift " + dateStr)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function buildHtml(grouped, shiftOrder, dateStr, dayName, shiftTimesRaw, tzLabel) {

  function fmtHour(h) {
    const period = h < 12 ? "AM" : "PM";
    const h12    = h % 12 === 0 ? 12 : h % 12;
    return h12 + ":00 " + period;
  }
  const shiftTimes = {
    C1: fmtHour(shiftTimesRaw.C1.s) + " – " + fmtHour(shiftTimesRaw.C1.e),
    C2: fmtHour(shiftTimesRaw.C2.s) + " – " + fmtHour(shiftTimesRaw.C2.e),
    C3: fmtHour(shiftTimesRaw.C3.s) + " – " + fmtHour(shiftTimesRaw.C3.e),
  };

  const meta = {
    C1:  { label:"Shift 1", time:shiftTimes.C1, grad:"#1d4ed8,#3b82f6", avBg:"#dbeafe", avTxt:"#1d4ed8", statClr:"#60a5fa", chipBg:"#1e1b4b", chipTxt:"#c7d2fe", chipBd:"#3730a3", lbBg:"#1d4ed8", lbTxt:"#eff6ff" },
    C2:  { label:"Shift 2", time:shiftTimes.C2, grad:"#b45309,#f97316", avBg:"#ffedd5", avTxt:"#b45309", statClr:"#fb923c", chipBg:"#431407", chipTxt:"#fed7aa", chipBd:"#9a3412", lbBg:"#c2410c", lbTxt:"#fff7ed" },
    C3:  { label:"Shift 3", time:shiftTimes.C3, grad:"#15803d,#22c55e", avBg:"#dcfce7", avTxt:"#15803d", statClr:"#4ade80", chipBg:"#052e16", chipTxt:"#bbf7d0", chipBd:"#166534", lbBg:"#15803d", lbTxt:"#f0fdf4" },
    ME:  { label:"Mac Energy — Nghỉ có công", time:"",               statClr:"#a78bfa", chipBg:"#2e1065", chipTxt:"#ddd6fe", chipBd:"#5b21b6", lbBg:"#4c1d95", lbTxt:"#ede9fe" },
    NP:  { label:"Nghỉ phép năm — Có công",  time:"",               statClr:"#fbbf24", chipBg:"#1c1407", chipTxt:"#fde68a", chipBd:"#92400e", lbBg:"#92400e", lbTxt:"#fef3c7" },
    HO:  { label:"Nghỉ lễ / Bù lễ — Có công", time:"",             statClr:"#f87171", chipBg:"#450a0a", chipTxt:"#fecaca", chipBd:"#7f1d1d", lbBg:"#991b1b", lbTxt:"#fee2e2" },
    OFF: { label:"Nghỉ",                     time:"",               statClr:"#94a3b8", chipBg:"#1e293b", chipTxt:"#f1f5f9", chipBd:"#475569", lbBg:"#334155", lbTxt:"#f1f5f9" },
  };

  function cleanName(raw) { return raw.replace(/\s*\([^)]*\)/g,"").trim(); }
  function cleanRole(raw) { return (raw || "").trim().replace(/[,;.]+$/, "").trim(); }
  function roleStyle(role) {
    const r = role.toLowerCase();
    if (r.includes("mgr"))                          return "background:#fef3c7;color:#92400e;border:1px solid #fcd34d;";   // Amber — Manager
    if (r.includes("sup") && !r.includes("build"))  return "background:#dbeafe;color:#1d4ed8;border:1px solid #93c5fd;";   // Blue — Tech Support
    if (r.includes("build"))                         return "background:#ffedd5;color:#c2410c;border:1px solid #fdba74;";   // Orange — Tech Build
    if (r.includes("lobby"))                         return "background:#ede9fe;color:#5b21b6;border:1px solid #c4b5fd;";   // Purple — Tech Lobby
    if (r.includes("ba"))                            return "background:#d1fae5;color:#065f46;border:1px solid #6ee7b7;";   // Green — BA
    return "background:#f1f5f9;color:#475569;border:1px solid #cbd5e1;";  // Gray — default
  }
  function initials(raw) {
    return cleanName(raw).split(/\s+/).slice(-2).map(w=>w[0].toUpperCase()).join("");
  }

  const totalWorking = ["C1","C2","C3"].reduce((s,k)=>s+(grouped[k]||[]).length,0);
  const totalAll     = shiftOrder.reduce((s,k)=>s+(grouped[k]||[]).length,0);

  // Top stat chips
  const statChipsHtml = shiftOrder.map(s => {
    const m = meta[s]; const n = (grouped[s]||[]).length;
    return `<div class="stat-chip"><span class="n" style="color:${m.statClr};">${n}</span><span class="l" style="color:${m.statClr};">${s}</span></div>`;
  }).join("");

  // Person row inside shift card
  function pRow(p, m) {
    const ini      = initials(p.name);
    const name     = cleanName(p.name);
    const role     = cleanRole(p.dept);   // column H = Function/Role
    const roleHtml = role ? `<span class="p-role" style="${roleStyle(role)}">${role}</span>` : "";
    return `<div class="p-row">
      <div class="p-av" style="background:${m.avBg};color:${m.avTxt};">${ini}</div>
      <span class="p-nm">${name}</span>${roleHtml}
    </div>`;
  }

  // Shift card C1/C2/C3
  function shiftCard(s) {
    const people = grouped[s]||[];
    const m = meta[s];
    const rows = people.map(p=>pRow(p,m)).join("");
    const id   = s === "C3" ? ' id="card-c3"' : '';
    return `<div${id} style="border-radius:10px;overflow:hidden;border:1px solid rgba(255,255,255,.08);display:flex;flex-direction:column;">
      <div class="card-hdr" style="background:linear-gradient(90deg,${m.grad});">
        <span class="shift-lbl">${s}</span>
        <span class="shift-sub">${m.label}</span>
        <span class="shift-time">${m.time} <span class="shift-tz">${tzLabel}</span></span>
        <span class="shift-cnt">${people.length}</span>
      </div>
      <div style="background:#fff;padding:3px 10px 6px;">${rows || '<span style="font-size:11px;color:#ccc;">—</span>'}</div>
    </div>`;
  }

  // Name chip for "other" col
  function chip(p, m) {
    const name = cleanName(p.name);
    const dept = p.dept && p.dept.toUpperCase() !== "TECH"
      ? ` <span style="font-size:9px;opacity:.7;">${p.dept}</span>` : "";
    return `<div class="oth-chip" style="background:${m.chipBg};color:${m.chipTxt};border-color:${m.chipBd};">${name}${dept}</div>`;
  }

  // Section inside "other" col
  function otherSec(s) {
    const people = grouped[s]||[]; const m = meta[s];
    const chips  = people.length
      ? people.map(p=>chip(p,m)).join("")
      : `<span style="font-size:10px;color:#334155;font-style:italic;">Không có</span>`;
    return `<div class="oth-sec">
      <div class="oth-lbl" style="background:${m.lbBg};color:${m.lbTxt};">
        ${s} — ${m.label}
        <span style="background:rgba(255,255,255,.2);border-radius:999px;padding:0 6px;">${people.length}</span>
      </div>
      <div class="oth-chips">${chips}</div>
    </div>`;
  }

  const otherHtml = ["ME","NP","HO","OFF"].map(s=>otherSec(s)).join("");

  return `<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Tech Shift ${dateStr}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"><\/script>
<style>
* { box-sizing:border-box; margin:0; padding:0; }
html { background:#0B1426; }
body {
  font-family:'Inter',sans-serif;
  background:#0B1426;
  padding-bottom:8px;
  position:relative;
}
body::before {
  content:''; position:fixed; inset:0; pointer-events:none; z-index:0;
  background-image:
    linear-gradient(rgba(99,179,237,.04) 1px,transparent 1px),
    linear-gradient(90deg,rgba(99,179,237,.04) 1px,transparent 1px);
  background-size:40px 40px;
}

/* ── MOBILE-FIRST (default) ── */
#topbar {
  position:relative; z-index:1;
  display:flex; align-items:center; justify-content:space-between;
  padding:6px 10px;
  background:linear-gradient(90deg,#0f2057,#1e3a8a);
  border-bottom:1px solid rgba(99,179,237,.2);
  margin-bottom:6px;
  gap:6px;
}
#topbar-brand { display:flex; align-items:center; gap:5px; flex-shrink:0; }
#topbar-date  { text-align:center; flex:1; }
#stats-row    { display:flex; align-items:center; gap:4px; flex-shrink:0; }
.stat-chip    {
  display:flex; flex-direction:column; align-items:center;
  background:rgba(255,255,255,.07); border:1px solid rgba(255,255,255,.12);
  border-radius:8px; padding:3px 7px; min-width:30px;
}
.stat-chip .n { font-size:14px; font-weight:900; line-height:1; }
.stat-chip .l { font-size:9px;  font-weight:700; margin-top:1px; }

#main {
  position:relative; z-index:1;
  display:grid;
  grid-template-columns:1fr 1fr;
  grid-template-rows:auto auto;
  gap:5px;
  padding:0 5px;
}
/* C3 + other đều full-width → hàng 2 */
#card-c3    { grid-column:1; grid-row:2; }
#other-col  { grid-column:2; grid-row:2; }

#other-col {
  border-radius:8px;
  background:#0d1b35;
  border:1px solid rgba(255,255,255,.08);
  padding:8px 8px;
}

/* card header compact */
.card-hdr {
  padding:6px 10px;
  display:flex; align-items:center; gap:5px;
}
.card-hdr .shift-lbl { font-size:14px; font-weight:900; color:#fff; }
.card-hdr .shift-sub { font-size:10px; font-weight:600; color:rgba(255,255,255,.75); }
.card-hdr .shift-time{ font-size:9px;  color:rgba(255,255,255,.6); margin-left:auto; }
.card-hdr .shift-tz  { font-size:8px;  color:rgba(255,255,255,.4); }
.card-hdr .shift-cnt { background:rgba(255,255,255,.22); border-radius:999px; padding:1px 8px; font-size:11px; font-weight:900; color:#fff; margin-left:4px; }

/* person row compact */
.p-row { display:flex; align-items:center; gap:6px; padding:3px 0; border-bottom:1px solid #f1f5f9; }
.p-row:last-child { border-bottom:none; }
.p-av  { width:22px; height:22px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:8px; font-weight:900; flex-shrink:0; }
.p-nm   { font-size:11px; font-weight:700; color:#0f172a; }
.p-role { font-size:9px; font-weight:700; margin-left:4px; padding:1px 6px; border-radius:4px; white-space:nowrap; }
.p-dp  { font-size:9px;  color:#94a3b8; font-weight:500; margin-left:3px; }

/* other section compact */
.oth-sec  { margin-bottom:8px; }
.oth-lbl  { display:inline-flex; align-items:center; gap:4px; border-radius:5px; padding:3px 8px; font-size:9px; font-weight:800; margin-bottom:5px; }
.oth-chips{ display:flex; flex-wrap:wrap; gap:3px; }
.oth-chip { border-radius:5px; padding:3px 7px; font-size:10px; font-weight:700; border-width:1px; border-style:solid; white-space:nowrap; }

#footer { display:none; }

#dlBtn {
  padding:4px 10px;
  background:rgba(255,255,255,.1); color:#e0f2fe;
  border:1px solid rgba(99,179,237,.35); border-radius:7px;
  font-family:'Inter',sans-serif; font-size:11px; font-weight:700;
  cursor:pointer; flex-shrink:0;
}
#dlBtn:disabled { opacity:.4; cursor:default; }
#status { display:none; }

/* ── DESKTOP (≥ 900px) ── */
@media (min-width:900px) {
  body { padding-bottom:20px; }
  #topbar { padding:12px 24px 10px; flex-direction:row; }
  #topbar-date .day { font-size:30px !important; }
  #topbar-date .dt  { font-size:13px !important; letter-spacing:2px !important; }
  .stat-chip { padding:6px 14px; border-radius:10px; min-width:52px; }
  .stat-chip .n { font-size:20px; }
  .stat-chip .l { font-size:11px; }
  #stats-row { gap:7px; }
  #dlBtn { font-size:13px; padding:7px 16px; }
  #status { display:inline; font-size:11px; color:#38bdf8; }
  #main {
    grid-template-columns:1fr 1fr 1fr 320px;
    grid-template-rows:auto;
    gap:10px; padding:0 12px;
  }
  #card-c3   { grid-column:auto; grid-row:auto; }
  #other-col { grid-column:auto; grid-row:auto; padding:14px 13px; }
  .card-hdr  { padding:10px 14px; }
  .card-hdr .shift-lbl { font-size:16px; }
  .card-hdr .shift-sub { font-size:13px; }
  .card-hdr .shift-time{ font-size:12px; }
  .card-hdr .shift-tz  { font-size:10px; }
  .card-hdr .shift-cnt { font-size:13px; }
  .p-row { gap:9px; padding:5px 0; }
  .p-av  { width:28px; height:28px; font-size:10px; }
  .p-nm   { font-size:14px; }
  .p-role { font-size:11px; }
  .p-dp  { font-size:11px; }
  .oth-lbl  { font-size:12px; padding:4px 11px; }
  .oth-chip { font-size:13px; padding:5px 11px; border-radius:6px; }
  .oth-sec  { margin-bottom:14px; }
  #other-col > div:first-child { font-size:13px; margin-bottom:12px; }
  #footer { display:block; text-align:center; margin-top:10px; padding:8px 0; border-top:1px solid rgba(255,255,255,.07); font-size:11px; color:#94a3b8; letter-spacing:.8px; }
}
</style>
</head>
<body>

<div id="topbar">
  <div id="topbar-brand">
    <div style="width:8px;height:8px;border-radius:50%;background:#38bdf8;flex-shrink:0;"></div>
    <span style="font-size:11px;font-weight:900;color:#e0f2fe;letter-spacing:2px;text-transform:uppercase;">Tech Support</span>
  </div>
  <div id="topbar-date">
    <div class="day" style="font-size:20px;font-weight:900;color:#fff;line-height:1;">${dayName}</div>
    <div class="dt"  style="font-size:10px;color:#93c5fd;margin-top:1px;font-weight:600;letter-spacing:1px;">${dateStr}</div>
  </div>
  <div id="stats-row">
    <div class="stat-chip"><span class="n" style="color:#4ade80;">${totalWorking}</span><span class="l" style="color:#4ade80;">Làm</span></div>
    ${statChipsHtml}
    <div class="stat-chip"><span class="n" style="color:#fbbf24;">${totalAll}</span><span class="l" style="color:#fbbf24;">Tổng</span></div>
    <button id="dlBtn" onclick="dlPng()">⬇ PNG</button>
    <span id="status"></span>
  </div>
</div>

<div id="main">
  ${shiftCard("C1")}
  ${shiftCard("C2")}
  ${shiftCard("C3")}
  <div id="other-col">
    <div style="font-size:11px;font-weight:900;color:#e2e8f0;letter-spacing:1px;text-transform:uppercase;margin-bottom:8px;">Trạng thái khác</div>
    ${otherHtml}
  </div>
</div>

<div id="footer">
  Shift 1: ${shiftTimes.C1} &nbsp;·&nbsp; Shift 2: ${shiftTimes.C2} &nbsp;·&nbsp; Shift 3: ${shiftTimes.C3} &nbsp;·&nbsp; All times <strong>${tzLabel}</strong> &nbsp;·&nbsp; TECH TEAM
</div>

<script>
function dlPng() {
  const btn = document.getElementById('dlBtn');
  const st  = document.getElementById('status');
  btn.disabled = true; btn.textContent = '⏳ Đang tạo...';
  st.textContent = '';
  html2canvas(document.body, {
    scale: 1.5,
    useCORS: true,
    backgroundColor: '#0B1426',
    logging: false,
    windowWidth: document.body.scrollWidth,
    windowHeight: document.body.scrollHeight
  }).then(canvas => {
    const d=new Date(), dd=String(d.getDate()).padStart(2,'0'), mm=String(d.getMonth()+1).padStart(2,'0');
    const a=document.createElement('a');
    a.href=canvas.toDataURL('image/png');
    a.download='tech-shift-'+dd+mm+d.getFullYear()+'.png';
    a.click();
    btn.disabled=false; btn.textContent='⬇ Tải PNG';
    st.textContent='✅ Đã tải!';
    setTimeout(()=>st.textContent='',3000);
  });
}
<\/script>
</body>
</html>`;
}
