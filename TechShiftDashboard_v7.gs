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
    .createHtmlOutput(buildHtml(grouped, shiftOrder, dateStr, dayName))
    .setTitle("Tech Shift " + dateStr)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function buildHtml(grouped, shiftOrder, dateStr, dayName) {

  const meta = {
    C1:  { label:"Ca 1 — Công chuẩn",     time:"7:00 AM – 3:00 PM",  grad:"#1d4ed8,#3b82f6", avBg:"#dbeafe", avTxt:"#1d4ed8", statClr:"#60a5fa", chipBg:"#1e1b4b", chipTxt:"#c7d2fe", chipBd:"#3730a3", lbBg:"#1d4ed8", lbTxt:"#eff6ff" },
    C2:  { label:"Ca 2 — Công chuẩn",     time:"3:00 PM – 11:00 PM", grad:"#b45309,#f97316", avBg:"#ffedd5", avTxt:"#b45309", statClr:"#fb923c", chipBg:"#431407", chipTxt:"#fed7aa", chipBd:"#9a3412", lbBg:"#c2410c", lbTxt:"#fff7ed" },
    C3:  { label:"Ca 3 — Công chuẩn",     time:"9:00 AM – 4:00 PM",  grad:"#15803d,#22c55e", avBg:"#dcfce7", avTxt:"#15803d", statClr:"#4ade80", chipBg:"#052e16", chipTxt:"#bbf7d0", chipBd:"#166534", lbBg:"#15803d", lbTxt:"#f0fdf4" },
    ME:  { label:"Mac Energy — Nghỉ có công", time:"",               statClr:"#a78bfa", chipBg:"#2e1065", chipTxt:"#ddd6fe", chipBd:"#5b21b6", lbBg:"#4c1d95", lbTxt:"#ede9fe" },
    NP:  { label:"Nghỉ phép năm — Có công",  time:"",               statClr:"#fbbf24", chipBg:"#1c1407", chipTxt:"#fde68a", chipBd:"#92400e", lbBg:"#92400e", lbTxt:"#fef3c7" },
    HO:  { label:"Nghỉ lễ / Bù lễ — Có công", time:"",             statClr:"#f87171", chipBg:"#450a0a", chipTxt:"#fecaca", chipBd:"#7f1d1d", lbBg:"#991b1b", lbTxt:"#fee2e2" },
    OFF: { label:"Nghỉ",                     time:"",               statClr:"#94a3b8", chipBg:"#1e293b", chipTxt:"#f1f5f9", chipBd:"#475569", lbBg:"#334155", lbTxt:"#f1f5f9" },
  };

  function cleanName(raw) { return raw.replace(/\s*\([^)]*\)/g,"").trim(); }
  function initials(raw) {
    return cleanName(raw).split(/\s+/).slice(-2).map(w=>w[0].toUpperCase()).join("");
  }

  const totalWorking = ["C1","C2","C3"].reduce((s,k)=>s+(grouped[k]||[]).length,0);
  const totalAll     = shiftOrder.reduce((s,k)=>s+(grouped[k]||[]).length,0);

  // Top stat chips
  const statChipsHtml = shiftOrder.map(s => {
    const m = meta[s]; const n = (grouped[s]||[]).length;
    return `<div style="display:flex;flex-direction:column;align-items:center;background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.12);border-radius:10px;padding:6px 14px;min-width:52px;">
      <span style="font-size:20px;font-weight:900;color:${m.statClr};line-height:1;">${n}</span>
      <span style="font-size:11px;font-weight:700;color:${m.statClr};margin-top:2px;">${s}</span>
    </div>`;
  }).join("");

  // Person row inside shift card
  function pRow(p, m, last) {
    const ini  = initials(p.name);
    const name = cleanName(p.name);
    const dept = p.dept && p.dept.toUpperCase() !== "TECH"
      ? `<span style="font-size:11px;color:#94a3b8;margin-left:5px;font-weight:500;">${p.dept}</span>` : "";
    return `<div style="display:flex;align-items:center;gap:9px;padding:5px 0;${last?"":"border-bottom:1px solid #f1f5f9;"}">
      <div style="width:28px;height:28px;border-radius:50%;background:${m.avBg};color:${m.avTxt};display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:900;flex-shrink:0;">${ini}</div>
      <span style="font-size:14px;font-weight:700;color:#0f172a;">${name}</span>${dept}
    </div>`;
  }

  // Shift card C1/C2/C3
  function shiftCard(s) {
    const people = grouped[s]||[];
    const m = meta[s];
    const rows = people.map((p,i)=>pRow(p,m,i===people.length-1)).join("");
    return `<div style="border-radius:12px;overflow:hidden;border:1px solid rgba(255,255,255,.08);display:flex;flex-direction:column;">
      <div style="background:linear-gradient(90deg,${m.grad});padding:10px 14px;display:flex;align-items:center;gap:8px;">
        <span style="font-size:16px;font-weight:900;color:#fff;letter-spacing:.5px;">${s}</span>
        <span style="font-size:13px;font-weight:600;color:rgba(255,255,255,.8);">${m.label}</span>
        <span style="font-size:12px;color:rgba(255,255,255,.55);margin-left:auto;">${m.time}</span>
        <span style="background:rgba(255,255,255,.22);border-radius:999px;padding:2px 10px;font-size:13px;font-weight:900;color:#fff;margin-left:6px;">${people.length}</span>
      </div>
      <div style="background:#fff;padding:4px 12px 8px;">${rows || '<span style="font-size:13px;color:#ccc;">—</span>'}</div>
    </div>`;
  }

  // Name chip for "other" col
  function chip(p, m) {
    const name = cleanName(p.name);
    const dept = p.dept && p.dept.toUpperCase() !== "TECH"
      ? ` <span style="font-size:10px;opacity:.7;">${p.dept}</span>` : "";
    return `<div style="border-radius:6px;padding:5px 11px;font-size:13px;font-weight:700;background:${m.chipBg};color:${m.chipTxt};border:1px solid ${m.chipBd};white-space:nowrap;">${name}${dept}</div>`;
  }

  // Section inside "other" col
  function otherSec(s) {
    const people = grouped[s]||[]; const m = meta[s];
    const chips  = people.length
      ? people.map(p=>chip(p,m)).join("")
      : `<span style="font-size:12px;color:#334155;font-style:italic;">Không có</span>`;
    return `<div style="margin-bottom:14px;">
      <div style="display:inline-flex;align-items:center;gap:5px;border-radius:6px;padding:4px 11px;background:${m.lbBg};color:${m.lbTxt};font-size:12px;font-weight:800;margin-bottom:7px;">
        ${s} — ${m.label}
        <span style="background:rgba(255,255,255,.2);border-radius:999px;padding:0 7px;font-size:12px;">${people.length}</span>
      </div>
      <div style="display:flex;flex-wrap:wrap;gap:5px;">${chips}</div>
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
  min-width:1200px;
  padding-bottom:20px;
  position:relative;
}
body::before {
  content:''; position:fixed; inset:0; pointer-events:none; z-index:0;
  background-image:
    linear-gradient(rgba(99,179,237,.04) 1px,transparent 1px),
    linear-gradient(90deg,rgba(99,179,237,.04) 1px,transparent 1px);
  background-size:40px 40px;
}
#topbar {
  position:relative; z-index:1;
  display:flex; align-items:center; justify-content:space-between;
  padding:12px 24px 10px;
  background:linear-gradient(90deg,#0f2057,#1e3a8a);
  border-bottom:1px solid rgba(99,179,237,.2);
  margin-bottom:12px;
}
#main {
  position:relative; z-index:1;
  display:grid;
  grid-template-columns:1fr 1fr 1fr 320px;
  gap:10px;
  padding:0 12px;
}
#other-col {
  border-radius:12px;
  background:#0d1b35;
  border:1px solid rgba(255,255,255,.08);
  padding:14px 13px;
}
#footer {
  position:relative; z-index:1;
  text-align:center; margin-top:10px; padding:8px 0;
  border-top:1px solid rgba(255,255,255,.07);
  font-size:11px; color:#94a3b8; letter-spacing:.8px;
}
#dlBtn {
  padding:7px 16px;
  background:rgba(255,255,255,.1); color:#e0f2fe;
  border:1px solid rgba(99,179,237,.35); border-radius:8px;
  font-family:'Inter',sans-serif; font-size:13px; font-weight:700;
  cursor:pointer; margin-left:12px;
}
#dlBtn:hover    { background:rgba(255,255,255,.18); }
#dlBtn:disabled { opacity:.4; cursor:default; }
#status { font-size:11px; color:#38bdf8; margin-left:10px; }
</style>
</head>
<body>

<div id="topbar">
  <div style="display:flex;align-items:center;gap:10px;">
    <div style="width:10px;height:10px;border-radius:50%;background:#38bdf8;flex-shrink:0;"></div>
    <span style="font-size:13px;font-weight:900;color:#e0f2fe;letter-spacing:3px;text-transform:uppercase;">Tech Support</span>
  </div>
  <div style="text-align:center;">
    <div style="font-size:30px;font-weight:900;color:#fff;line-height:1;">${dayName}</div>
    <div style="font-size:13px;color:#93c5fd;margin-top:2px;font-weight:600;letter-spacing:2px;">${dateStr}</div>
  </div>
  <div style="display:flex;align-items:center;gap:7px;flex-wrap:wrap;justify-content:flex-end;">
    <div style="display:flex;flex-direction:column;align-items:center;background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.12);border-radius:10px;padding:6px 14px;">
      <span style="font-size:20px;font-weight:900;color:#4ade80;line-height:1;">${totalWorking}</span>
      <span style="font-size:11px;font-weight:700;color:#4ade80;margin-top:2px;">Đi làm</span>
    </div>
    ${statChipsHtml}
    <div style="display:flex;flex-direction:column;align-items:center;background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.12);border-radius:10px;padding:6px 14px;">
      <span style="font-size:20px;font-weight:900;color:#fbbf24;line-height:1;">${totalAll}</span>
      <span style="font-size:11px;font-weight:700;color:#fbbf24;margin-top:2px;">Tổng</span>
    </div>
    <button id="dlBtn" onclick="dlPng()">⬇ Tải PNG</button>
    <span id="status"></span>
  </div>
</div>

<div id="main">
  ${shiftCard("C1")}
  ${shiftCard("C2")}
  ${shiftCard("C3")}
  <div id="other-col">
    <div style="font-size:13px;font-weight:900;color:#e2e8f0;letter-spacing:1px;text-transform:uppercase;margin-bottom:12px;">Trạng thái khác</div>
    ${otherHtml}
  </div>
</div>

<div id="footer">
  C1: 7:00 AM – 3:00 PM &nbsp;·&nbsp; C2: 3:00 PM – 11:00 PM &nbsp;·&nbsp; C3: 9:00 AM – 4:00 PM &nbsp;·&nbsp; All times in <strong>CST (Central Standard Time, UTC−6)</strong> &nbsp;·&nbsp; TECH TEAM — Have a great day!
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
