// ============================================================
//  TECH SHIFT DASHBOARD v6 — FINAL
//  Layout 1920×1080 · Dark theme · Font Inter
//  Tab: "Filter" | Cột G = Tên | H = Dept | I = Shift
//  Shift values: C1 C2 C3 ME NP HO OFF
//  Row 1: header (I1 = ngày) | Row 2: tiêu đề | Row 3+: data
//  Deploy: Execute as Me · Who has access: Anyone
// ============================================================

function doGet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Filter");

  const dateCell = sheet.getRange("I1").getValue();
  const tz       = ss.getSpreadsheetTimeZone();
  const today    = dateCell ? new Date(dateCell) : new Date();
  const dateStr  = Utilities.formatDate(today, tz, "dd/MM/yyyy");
  const dayNames = ["Chủ nhật","Thứ hai","Thứ ba","Thứ tư","Thứ năm","Thứ sáu","Thứ bảy"];
  const dayName  = dayNames[today.getDay()];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return HtmlService.createHtmlOutput("<p>Không có dữ liệu</p>");
  const raw = sheet.getRange(3, 7, lastRow - 2, 3).getValues();

  const shiftOrder = ["C1","C2","C3","ME","NP","HO","OFF"];
  const grouped    = {};
  shiftOrder.forEach(s => grouped[s] = []);

  raw.forEach(row => {
    const name  = (row[0] || "").toString().trim();
    const dept  = (row[1] || "").toString().trim();
    const shift = (row[2] || "").toString().trim().toUpperCase();
    if (!name) return;
    if (grouped[shift] !== undefined) {
      grouped[shift].push({ name, dept });
    }
  });

  return HtmlService
    .createHtmlOutput(buildHtml(grouped, shiftOrder, dateStr, dayName))
    .setTitle("Tech Shift " + dateStr)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
function buildHtml(grouped, shiftOrder, dateStr, dayName) {

  const meta = {
    C1:  { label:"Ca 1",            time:"7:00 AM – 3:00 PM",  hdrGrad:"#1d4ed8,#3b82f6", avBg:"#dbeafe", avTxt:"#1d4ed8", statClr:"#60a5fa", chipBg:"#1e1b4b", chipTxt:"#a5b4fc", chipBorder:"#3730a3", labelBg:"#1d4ed8", labelTxt:"#eff6ff" },
    C2:  { label:"Ca 2",            time:"3:00 PM – 11:00 PM", hdrGrad:"#c2410c,#f97316", avBg:"#ffedd5", avTxt:"#c2410c", statClr:"#fb923c", chipBg:"#431407", chipTxt:"#fdba74", chipBorder:"#9a3412", labelBg:"#c2410c", labelTxt:"#fff7ed" },
    C3:  { label:"Ca 3",            time:"9:00 AM – 4:00 PM",  hdrGrad:"#15803d,#22c55e", avBg:"#dcfce7", avTxt:"#15803d", statClr:"#34d399", chipBg:"#052e16", chipTxt:"#86efac", chipBorder:"#166534", labelBg:"#15803d", labelTxt:"#f0fdf4" },
    ME:  { label:"Meeting",         time:"—",                  statClr:"#a78bfa", chipBg:"#1e1b4b", chipTxt:"#a5b4fc", chipBorder:"#4338ca", labelBg:"#3b0764", labelTxt:"#e9d5ff" },
    NP:  { label:"Nghỉ phép",       time:"—",                  statClr:"#f472b6", chipBg:"#1f0d1a", chipTxt:"#f9a8d4", chipBorder:"#9d174d", labelBg:"#500724", labelTxt:"#fce7f3" },
    HO:  { label:"Nghỉ lễ / Bù",   time:"—",                  statClr:"#f87171", chipBg:"#450a0a", chipTxt:"#fca5a5", chipBorder:"#7f1d1d", labelBg:"#991b1b", labelTxt:"#fecaca" },
    OFF: { label:"Nghỉ",            time:"—",                  statClr:"#94a3b8", chipBg:"#1f2937", chipTxt:"#f1f5f9", chipBorder:"#374151", labelBg:"#374151", labelTxt:"#f1f5f9" },
  };

  const totalWorking = ["C1","C2","C3"].reduce((s,k) => s + (grouped[k]||[]).length, 0);
  const totalAll     = shiftOrder.reduce((s,k) => s + (grouped[k]||[]).length, 0);

  // Clean name: strip parentheses content for display name
  function cleanName(raw) {
    return raw.replace(/\s*\([^)]*\)/g, "").trim();
  }
  function initials(raw) {
    const words = cleanName(raw).split(/\s+/);
    return words.slice(-2).map(w => w[0].toUpperCase()).join("");
  }

  // Stat chips top bar
  const statChips = shiftOrder.map(s => {
    const n = (grouped[s]||[]).length;
    const m = meta[s];
    return `<div style="display:flex;flex-direction:column;align-items:center;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);border-radius:8px;padding:4px 11px;min-width:44px;">
      <span style="font-size:16px;font-weight:900;color:${m.statClr};line-height:1;">${n}</span>
      <span style="font-size:9px;font-weight:700;color:${m.statClr};margin-top:1px;opacity:.85;">${s}</span>
    </div>`;
  }).join("");

  // Person row for C1/C2/C3 columns
  function personRow(p, m, isLast) {
    const ini  = initials(p.name);
    const name = cleanName(p.name);
    const dept = p.dept && p.dept !== "Tech" ? `<span style="font-size:9.5px;color:#94a3b8;margin-left:3px;">${p.dept}</span>` : "";
    return `<div style="display:flex;align-items:center;gap:6px;padding:3.5px 0;${isLast ? "" : "border-bottom:1px solid #f1f5f9;"}">
      <div style="width:22px;height:22px;border-radius:50%;background:${m.avBg};color:${m.avTxt};display:flex;align-items:center;justify-content:center;font-size:8px;font-weight:900;flex-shrink:0;">${ini}</div>
      <span style="font-size:11.5px;font-weight:700;color:#0f172a;">${name}</span>${dept}
    </div>`;
  }

  // Shift card (C1/C2/C3)
  function shiftCard(s) {
    const people = grouped[s]||[];
    const m = meta[s];
    const rows = people.map((p,i) => personRow(p, m, i===people.length-1)).join("");
    return `<div style="border-radius:10px;overflow:hidden;display:flex;flex-direction:column;min-height:0;border:1px solid rgba(255,255,255,.07);">
      <div style="background:linear-gradient(90deg,${m.hdrGrad});padding:7px 11px;display:flex;align-items:center;gap:7px;flex-shrink:0;">
        <span style="font-size:14px;font-weight:900;color:#fff;letter-spacing:.5px;">${s}</span>
        <span style="font-size:11px;font-weight:600;color:rgba(255,255,255,.75);">${m.label}</span>
        <span style="font-size:10px;color:rgba(255,255,255,.5);margin-left:auto;">${m.time}</span>
        <span style="background:rgba(255,255,255,.2);border-radius:999px;padding:1px 8px;font-size:12px;font-weight:900;color:#fff;margin-left:5px;">${people.length}</span>
      </div>
      <div style="flex:1;overflow:hidden;padding:3px 10px 5px;background:#fff;">${rows}</div>
    </div>`;
  }

  // Chip for "other" column
  function nameChip(p, m) {
    const name = cleanName(p.name);
    const dept = p.dept && p.dept !== "Tech" ? ` <span style="opacity:.7;font-size:9px;">${p.dept}</span>` : "";
    return `<div style="border-radius:5px;padding:4px 9px;font-size:10.5px;font-weight:700;background:${m.chipBg};color:${m.chipTxt};border:1px solid ${m.chipBorder};white-space:nowrap;">${name}${dept}</div>`;
  }

  // "Other" section block inside the 4th column
  function otherSection(s) {
    const people = grouped[s]||[];
    const m = meta[s];
    const chips = people.length
      ? people.map(p => nameChip(p, m)).join("")
      : `<span style="font-size:10px;color:#1e293b;">—</span>`;
    return `<div style="flex-shrink:0;">
      <div style="display:inline-flex;align-items:center;gap:4px;border-radius:5px;padding:3px 9px;font-size:10px;font-weight:800;background:${m.labelBg};color:${m.labelTxt};margin-bottom:5px;">
        ${s} — ${m.label} <span style="opacity:.7;font-weight:600;">(${people.length})</span>
      </div>
      <div style="display:flex;flex-wrap:wrap;gap:4px;">${chips}</div>
    </div>`;
  }

  const otherSections = ["ME","NP","HO","OFF"].map(s => otherSection(s)).join("");

  return `<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=1920">
<title>Tech Shift ${dateStr}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"><\/script>
<style>
* { box-sizing:border-box; margin:0; padding:0; }
html, body {
  width:1920px; height:1080px;
  overflow:hidden;
  font-family:'Inter',sans-serif;
  background:#0B1426;
}
body {
  display:flex; flex-direction:column;
  position:relative;
}
body::before {
  content:''; position:absolute; inset:0; pointer-events:none;
  background-image:
    linear-gradient(rgba(99,179,237,.04) 1px,transparent 1px),
    linear-gradient(90deg,rgba(99,179,237,.04) 1px,transparent 1px);
  background-size:40px 40px;
}
#topbar {
  flex-shrink:0; display:flex; align-items:center; justify-content:space-between;
  padding:10px 22px 8px;
  background:linear-gradient(90deg,#0f2057,#1e3a8a);
  border-bottom:1px solid rgba(99,179,237,.2);
  position:relative; z-index:1;
}
#main {
  flex:1; display:grid;
  grid-template-columns:1fr 1fr 1fr 1fr;
  gap:8px; padding:8px 12px 8px;
  min-height:0; position:relative; z-index:1;
}
#other-col {
  border-radius:10px; display:flex; flex-direction:column;
  gap:9px; padding:10px 11px; min-height:0; overflow:hidden;
  background:#0d1b35; border:1px solid rgba(255,255,255,.07);
}
#footer {
  flex-shrink:0; text-align:center; padding:5px 0 6px;
  background:#080e1c; border-top:1px solid rgba(255,255,255,.06);
  font-size:9.5px; color:#94a3b8; letter-spacing:.8px; text-transform:uppercase;
  position:relative; z-index:1;
}
#dlBtn {
  padding:5px 14px;
  background:rgba(255,255,255,.1); color:#e0f2fe;
  border:1px solid rgba(99,179,237,.3); border-radius:7px;
  font-family:'Inter',sans-serif; font-size:12px; font-weight:700;
  cursor:pointer; white-space:nowrap; margin-left:10px;
}
#dlBtn:hover    { background:rgba(255,255,255,.18); }
#dlBtn:disabled { opacity:.4; cursor:default; }
</style>
</head>
<body>

<div id="topbar">
  <div style="display:flex;align-items:center;gap:9px;">
    <div style="width:9px;height:9px;border-radius:50%;background:#38bdf8;"></div>
    <span style="font-size:11px;font-weight:900;color:#e0f2fe;letter-spacing:3px;text-transform:uppercase;">Tech Support</span>
  </div>
  <div style="text-align:center;">
    <div style="font-size:26px;font-weight:900;color:#fff;line-height:1;">${dayName}</div>
    <div style="font-size:11px;color:#93c5fd;margin-top:1px;font-weight:600;letter-spacing:2px;">${dateStr}</div>
  </div>
  <div style="display:flex;align-items:center;gap:6px;">
    <div style="display:flex;flex-direction:column;align-items:center;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);border-radius:8px;padding:4px 11px;min-width:50px;">
      <span style="font-size:16px;font-weight:900;color:#4ade80;line-height:1;">${totalWorking}</span>
      <span style="font-size:9px;font-weight:700;color:#4ade80;margin-top:1px;">Đi làm</span>
    </div>
    ${statChips}
    <div style="display:flex;flex-direction:column;align-items:center;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);border-radius:8px;padding:4px 11px;min-width:50px;">
      <span style="font-size:16px;font-weight:900;color:#fbbf24;line-height:1;">${totalAll}</span>
      <span style="font-size:9px;font-weight:700;color:#fbbf24;margin-top:1px;">Tổng</span>
    </div>
    <button id="dlBtn" onclick="downloadPng()">⬇ Tải PNG</button>
  </div>
</div>

<div id="main">
  ${shiftCard("C1")}
  ${shiftCard("C2")}
  ${shiftCard("C3")}
  <div id="other-col">
    <div style="font-size:11px;font-weight:900;color:#e2e8f0;letter-spacing:1px;text-transform:uppercase;flex-shrink:0;">
      Trạng thái khác
    </div>
    ${otherSections}
  </div>
</div>

<div id="footer">
  Shift 1: 7:00 AM – 3:00 PM &nbsp;·&nbsp; Shift 2: 3:00 PM – 11:00 PM &nbsp;·&nbsp; Shift 3: 9:00 AM – 4:00 PM (Central Time) &nbsp;·&nbsp; TECH TEAM — Have a great day!
</div>

<script>
function downloadPng() {
  const btn = document.getElementById('dlBtn');
  btn.disabled = true;
  btn.textContent = '⏳ Đang tạo...';
  html2canvas(document.body, {
    scale: 1,
    width: 1920, height: 1080,
    windowWidth: 1920, windowHeight: 1080,
    useCORS: true,
    backgroundColor: '#0B1426',
    logging: false
  }).then(canvas => {
    const d  = new Date();
    const dd = String(d.getDate()).padStart(2,'0');
    const mm = String(d.getMonth()+1).padStart(2,'0');
    const yy = d.getFullYear();
    const a  = document.createElement('a');
    a.href     = canvas.toDataURL('image/png');
    a.download = 'tech-shift-' + dd + mm + yy + '.png';
    a.click();
    btn.disabled    = false;
    btn.textContent = '⬇ Tải PNG';
  });
}
<\/script>
</body>
</html>`;
}
