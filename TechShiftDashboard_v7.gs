// ============================================================
//  TECH SHIFT DASHBOARD v7
//  Layout tự co theo nội dung · Chữ to · Không cắt
//  Tab: "Filter" | Cột G = Tên | H = Dept | I = Shift
//  Shift: C1 C2 C3 ME NP HO OFF
//  Deploy: Execute as Me · Who has access: Anyone
// ============================================================

// Offset & nhãn giờ US Central, tự xử lý DST theo ngày.
// DST: từ Chủ nhật thứ 2 của tháng 3 đến Chủ nhật đầu tiên của tháng 11.
function centralOffsetLabel(date, tz) {
  const y = parseInt(Utilities.formatDate(date, tz, "yyyy"), 10);
  const m = parseInt(Utilities.formatDate(date, tz, "M"), 10);   // 1-12
  const d = parseInt(Utilities.formatDate(date, tz, "d"), 10);
  function nthSunday(year, month1, n) {     // ngày-trong-tháng của Chủ nhật thứ n
    const dow = new Date(Date.UTC(year, month1 - 1, 1)).getUTCDay();  // 0 = Chủ nhật
    return 1 + ((7 - dow) % 7) + (n - 1) * 7;
  }
  let dst;
  if (m < 3 || m > 11)      dst = false;
  else if (m > 3 && m < 11) dst = true;
  else if (m === 3)         dst = d >= nthSunday(y, 3, 2);
  else /* m === 11 */       dst = d <  nthSunday(y, 11, 1);
  return dst ? { offset: -5, label: "CDT (UTC−5)" } : { offset: -6, label: "CST (UTC−6)" };
}

function doGet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Filter");
  if (!sheet) return HtmlService.createHtmlOutput("<p>❌ Không tìm thấy tab 'Filter'. Vui lòng kiểm tra tên tab trong Google Sheet.</p>");

  const dateCell = sheet.getRange("I1").getValue();
  const tz       = ss.getSpreadsheetTimeZone();
  let today      = dateCell ? new Date(dateCell) : new Date();
  if (isNaN(today.getTime())) today = new Date();  // I1 không phải ngày hợp lệ → dùng hôm nay
  const dateStr  = Utilities.formatDate(today, tz, "MMMM dd, yyyy");
  const dayName  = Utilities.formatDate(today, tz, "EEEE");

  // US Central tự động theo DST: hè CDT (UTC−5), đông CST (UTC−6)
  const ct          = centralOffsetLabel(today, tz);
  const localOffset = ct.offset;
  const tzLabel     = ct.label;

  function vnToLocal(vnHour) {
    // VN = UTC+7, so UTC = VN - 7; local = UTC + localOffset
    return ((vnHour - 7 + localOffset) % 24 + 24) % 24;
  }

  // VN fixed: C1 8PM→4AM, C2 4AM→12PM(noon), C3 10PM→5AM
  const shiftTimesRaw = {
    C1: { s: vnToLocal(20), e: vnToLocal(4)  },
    C2: { s: vnToLocal(4),  e: vnToLocal(12) },
    C3: { s: vnToLocal(22), e: vnToLocal(5)  },
  };

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return HtmlService.createHtmlOutput("<p>Không có dữ liệu</p>");
  const raw = sheet.getRange(3, 5, lastRow - 2, 6).getValues();  // cột E..J

  // Daily motivational quote — changes every day
  const quotes = [
    { text: "The only way to do great work is to love what you do.", author: "Steve Jobs" },
    { text: "Success is not final, failure is not fatal: it is the courage to continue that counts.", author: "Winston Churchill" },
    { text: "Believe you can and you're halfway there.", author: "Theodore Roosevelt" },
    { text: "It does not matter how slowly you go as long as you do not stop.", author: "Confucius" },
    { text: "Hard work beats talent when talent doesn't work hard.", author: "Tim Notke" },
    { text: "The secret of getting ahead is getting started.", author: "Mark Twain" },
    { text: "Don't watch the clock; do what it does. Keep going.", author: "Sam Levenson" },
    { text: "Great things never come from comfort zones.", author: "Unknown" },
    { text: "Push yourself, because no one else is going to do it for you.", author: "Unknown" },
    { text: "Dream it. Wish it. Do it.", author: "Unknown" },
    { text: "Success doesn't just find you. You have to go out and get it.", author: "Unknown" },
    { text: "The harder you work for something, the greater you'll feel when you achieve it.", author: "Unknown" },
    { text: "Dream bigger. Do bigger.", author: "Unknown" },
    { text: "Don't stop when you're tired. Stop when you're done.", author: "Unknown" },
    { text: "Wake up with determination. Go to bed with satisfaction.", author: "Unknown" },
    { text: "Do something today that your future self will thank you for.", author: "Sean Patrick Flanery" },
    { text: "Little things make big days.", author: "Unknown" },
    { text: "It's going to be hard, but hard does not mean impossible.", author: "Unknown" },
    { text: "Don't wait for opportunity. Create it.", author: "Unknown" },
    { text: "Sometimes we're tested not to show our weaknesses, but to discover our strengths.", author: "Unknown" },
    { text: "The key to success is to focus on goals, not obstacles.", author: "Unknown" },
    { text: "Enthusiasm is the electricity of life.", author: "Gordon Parks" },
    { text: "If you want to achieve greatness, stop asking for permission.", author: "Unknown" },
    { text: "Things work out best for those who make the best of how things work out.", author: "John Wooden" },
    { text: "To live a creative life, we must lose our fear of being wrong.", author: "Joseph Chilton Pearce" },
    { text: "If you are not willing to risk the usual, you will have to settle for the ordinary.", author: "Jim Rohn" },
    { text: "Trust because you are willing to accept the risk, not because it's safe or certain.", author: "Unknown" },
    { text: "Take up one idea. Make that one idea your life.", author: "Swami Vivekananda" },
    { text: "All our dreams can come true if we have the courage to pursue them.", author: "Walt Disney" },
    { text: "Good things come to people who wait, but better things come to those who go out and get them.", author: "Unknown" },
    { text: "If you do what you always did, you will get what you always got.", author: "Albert Einstein" },
  ];
  const dayOfYear = parseInt(Utilities.formatDate(today, tz, "D"), 10);
  const quote     = quotes[dayOfYear % quotes.length];

  // Cột: E = ca cố định | H = tên | I = vai trò | J = trạng thái hôm nay
  const WORK  = { C1:1, C2:1, C3:1 };
  const LEAVE = { OFF:1, ME:1, NP:1, HO:1 };
  const grouped = { C1:[], C2:[], C3:[] };
  const norm = v => (v == null ? "" : v).toString().trim().toUpperCase();

  raw.forEach(row => {
    const home   = norm(row[0]);                       // E = ca cố định
    const name   = (row[3] || "").toString().trim();   // H = tên
    const role   = (row[4] || "").toString().trim();   // I = vai trò
    const status = norm(row[5]);                        // J = trạng thái hôm nay
    if (!name) return;
    if (/[0-9:\/]/.test(name)) return;                 // dòng rác (legend/giờ giấc)
    if (WORK[status]) {                                // đi làm → card theo ca hôm nay (J)
      grouped[status].push({ name, role, st: "WORK" });
    } else {                                           // OFF/ME/NP/HO → card theo ca cố định (E)
      if (!WORK[home]) return;                         // không rõ ca cố định → bỏ qua
      const st = LEAVE[status] ? status : (role ? "OFF" : null);  // J lạ: có vai trò → OFF, không → bỏ
      if (st) grouped[home].push({ name, role, st });
    }
  });

  return HtmlService
    .createHtmlOutput(buildHtml(grouped, dateStr, dayName, shiftTimesRaw, tzLabel, quote))
    .setTitle("Tech Shift " + dateStr)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function buildHtml(grouped, dateStr, dayName, shiftTimesRaw, tzLabel, quote) {

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

  // Màu hàng cho người nghỉ (nằm trong card ca): OFF xám · ME xanh · NP vàng · HO đỏ
  const LEAVE_META = {
    OFF: { tint:"#f8fafc", avBg:"#e2e8f0", avTxt:"#475569", pillBg:"#e2e8f0", pillTxt:"#475569" },
    ME:  { tint:"#eff6ff", avBg:"#dbeafe", avTxt:"#1d4ed8", pillBg:"#dbeafe", pillTxt:"#1d4ed8" },
    NP:  { tint:"#fffbeb", avBg:"#fef3c7", avTxt:"#92400e", pillBg:"#fde68a", pillTxt:"#92400e" },
    HO:  { tint:"#fef2f2", avBg:"#fee2e2", avTxt:"#991b1b", pillBg:"#fecaca", pillTxt:"#991b1b" },
  };

  function cleanName(raw) { return raw.replace(/\s*\([^)]*\)/g,"").trim(); }
  function cleanRole(raw) { return (raw || "").trim().replace(/[,;.]+$/, "").trim(); }
  function esc(s) { return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
  function roleStyle(role) {
    const r = role.toLowerCase();
    if (r.includes("mgr"))                                    return "background:#fef3c7;color:#92400e;border:1px solid #fcd34d;font-weight:800;";  // Amber — Manager
    if (r.includes("supr") || r.includes("supervisor"))       return "background:#dbeafe;color:#1d4ed8;border:1px solid #93c5fd;font-weight:800;";  // Blue — Supervisor
    return "";  // no style for other roles
  }
  function initials(raw) {
    return cleanName(raw).split(/\s+/).filter(w=>w.length>0).slice(-2).map(w=>w[0].toUpperCase()).join("") || "?";
  }

  const workCount  = { C1:0, C2:0, C3:0 };
  const leaveCount = { ME:0, NP:0, HO:0, OFF:0 };
  ["C1","C2","C3"].forEach(s => (grouped[s]||[]).forEach(p => {
    if (p.st === "WORK") workCount[s]++;
    else if (leaveCount[p.st] !== undefined) leaveCount[p.st]++;
  }));
  const totalWorking = workCount.C1 + workCount.C2 + workCount.C3;
  const totalAll     = totalWorking + leaveCount.ME + leaveCount.NP + leaveCount.HO + leaveCount.OFF;
  const statCount    = Object.assign({}, workCount, leaveCount);

  // Top stat chips
  const statChipsHtml = ["C1","C2","C3","ME","NP","HO","OFF"].map(s => {
    const m = meta[s]; const n = statCount[s] || 0;
    return `<div class="stat-chip"><span class="n outfit" style="color:${m.statClr};">${n}</span><span class="l" style="color:${m.statClr};">${s}</span></div>`;
  }).join("");

  // Person row inside shift card
  // WORK  = sáng, chấm xanh hoạt động
  // LEAVE = mờ hơn (chỉ pill có màu), không tô nền
  function personRow(p, m) {
    const ini     = esc(initials(p.name));
    const name    = esc(cleanName(p.name));
    const rawRole = cleanRole(p.role);
    const role    = esc(rawRole);
    if (p.st === "WORK") {
      const rs       = roleStyle(rawRole);
      const roleHtml = role ? `<span class="${rs ? 'p-role p-badge' : 'p-role'}" style="${rs}">${role}</span>` : "";
      return `<div class="p-row">
        <span class="active-dot"></span>
        <div class="p-av" style="background:${m.avBg};color:${m.avTxt};">${ini}</div>
        <span class="p-nm">${name}</span>${roleHtml}
      </div>`;
    }
    const lm       = LEAVE_META[p.st] || LEAVE_META.OFF;
    const roleHtml = role ? `<span class="p-role p-role-dim">${role}</span>` : "";
    return `<div class="p-row p-row-leave">
      <div class="p-av p-av-dim" style="background:${lm.avBg};color:${lm.avTxt};">${ini}</div>
      <span class="p-nm p-nm-dim">${name}</span>${roleHtml}
      <span class="leave-pill" style="background:${lm.pillBg};color:${lm.pillTxt};">${p.st}</span>
    </div>`;
  }

  // Shift card C1/C2/C3 — đi làm trước, nghỉ (OFF/ME/NP/HO) sau; badge = số người đi làm
  function shiftCard(s) {
    const people = grouped[s] || [];
    const work   = people.filter(p => p.st === "WORK");
    const leave  = people.filter(p => p.st !== "WORK");
    const m = meta[s];
    const rows = work.concat(leave).map(p => personRow(p, m)).join("");
    const id   = s === "C3" ? ' id="card-c3"' : (s === "C2" ? ' id="card-c2"' : '');
    return `<div${id} style="border-radius:10px;overflow:hidden;border:1px solid rgba(255,255,255,.08);display:flex;flex-direction:column;">
      <div class="card-hdr" style="background:linear-gradient(90deg,${m.grad});">
        <span class="shift-lbl outfit">${s}</span>
        <span class="shift-sub">${m.label}</span>
        <span class="shift-time">${m.time} <span class="shift-tz">${tzLabel}</span></span>
        <span class="shift-cnt outfit">${work.length}</span>
      </div>
      <div style="background:#fff;padding:3px 10px 6px;">${rows || '<span style="font-size:11px;color:#ccc;">—</span>'}</div>
    </div>`;
  }

  return `<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Tech Shift ${dateStr}</title>
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@700;800;900&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"><\/script>
<style>
* { box-sizing:border-box; margin:0; padding:0; }
html { background:#0B1426; }
body {
  font-family:'Plus Jakarta Sans',sans-serif;
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
/* Outfit — tiêu đề & số */
.outfit { font-family:'Outfit',sans-serif; }
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
  grid-template-columns:1fr;
  gap:5px;
  padding:0 5px;
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
.p-nm   { font-size:13px; font-weight:800; color:#0f172a; }
.p-role  { font-size:10px; font-weight:600; margin-left:4px; color:#64748b; white-space:nowrap; }
.p-badge { padding:1px 6px; border-radius:4px; border-width:1px; border-style:solid; }
.p-dp  { font-size:9px;  color:#94a3b8; font-weight:500; margin-left:3px; }
.leave-pill { margin-left:auto; flex-shrink:0; font-size:9px; font-weight:800; padding:1px 7px; border-radius:999px; letter-spacing:.5px; }
/* active dot — người đang làm việc */
.active-dot { width:7px; height:7px; border-radius:50%; background:#22c55e; box-shadow:0 0 5px #22c55e; flex-shrink:0; }
/* leave row — mờ hơn, không tô nền */
.p-row-leave { border-bottom-color:#e2e8f0; }
.p-av-dim    { opacity:.5; }
.p-nm-dim    { color:#94a3b8 !important; font-weight:600; }
.p-role-dim  { font-size:10px; font-weight:500; margin-left:4px; color:#b0bec5; white-space:nowrap; }

#footer { display:none; }

#dlBtn {
  padding:4px 10px;
  background:rgba(255,255,255,.1); color:#e0f2fe;
  border:1px solid rgba(99,179,237,.35); border-radius:7px;
  font-family:'Plus Jakarta Sans',sans-serif; font-size:11px; font-weight:700;
  cursor:pointer; flex-shrink:0;
}
#dlBtn:disabled { opacity:.4; cursor:default; }
#status { display:none; }
#quote-bar {
  position:relative; z-index:1; overflow:hidden;
  margin:10px 5px 6px;
  padding:14px 18px;
  background:linear-gradient(135deg,rgba(30,58,138,.55),rgba(15,32,87,.6));
  border:1px solid rgba(99,179,237,.35);
  border-radius:12px;
  box-shadow:inset 0 0 0 1px rgba(99,179,237,.08), 0 6px 22px rgba(0,0,0,.35), 0 0 26px rgba(59,130,246,.12);
  display:flex; align-items:center; justify-content:center; gap:10px; flex-wrap:nowrap;
  text-align:center;
}
#quote-bar::before {
  content:''; position:absolute; top:0; left:0; right:0; height:3px;
  background:linear-gradient(90deg,transparent,#3b82f6,#60a5fa,#3b82f6,transparent);
}
.qmark        { font-family:'Outfit',serif; font-weight:900; color:rgba(96,165,250,.45); line-height:.8; flex-shrink:0; font-size:34px; }
.qmark-l      { align-self:flex-start; }
.qmark-r      { align-self:flex-end; }
#quote-inner  { display:flex; flex-direction:column; align-items:center; gap:4px; }
#quote-text   { font-size:13px; color:#e2e8f0; font-style:italic; line-height:1.45; font-weight:500; }
#quote-author { font-size:11px; color:#60a5fa; font-weight:800; letter-spacing:1px; text-transform:uppercase; white-space:nowrap; }

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
    grid-template-columns:1fr 1fr 1fr;
    gap:10px; padding:0 12px;
  }
  #card-c2   { margin-left:16px; }
  .card-hdr  { padding:10px 14px; }
  .card-hdr .shift-lbl { font-size:16px; }
  .card-hdr .shift-sub { font-size:14px; }
  .card-hdr .shift-time{ font-size:12px; }
  .card-hdr .shift-tz  { font-size:10px; }
  .card-hdr .shift-cnt { font-size:13px; }
  .p-row { gap:10px; padding:6px 0; }
  .p-av  { width:31px; height:31px; font-size:11px; }
  .p-nm   { font-size:16px; font-weight:800; }
  .p-role { font-size:12px; }
  .p-dp  { font-size:12px; }
  .leave-pill { font-size:11px; padding:2px 9px; }
  .active-dot { width:8px; height:8px; }
  .p-role-dim { font-size:12px; }
  #footer { display:block; text-align:center; margin-top:12px; padding:12px 0; border-top:1px solid rgba(255,255,255,.1); font-size:15px; color:#cbd5e1; letter-spacing:1px; font-weight:600; }
  #footer strong { color:#93c5fd; font-weight:800; }
  #quote-bar { margin:12px 12px 10px; padding:13px 28px; border-radius:14px; gap:14px; }
  #quote-bar::before { height:3px; }
  .qmark        { font-size:42px; }
  #quote-inner  { gap:5px; }
  #quote-text   { font-size:15px; line-height:1.45; }
  #quote-author { font-size:11px; letter-spacing:1.5px; }
}
</style>
</head>
<body>

<div id="topbar">
  <div id="topbar-brand">
    <div style="width:12px;height:12px;border-radius:50%;background:#22c55e;box-shadow:0 0 6px #22c55e;flex-shrink:0;"></div>
    <span class="outfit" style="font-size:16px;font-weight:900;color:#e0f2fe;letter-spacing:3px;text-transform:uppercase;">Tech Support</span>
  </div>
  <div id="topbar-date">
    <div class="day outfit" style="font-size:20px;font-weight:900;color:#fff;line-height:1;">${dayName}</div>
    <div class="dt outfit" style="font-size:10px;color:#93c5fd;margin-top:1px;font-weight:700;letter-spacing:1px;">${dateStr}</div>
  </div>
  <div id="stats-row">
    <div class="stat-chip"><span class="n outfit" style="color:#4ade80;">${totalWorking}</span><span class="l" style="color:#4ade80;">Work</span></div>
    ${statChipsHtml}
    <div class="stat-chip"><span class="n outfit" style="color:#fbbf24;">${totalAll}</span><span class="l" style="color:#fbbf24;">Total</span></div>
    <button id="dlBtn" onclick="copyImg()">📋 Copy</button>
    <span id="status"></span>
  </div>
</div>

<div id="main">
  ${shiftCard("C1")}
  ${shiftCard("C3")}
  ${shiftCard("C2")}
</div>

<div id="footer">
  Shift 1: ${shiftTimes.C1} &nbsp;·&nbsp; Shift 2: ${shiftTimes.C2} &nbsp;·&nbsp; Shift 3: ${shiftTimes.C3} &nbsp;·&nbsp; All times <strong>${tzLabel}</strong> &nbsp;·&nbsp; TECH TEAM
</div>
<div id="quote-bar">
  <span class="qmark qmark-l">&ldquo;</span>
  <div id="quote-inner">
    <span id="quote-text">${esc(quote.text)}</span>
    <span id="quote-author">— ${esc(quote.author)}</span>
  </div>
  <span class="qmark qmark-r">&rdquo;</span>
</div>

<script>
function copyImg() {
  const btn = document.getElementById('dlBtn');
  const st  = document.getElementById('status');
  btn.disabled = true; btn.textContent = '⏳ Đang tạo...';
  st.textContent = '';

  function reset() { btn.disabled = false; btn.textContent = '📋 Copy'; }
  function flash(msg, ms) { st.textContent = msg; setTimeout(() => st.textContent = '', ms || 4000); }
  function download(canvas) {
    const a = document.createElement('a');
    a.href = canvas.toDataURL('image/png');
    a.download = 'tech-shift.png'; a.click();
  }

  // Chờ font web load xong (tránh chữ bị mờ / fallback) rồi mới chụp
  const ready = document.fonts && document.fonts.ready ? document.fonts.ready : Promise.resolve();
  ready.then(function() {
    const W = document.body.scrollWidth;
    const H = document.body.scrollHeight;
    // Telegram nén ảnh paste về tối đa ~1280px cạnh dài. Render ~1920px (supersample 1.5x)
    // để khi Telegram thu xuống vẫn sắc, rõ chữ — file nhẹ, mở nhanh trong tin nhắn.
    const scale = Math.max(1.25, Math.min(2, 1920 / W));
    return html2canvas(document.body, {
      scale: scale,
      useCORS: true,
      backgroundColor: '#0B1426',
      logging: false,
      imageTimeout: 0,
      windowWidth: W,
      windowHeight: H
    });
  }).then(canvas => {
    canvas.toBlob(blob => {
      if (!blob) { reset(); flash('❌ Không tạo được ảnh, thử lại.'); return; }
      const canCopy = navigator.clipboard && window.ClipboardItem;
      if (!canCopy) { download(canvas); reset(); flash('⚠️ Trình duyệt không hỗ trợ copy, đã tải xuống.'); return; }
      navigator.clipboard.write([new ClipboardItem({'image/png': blob})])
        .then(() => { reset(); flash('✅ Đã copy! Paste vào Telegram.'); })
        .catch(() => { download(canvas); reset(); flash('⚠️ Clipboard bị chặn, đã tải xuống.'); });
    }, 'image/png');
  }).catch(err => {
    reset();
    flash('❌ Lỗi tạo ảnh: ' + (err && err.message ? err.message : err), 6000);
  });
}
<\/script>
</body>
</html>`;
}
