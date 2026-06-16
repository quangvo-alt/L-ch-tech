const fs = require("fs");

// Stub Utilities
global.Utilities = {
  formatDate(date, _tz, fmt) {
    const d = new Date(date);
    const pad = n => String(n).padStart(2,"0");
    if (fmt === "yyyy") return String(d.getFullYear());
    if (fmt === "M")    return String(d.getMonth() + 1);
    if (fmt === "d")    return String(d.getDate());
    if (fmt === "D")    return String(Math.ceil((d - new Date(d.getFullYear(),0,0))/86400000));
    if (fmt === "EEEE") return ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"][d.getDay()];
    if (fmt === "MMMM dd, yyyy") return d.toLocaleDateString("en-US",{year:"numeric",month:"long",day:"2-digit"});
    return fmt;
  }
};

// Extract and eval only centralOffsetLabel + buildHtml
const src = fs.readFileSync("TechShiftDashboard_v7.gs","utf8");

// Extract a top-level function by counting braces
function extractFn(source, name) {
  const start = source.indexOf("function " + name);
  if (start === -1) return null;
  let depth = 0, i = start, found = false;
  while (i < source.length) {
    if (source[i] === "{") { depth++; found = true; }
    else if (source[i] === "}") { depth--; if (found && depth === 0) return source.slice(start, i + 1); }
    i++;
  }
  return null;
}
const fn1 = extractFn(src, "centralOffsetLabel");
const fn2 = extractFn(src, "buildHtml");
if (!fn1 || !fn2) { console.error("Could not extract functions"); process.exit(1); }

eval(fn1 + "\n" + fn2);

// Mock data
const grouped = {
  C1: [
    { name:"Nguyen Van An",    role:"Supervisor", st:"WORK" },
    { name:"Tran Thi Bich",    role:"Tech",       st:"WORK" },
    { name:"Le Van Cuong",     role:"Tech",       st:"WORK" },
    { name:"Pham Thi Dung",    role:"Tech",       st:"OFF"  },
    { name:"Hoang Van Em",     role:"Tech",       st:"ME"   },
  ],
  C2: [
    { name:"Vo Thi Phuong",    role:"Manager",    st:"WORK" },
    { name:"Nguyen Minh Quan", role:"Tech",       st:"WORK" },
    { name:"Tran Van Son",     role:"Tech",       st:"NP"   },
    { name:"Le Thi Thu",       role:"Tech",       st:"HO"   },
  ],
  C3: [
    { name:"Do Van Uyen",      role:"Tech",       st:"WORK" },
    { name:"Bui Thi Viet",     role:"Tech",       st:"WORK" },
    { name:"Cao Minh Xuan",    role:"Supervisor", st:"WORK" },
    { name:"Dinh Van Yen",     role:"Tech",       st:"OFF"  },
  ],
};

const today  = new Date("2026-06-15");
const ct     = centralOffsetLabel(today, "America/Chicago");
const off    = ct.offset;
const tzLabel = ct.label;
function vnToLocal(h) { return ((h - 7 + off) % 24 + 24) % 24; }

const shiftTimesRaw = {
  C1: { s:vnToLocal(20), e:vnToLocal(4)  },
  C2: { s:vnToLocal(4),  e:vnToLocal(12) },
  C3: { s:vnToLocal(22), e:vnToLocal(5)  },
};

const quote = { text:"The only way to do great work is to love what you do.", author:"Steve Jobs" };
const html  = buildHtml(grouped, "June 15, 2026", "Monday", shiftTimesRaw, tzLabel, quote);
fs.writeFileSync("preview.html", html, "utf8");
console.log("Done: preview.html");
