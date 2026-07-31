import { useState, useEffect, useMemo, useRef, createContext, useContext } from "react";
import * as XLSX from "xlsx";
import { getClubDoc, saveClubDoc, subscribeClubDoc, storageEnabled } from "./lib/clubStore";
import { uploadFile } from "./lib/files";
import { registerPush, sendPush } from "./lib/push";

/* ============ Theme (club-configurable; Rotaract defaults) ============ */
let CRAN = "#D41367", CRAN_DK = "#A50D50", AZURE = "#0067C8";
const GOLD = "#F7A81B", INK = "#2A1420", PAPER = "#F6F3F5", CARD = "#FFFFFF", LINE = "#E9E1E6", OK = "#1E8E5A", BAD = "#C22F2F";
const DISPLAY = "'Archivo', system-ui, sans-serif";
const BODY = "'Public Sans', system-ui, sans-serif";
function shade(hex, f) {
  const n = hex.replace("#", "");
  if (n.length !== 6) return hex;
  const r = Math.round(parseInt(n.slice(0, 2), 16) * (1 + f)), g = Math.round(parseInt(n.slice(2, 4), 16) * (1 + f)), b = Math.round(parseInt(n.slice(4, 6), 16) * (1 + f));
  const c = (x) => Math.max(0, Math.min(255, x)).toString(16).padStart(2, "0");
  return `#${c(r)}${c(g)}${c(b)}`;
}
function applyTheme(colors) {
  CRAN = colors?.primary || "#D41367";
  CRAN_DK = shade(CRAN, -0.25);
  AZURE = colors?.secondary || "#0067C8";
}
const fontCss = () => `
@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@600;700;800;900&family=Public+Sans:wght@400;500;600;700&display=swap');
*{-webkit-tap-highlight-color:transparent}
::-webkit-scrollbar{display:none}
input,textarea,select{outline:none}
input:focus,textarea:focus,select:focus{border-color:${CRAN} !important; box-shadow:0 0 0 3px ${CRAN}1f}
button:focus-visible{outline:2px solid ${AZURE}; outline-offset:2px}
@media (prefers-reduced-motion: reduce){*{transition:none !important; animation:none !important}}
`;

/* ============ Domain constants ============ */
const ROLES = ["President", "Vice President", "Secretary", "Treasurer", "Past President", "Sergeant-at-Arms", "Board Member", "Member", "Prospect"];
const EBOD = ["President", "Vice President", "Secretary", "Treasurer", "Past President"];
const ROLE_COLOR = {
  President: "#D41367", "Vice President": "#0067C8", Secretary: "#B07400", Treasurer: "#1E8E5A",
  "Past President": "#7A4BA6", "Sergeant-at-Arms": "#B34700", "Board Member": "#0E8F8F", Member: "#6B5A64", Prospect: "#9A8B93",
};
const MEMBER_STATUSES = ["Active", "Prospect", "Applied", "Invited", "On Leave", "Transferred", "Resigned", "Alumni", "Inactive"];
const ACTIVE_LIKE = ["Active", "On Leave"];
const AREAS = ["Peacebuilding & Conflict Prevention", "Disease Prevention & Treatment", "Water, Sanitation & Hygiene", "Maternal & Child Health", "Basic Education & Literacy", "Community Economic Development", "Environment"];
const TX_CATS = ["Dues", "Fundraising", "Donations", "Grants", "Events", "Service Projects", "Supplies", "Venue", "Transport", "Fees", "Other"];
const PAY_METHODS = ["Cash", "Bank transfer", "Cheque", "Mobile money", "Card", "Other"];
const CHARGE_KINDS = { monthly: "Monthly dues", district: "District dues", ri: "Rotary International", penalty: "Penalty", happy: "Happy Dollars", late: "Late charge", carryforward: "Balance carried forward", other: "Other charge" };
const ATT_STATUSES = ["present", "late", "virtual", "excused", "absent"];
const ATT_COLOR = { present: OK, late: GOLD, virtual: AZURE, excused: "#7A4BA6", absent: BAD };
const EVENT_KINDS = {
  general: { label: "General Meeting", short: "General", color: "#D41367" },
  professional: { label: "Professional Development", short: "Prof. Dev", color: "#0067C8" },
  social: { label: "Social", short: "Social", color: "#7A4BA6" },
  orientation: { label: "Club Orientation", short: "Orientation", color: "#0E8F8F" },
  drive: { label: "Membership Drive", short: "Drive", color: "#B07400" },
  fundraiser: { label: "Fundraiser", short: "Fundraiser", color: "#1E8E5A" },
  project: { label: "Project Date", short: "Project", color: "#B34700" },
};
const eventKind = (m) => EVENT_KINDS[m?.kind] || EVENT_KINDS.general;
const TASK_STATUSES = ["Not Started", "In Progress", "Blocked", "Awaiting Review", "Completed"];
const TASK_COLOR = { "Not Started": "#8A7580", "In Progress": GOLD, Blocked: BAD, "Awaiting Review": "#0067C8", Completed: OK };
const PRIORITIES = ["Low", "Medium", "High"];
const PRIORITY_COLOR = { Low: "#8A7580", Medium: "#B07400", High: BAD };
const NOTIF_CATS = {
  meetings: { label: "Meetings & agendas", essential: true },
  minutes: { label: "Published minutes", essential: false },
  dues: { label: "Dues & payments", essential: true },
  projects: { label: "Project decisions & updates", essential: false },
  tasks: { label: "Task assignments & deadlines", essential: false },
  announcements: { label: "Club announcements", essential: true },
  library: { label: "New library documents", essential: false },
  membership: { label: "Membership changes", essential: true },
};
const LIB_CATS = ["Constitution & Bylaws", "Guidelines", "Branding", "Training", "Policies", "Reports", "Assets & Equipment", "Photos & Media", "Other"];
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/* ============ Helpers ============ */
const uid = () => Math.random().toString(36).slice(2, 10);
const todayStr = () => new Date().toISOString().slice(0, 10);
const clone = (o) => JSON.parse(JSON.stringify(o));
const money = (n, c = "$") => `${n < 0 ? "−" : ""}${c}${Math.abs(Number(n || 0)).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const fmtDate = (d) => d ? new Date(d + (d.length === 10 ? "T12:00:00" : "")).toLocaleDateString(undefined, { weekday: "short", month: "short", day: "numeric", year: "numeric" }) : "";
const fmtShort = (d) => d ? new Date(d + (d.length === 10 ? "T12:00:00" : "")).toLocaleDateString(undefined, { month: "short", day: "numeric" }) : "";
const timeAgo = (ts) => { const s = Math.floor((Date.now() - ts) / 1000); if (s < 60) return "just now"; if (s < 3600) return `${Math.floor(s / 60)}m ago`; if (s < 86400) return `${Math.floor(s / 3600)}h ago`; return `${Math.floor(s / 86400)}d ago`; };
const initials = (name) => (name || "?").split(" ").map((p) => p[0]).slice(0, 2).join("").toUpperCase();
const pad2 = (n) => String(n).padStart(2, "0");

/* Rotary year: 1 June — 31 May */
const ryStartYearFor = (dateStr) => { const d = new Date(dateStr + "T12:00:00"); return d.getMonth() + 1 >= 6 ? d.getFullYear() : d.getFullYear() - 1; };
const ryId = (startYear) => `RY${startYear}`;
const ryLabel = (startYear) => `${startYear}–${String(startYear + 1).slice(2)}`;
const ryBounds = (startYear) => ({ start: `${startYear}-06-01`, end: `${startYear + 1}-05-31` });
const ryMonths = (startYear) => { const out = []; for (let i = 0; i < 12; i++) { const m = ((5 + i) % 12) + 1; const y = m >= 6 ? startYear : startYear + 1; out.push(`${y}-${pad2(m)}`); } return out; };
const inYear = (dateStr, year) => dateStr >= year.start && dateStr <= year.end;

/* Downloads */
function downloadBlob(name, content, type) {
  const blob = content instanceof Blob ? content : new Blob([content], { type: type || "text/plain" });
  const u = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = u; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(u), 8000);
}
function downloadDataUrl(name, dataUrl) {
  const a = document.createElement("a"); a.href = dataUrl; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
}
function csvOf(rows) { return rows.map((r) => r.map((c) => `"${String(c ?? "").replace(/"/g, '""')}"`).join(",")).join("\n"); }
function exportCsv(name, rows) { downloadBlob(name + ".csv", csvOf(rows), "text/csv"); }
function exportXlsx(name, rows) {
  try {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadBlob(name + ".xlsx", new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  } catch (e) { exportCsv(name, rows); }
}
function docShell(title, inner, accent) {
  return `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title><style>
  body{font-family:'Segoe UI',system-ui,sans-serif;color:#2A1420;max-width:760px;margin:32px auto;padding:0 20px}
  h1{color:${accent};margin-bottom:4px} h2{border-bottom:2px solid ${accent};padding-bottom:4px;margin-top:28px;font-size:18px}
  table{width:100%;border-collapse:collapse;margin:10px 0;font-size:13px} th,td{border:1px solid #ddd;padding:6px 8px;text-align:left}
  th{background:${accent}14} .muted{color:#8A7580;font-size:12px} .big{font-size:26px;font-weight:800}
  @media print{body{margin:0}}
  </style></head><body>${inner}<p class="muted">Generated ${new Date().toLocaleString()}</p></body></html>`;
}
function exportHtml(name, title, inner, accent) { downloadBlob(name + ".html", docShell(title, inner, accent || CRAN), "text/html"); }
// Prints a complete docShell() document via a hidden iframe (rather than
// window.open, which installed/standalone PWAs block — there's no "new tab"
// in standalone mode, so it silently fell back to an .html download there).
// The browser's print dialog then offers "Save as PDF" as a destination.
function printHtml(name, html) {
  const cleanup = (iframe) => { if (iframe.parentNode) iframe.parentNode.removeChild(iframe); };
  try {
    const iframe = document.createElement("iframe");
    Object.assign(iframe.style, { position: "fixed", right: 0, bottom: 0, width: 0, height: 0, border: 0 });
    document.body.appendChild(iframe);
    iframe.onload = () => {
      const win = iframe.contentWindow;
      win.onafterprint = () => cleanup(iframe);
      win.focus();
      win.print();
      setTimeout(() => cleanup(iframe), 60000); // fallback if afterprint never fires
    };
    iframe.srcdoc = html;
  } catch (e) {
    downloadBlob(`${name}.html`, html, "text/html");
  }
}

/* File handling: cloud storage when configured, small base64 files in local demo mode */
const MAX_FILE_KB = 400;           // demo-mode cap (files live inside the club document)
const MAX_FILE_MB_CLOUD = 15;      // cloud cap (files live in Supabase Storage)
function readFileAsDataUrl(file, cb, onErr) {
  if (storageEnabled()) {
    if (file.size > MAX_FILE_MB_CLOUD * 1024 * 1024) { onErr(`File is too large — keep uploads under ${MAX_FILE_MB_CLOUD} MB.`); return; }
    uploadFile(CLUB_ID, file)
      .then((url) => cb({ id: uid(), name: file.name, type: file.type, size: file.size, dataUrl: url, at: Date.now() }))
      .catch((e) => onErr("Upload failed: " + (e.message || e)));
    return;
  }
  if (file.size > MAX_FILE_KB * 1024) { onErr(`File is ${Math.round(file.size / 1024)} KB — keep uploads under ${MAX_FILE_KB} KB in demo mode.`); return; }
  const r = new FileReader();
  r.onload = () => cb({ id: uid(), name: file.name, type: file.type, size: file.size, dataUrl: r.result, at: Date.now() });
  r.onerror = () => onErr("Could not read that file.");
  r.readAsDataURL(file);
}

/* ============ Storage (Supabase-backed; localStorage in demo mode) ============ */
let CLUB_ID = "demo";
let _freshHandler = null; // called when a save loses an optimistic-concurrency race
async function loadShared(_key, fallback) {
  try { const d = await getClubDoc(CLUB_ID); return d ?? fallback; } catch (e) { console.error(e); return fallback; }
}
async function saveShared(_key, val) {
  try { await saveClubDoc(CLUB_ID, val); }
  catch (e) { if (e && e.fresh && _freshHandler) _freshHandler(e.fresh); else console.error(e); }
}
function loadMine(key, fallback) {
  try { const v = localStorage.getItem(`rotaract:${CLUB_ID}:${key}`); return v ? JSON.parse(v) : fallback; } catch (e) { return fallback; }
}
function saveMine(key, val) {
  try {
    if (val === null) localStorage.removeItem(`rotaract:${CLUB_ID}:${key}`);
    else localStorage.setItem(`rotaract:${CLUB_ID}:${key}`, JSON.stringify(val));
  } catch (e) {}
}

/* ============ Seed ============ */
function blankMemberExtras() {
  return { photo: null, dob: "", address: "", occupation: "", employer: "", emergency: "", rotaryId: "", interests: "", skills: "", positions: [], statusHistory: [], notes: "" };
}
function seedDb() {
  const now = Date.now();
  const curRY = ryStartYearFor(todayStr());
  const prevRY = curRY - 1;
  const yCur = { id: ryId(curRY), startYear: curRY, label: ryLabel(curRY), ...ryBounds(curRY), active: true, archivedAt: null, committees: [
    { id: "c1", name: "Service Projects", chair: "m6", members: ["m6", "m7"] },
    { id: "c2", name: "Fundraising", chair: "m2", members: ["m2", "m4"] },
  ] };
  const yPrev = { id: ryId(prevRY), startYear: prevRY, label: ryLabel(prevRY), ...ryBounds(prevRY), active: false, archivedAt: now - 60 * 86400000, committees: [] };
  const M = (id, name, role, email, joined, extra = {}) => ({
    id, name, role, email, phone: "", joined, status: "Active", ...blankMemberExtras(), ...extra,
  });
  const members = [
    M("m1", "Ana Reyes", "President", "ana@rotaract.club", `${curRY - 2}-07-01`, { occupation: "Teacher", positions: [{ yearId: yPrev.id, role: "Vice President" }] }),
    M("m2", "Devon Castillo", "Vice President", "devon@rotaract.club", `${curRY - 2}-07-01`, { occupation: "Accountant" }),
    M("m3", "Marisol Tun", "Secretary", "marisol@rotaract.club", `${curRY - 1}-07-01`),
    M("m4", "Jared Novelo", "Treasurer", "jared@rotaract.club", `${curRY - 1}-07-01`, { rotaryId: "RI-88231" }),
    M("m5", "Kimberly Chan", "Past President", "kim@rotaract.club", `${curRY - 3}-07-01`, { positions: [{ yearId: yPrev.id, role: "President" }] }),
    M("m6", "Luis Ayuso", "Board Member", "luis@rotaract.club", `${curRY - 1}-09-15`),
    M("m7", "Sasha Flores", "Member", "sasha@rotaract.club", `${curRY}-01-10`),
    M("m8", "Omar Bacab", "Prospect", "omar@mail.com", todayStr(), { status: "Applied" }),
  ];
  const duesConfig = { monthly: 10, district: 15, ri: 8, currency: "$", dueDay: 5, graceDays: 10, lateFee: 2, lateFeeOn: false };
  const nextThu = new Date(); nextThu.setDate(nextThu.getDate() + ((4 - nextThu.getDay() + 7) % 7 || 7));
  const lastThu = new Date(); lastThu.setDate(lastThu.getDate() - ((lastThu.getDay() - 4 + 7) % 7 || 7));
  const plusDays = (n) => { const d = new Date(); d.setDate(d.getDate() + n); return d.toISOString().slice(0, 10); };
  const meetings = [
    { id: "mt1", yearId: yCur.id, kind: "general", projectId: "", title: "General Meeting", date: nextThu.toISOString().slice(0, 10), time: "19:00", type: "hybrid",
      location: "Community Center, Room B", link: "https://meet.example.com/rotaract", status: "published",
      agenda: ["Call to order & roll call", "Treasurer's report", "River cleanup update", "Induction planning", "Open floor"],
      presenters: ["m4", "m6"], rsvpDeadline: nextThu.toISOString().slice(0, 10), rsvps: { m1: "yes", m3: "yes" },
      attachments: [], attendance: {}, guests: [], minutes: null, createdBy: "m1", publishedAt: now - 3 * 86400000 },
    { id: "mt0", yearId: yCur.id, kind: "general", projectId: "", title: "General Meeting", date: lastThu.toISOString().slice(0, 10), time: "19:00", type: "physical",
      location: "Community Center, Room B", link: "", status: "published",
      agenda: ["Call to order", "Bake sale results", "Cleanup scouting", "Adjournment"], presenters: [], rsvpDeadline: "", rsvps: {},
      attachments: [], createdBy: "m1", publishedAt: now - 10 * 86400000,
      attendance: { m1: "present", m2: "present", m3: "present", m4: "late", m5: "excused", m6: "present", m7: "virtual" },
      guests: [{ id: uid(), name: "Rtn. Pablo Marin", kind: "guest" }, { id: uid(), name: "Omar Bacab", kind: "prospect" }],
      minutes: { versions: [{ id: uid(), fileName: "Minutes note (text)", dataUrl: null, text: "Called to order 7:05 PM. Bake sale netted a solid surplus. Cleanup scouting approved. Adjourned 8:20 PM.", by: "m3", at: now - 6 * 86400000, note: "Original" }],
        actionItems: [{ id: uid(), text: "Collect vendor quotes for cleanup supplies", assignee: "m6", due: todayStr(), done: false }] } },
    { id: "ev1", yearId: yCur.id, kind: "social", projectId: "", title: "Games Night Social", date: plusDays(9), time: "18:30", type: "physical",
      location: "Kim's place, Santa Rita", link: "", status: "published", agenda: ["Potluck & board games — bring a friend!"],
      presenters: [], rsvpDeadline: plusDays(8), rsvps: {}, attachments: [], attendance: {}, guests: [], minutes: null, createdBy: "m2", publishedAt: now - 86400000 },
    { id: "ev2", yearId: yCur.id, kind: "fundraiser", projectId: "p1", title: "Car Wash Fundraiser", date: plusDays(16), time: "09:00", type: "physical",
      location: "Central Park corner lot", link: "", status: "published", agenda: ["Shift 1: 9–12", "Shift 2: 12–3"],
      presenters: [], rsvpDeadline: plusDays(14), rsvps: {}, attachments: [], attendance: {}, guests: [], minutes: null, createdBy: "m1", publishedAt: now - 86400000 },
    { id: "ev3", yearId: yCur.id, kind: "project", projectId: "p1", title: "River Cleanup Day", date: plusDays(23), time: "07:30", type: "physical",
      location: "New River bank — north access", link: "", status: "published", agenda: ["Safety briefing", "Team assignments", "Cleanup", "Weigh-in & photos"],
      presenters: ["m7"], rsvpDeadline: plusDays(21), rsvps: {}, attachments: [], attendance: {}, guests: [], minutes: null, createdBy: "m1", publishedAt: now - 86400000 },
  ];
  const projects = [
    { id: "p1", yearId: yCur.id, title: "New River Cleanup Day", area: "Environment", lead: "m7", submittedBy: "m7", at: now - 12 * 86400000,
      description: "One-day volunteer cleanup along the New River bank with the town council collecting waste.",
      problem: "Plastic and household waste accumulate along a 2 km stretch used by families and fishers.",
      objectives: "Remove 40+ bags of waste; recruit 25+ volunteers.", goals: "Establish an annual cleanup tradition.",
      impact: "Cleaner riverbank, community awareness.", beneficiaries: "Riverside residents and fishers (~300 people)",
      location: "New River bank, north stretch", dates: "", volunteersNeeded: 25, partners: "Town Council", budget: 350, expectedIncome: 100,
      materials: "Gloves, bags, grabbers, water", risks: "Weather; sharp objects", success: "Bags collected; volunteer count", sustainability: "Annual repeat with council MOU",
      files: [], status: "Approved",
      decisions: [{ id: uid(), by: "m1", action: "Approved", reason: "Strong impact, modest budget.", at: now - 9 * 86400000 }],
      votes: { m1: { v: "approve", comment: "Strong community impact." }, m2: { v: "approve", comment: "" }, m4: { v: "approve", comment: "Budget is fine." } },
      questions: [], ebodNotes: [], recommendations: "Coordinate pickup with council early.",
      tasks: [
        { id: "t1", title: "Confirm council waste pickup", assignee: "m6", deadline: todayStr(), status: "In Progress", priority: "High", checklist: [{ id: uid(), text: "Call sanitation office", done: true }, { id: uid(), text: "Get written confirmation", done: false }], comments: [], deps: [], attachments: [] },
        { id: "t2", title: "Design volunteer sign-up flyer", assignee: "m7", deadline: todayStr(), status: "Not Started", priority: "Medium", checklist: [], comments: [], deps: ["t1"], attachments: [] },
      ],
      milestones: [{ id: uid(), title: "Volunteers confirmed", date: "", done: false }],
      volunteerRoles: [{ id: uid(), title: "Team captain", slots: 3, filled: ["m6"] }, { id: uid(), title: "Refreshments", slots: 2, filled: [] }],
      updates: [], discussion: [], serviceHours: [], risksLog: [], finalReport: null, completedAt: null },
    { id: "p2", yearId: yCur.id, title: "Primary School Reading Corner", area: "Basic Education & Literacy", lead: "m6", submittedBy: "m6", at: now - 2 * 86400000,
      description: "Build and stock a reading corner with 100+ books at a local primary school.",
      problem: "Standard 1–3 students lack access to age-appropriate books.", objectives: "Collect 100 books; furnish one corner.",
      goals: "Improve early reading access.", impact: "Better literacy outcomes.", beneficiaries: "~90 students",
      location: "St. Paul's Primary", dates: "", volunteersNeeded: 8, partners: "", budget: 500, expectedIncome: 0,
      materials: "Shelving, mats, books", risks: "Book supply shortfall", success: "Books shelved; corner in use", sustainability: "Teacher-librarian upkeep",
      files: [], status: "Submitted", decisions: [{ id: uid(), by: "m6", action: "Submitted", reason: "", at: now - 2 * 86400000 }], votes: {}, questions: [], ebodNotes: [], recommendations: "",
      tasks: [], milestones: [], volunteerRoles: [], updates: [], discussion: [], serviceHours: [], risksLog: [], finalReport: null, completedAt: null },
  ];
  const library = [
    { id: "l0", category: "Constitution & Bylaws", title: "Club Constitution", desc: "The club's governing constitution as adopted.", content: "Standard Rotaract Club Constitution — replace with your club's adopted text.", url: "", files: [] },
    { id: "l1", category: "Guidelines", title: "Rotaract Handbook", desc: "Official guide to running a Rotaract club.", url: "https://www.rotary.org/en/get-involved/rotaract-clubs", files: [] },
    { id: "l2", category: "Branding", title: "Rotaract Brand Center", desc: "Logos, colors, and templates (My Rotary login).", url: "https://brandcenter.rotary.org/", files: [] },
    { id: "l3", category: "Training", title: "Project Lifecycle Guide", desc: "How proposals move through this portal.", content: "Draft → Submitted → Under EBOD Review → Approved / Denied / Returned for revision → active workspace → final report → EBOD closes.", url: "", files: [] },
    { id: "l4", category: "Assets & Equipment", title: "Club Asset Register", desc: "Banner, PA speaker, pop-up tent — sign out with the Sergeant-at-Arms.", content: "1× ceremonial banner\n1× PA speaker\n1× pop-up tent (3×3 m)\n2× cash boxes", url: "", files: [] },
  ];
  const transactions = [
    { id: "tx1", yearId: yCur.id, type: "income", category: "Fundraising", desc: "Bake sale proceeds", payee: "", amount: 264.5, date: yCur.start > todayStr() ? todayStr() : `${curRY}-06-20`, by: "m4", approval: "approved", projectId: "", receipt: null, reversed: false, reversalOf: null },
    { id: "tx2", yearId: yCur.id, type: "expense", category: "Venue", desc: "Community Center hall fee", payee: "Community Center", amount: 60, date: `${curRY}-06-10`, by: "m4", approval: "approved", projectId: "", receipt: null, reversed: false, reversalOf: null },
    { id: "tx0", yearId: yPrev.id, type: "income", category: "Fundraising", desc: "Car wash (last year)", payee: "", amount: 410, date: `${prevRY}-11-12`, by: "m4", approval: "approved", projectId: "", receipt: null, reversed: false, reversalOf: null },
  ];
  return {
    club: { name: "Rotaract Club", tagline: "Fellowship through service", logo: null, colors: { primary: "#D41367", secondary: "#0067C8" }, charterDate: `${curRY - 3}-06-01` },
    years: [yCur, yPrev], activeYearId: yCur.id,
    members, duesConfig,
    charges: [], payments: [], credits: [], arrangements: [], exemptions: [],
    transactions, seq: { receipt: 0 }, locks: {}, reconciliations: {},
    audit: [{ id: uid(), ts: now, by: "system", action: "Club initialized", detail: "Demo data seeded" }],
    meetings, projects, library,
    notifications: [
      { id: uid(), ts: now - 3 * 86400000, type: "meetings", title: "Meeting published", body: "General Meeting scheduled — see agenda and RSVP." },
      { id: uid(), ts: now - 2 * 86400000, type: "projects", title: "New proposal", body: "Primary School Reading Corner submitted for EBOD review." },
    ],
    seededAt: now, v: 2,
  };
}

function seedProductionDb(clubName, creator) {
  const now = Date.now();
  const curRY = ryStartYearFor(todayStr());
  const yCur = { id: ryId(curRY), startYear: curRY, label: ryLabel(curRY), ...ryBounds(curRY), active: true, archivedAt: null, committees: [] };
  return {
    club: { name: clubName || "Rotaract Club", tagline: "Fellowship through service", logo: null, colors: { primary: "#D41367", secondary: "#0067C8" }, charterDate: todayStr() },
    years: [yCur], activeYearId: yCur.id,
    members: [{ id: uid(), name: creator?.name || "President", role: "President", email: creator?.email || "", phone: "", joined: todayStr(), status: "Active", ...blankMemberExtras() }],
    duesConfig: { monthly: 0, district: 0, ri: 0, currency: "$", dueDay: 5, graceDays: 10, lateFee: 0, lateFeeOn: false },
    charges: [], payments: [], credits: [], arrangements: [], exemptions: [],
    transactions: [], seq: { receipt: 0 }, locks: {}, reconciliations: {},
    audit: [{ id: uid(), ts: now, by: creator?.name || "system", action: "Club created", detail: clubName || "" }],
    meetings: [], projects: [],
    library: [
      { id: uid(), category: "Guidelines", title: "Rotaract Handbook", desc: "Official guide to running a Rotaract club.", url: "https://www.rotary.org/en/get-involved/rotaract-clubs", content: "", files: [] },
      { id: uid(), category: "Branding", title: "Rotary Brand Center", desc: "Logos, colours, and templates (My Rotary login).", url: "https://brandcenter.rotary.org/", content: "", files: [] },
    ],
    notifications: [{ id: uid(), ts: now, type: "announcements", title: `Welcome to ${clubName || "your club"}!`, body: "Set the dues configuration, share your invite code from Club settings, and schedule your first meeting." }],
    seededAt: now, v: 2,
  };
}

/* ============ Finance derivations ============ */
function ensureObligations(d) {
  const y = d.years.find((x) => x.id === d.activeYearId);
  if (!y) return false;
  const cfg = d.duesConfig; const t = todayStr(); let changed = false;
  const activeMembers = d.members.filter((m) => ACTIVE_LIKE.includes(m.status));
  const exemptIds = new Set(d.exemptions.filter((e) => e.active).map((e) => e.memberId));
  const months = ryMonths(y.startYear).filter((p) => `${p}-01` <= t);
  activeMembers.forEach((m) => {
    if (exemptIds.has(m.id)) return;
    months.forEach((period) => {
      if (cfg.monthly > 0 && !d.charges.some((c) => c.memberId === m.id && c.kind === "monthly" && c.period === period && !c.reversed)) {
        d.charges.push({ id: uid(), yearId: y.id, memberId: m.id, kind: "monthly", label: `Monthly dues ${period}`, period, amount: cfg.monthly, dueDate: `${period}-${pad2(cfg.dueDay)}`, at: Date.now(), by: "system", reversed: false });
        changed = true;
      }
    });
    const firstP = ryMonths(y.startYear)[0];
    if (`${firstP}-01` <= t) {
      ["district", "ri"].forEach((k) => {
        const amt = k === "district" ? cfg.district : cfg.ri;
        if (amt > 0 && !d.charges.some((c) => c.memberId === m.id && c.kind === k && c.yearId === y.id && !c.reversed)) {
          d.charges.push({ id: uid(), yearId: y.id, memberId: m.id, kind: k, label: `${CHARGE_KINDS[k]} ${y.label}`, period: firstP, amount: amt, dueDate: `${firstP}-${pad2(cfg.dueDay)}`, at: Date.now(), by: "system", reversed: false });
          changed = true;
        }
      });
    }
  });
  return changed;
}
function memberAccount(d, memberId, yearId) {
  const ch = d.charges.filter((c) => c.memberId === memberId && c.yearId === yearId && !c.reversed);
  const pays = d.payments.filter((p) => p.memberId === memberId && p.yearId === yearId && !p.reversed);
  const crs = d.credits.filter((c) => c.memberId === memberId && c.yearId === yearId && !c.reversed);
  const charged = ch.reduce((s, c) => s + c.amount, 0);
  const paid = pays.reduce((s, p) => s + p.amount, 0);
  const credited = crs.reduce((s, c) => s + c.amount, 0);
  const balance = charged - paid - credited;
  const t = todayStr(); const grace = d.duesConfig.graceDays || 0;
  const pastDueCharges = ch.filter((c) => { const dd = new Date(c.dueDate + "T12:00:00"); dd.setDate(dd.getDate() + grace); return dd.toISOString().slice(0, 10) < t; });
  const pastDueTotal = pastDueCharges.reduce((s, c) => s + c.amount, 0);
  const overdue = Math.max(0, Math.min(balance, pastDueTotal - paid - credited));
  const arr = d.arrangements.find((a) => a.memberId === memberId && a.active);
  const upcoming = ch.filter((c) => c.dueDate >= t).sort((a, b) => a.dueDate.localeCompare(b.dueDate))[0];
  const nextDue = arr?.nextDate || upcoming?.dueDate || (balance > 0 ? t : null);
  return { charges: ch, payments: pays, credits: crs, charged, paid, credited, balance, overdue, nextDue, arrangement: arr };
}
function clubTotals(d, yearId) {
  const tx = d.transactions.filter((t) => t.yearId === yearId && !t.reversed && (t.type === "income" || t.approval === "approved"));
  const income = tx.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const expense = tx.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
  return { income, expense, balance: income - expense };
}
function isLocked(d, dateStr) { return !!d.locks[dateStr?.slice(0, 7)]; }
function attendanceStats(d, memberId, yearId) {
  const past = d.meetings.filter((m) => m.yearId === yearId && m.status === "published" && ["general", "professional", "orientation"].includes(m.kind || "general") && m.date < todayStr() && Object.keys(m.attendance || {}).length > 0);
  const marked = past.filter((m) => m.attendance[memberId]);
  const attended = marked.filter((m) => ["present", "late", "virtual"].includes(m.attendance[memberId]));
  return { total: past.length, attended: attended.length, pct: past.length ? Math.round((attended.length / past.length) * 100) : null };
}
function serviceHoursOf(d, memberId) {
  return d.projects.reduce((s, p) => s + (p.serviceHours || []).filter((h) => h.memberId === memberId).reduce((a, h) => a + Number(h.hours || 0), 0), 0);
}

/* ============ UI primitives ============ */
const Ctx = createContext(null);
const useApp = () => useContext(Ctx);

function Icon({ name, size = 20, color = "currentColor", strokeWidth = 2 }) {
  const p = { fill: "none", stroke: color, strokeWidth, strokeLinecap: "round", strokeLinejoin: "round" };
  const paths = {
    home: <path {...p} d="M3 10.5 12 3l9 7.5M5 9.5V21h5v-6h4v6h5V9.5" />,
    calendar: <g {...p}><rect x="3" y="5" width="18" height="16" rx="2" /><path d="M8 3v4M16 3v4M3 10h18" /></g>,
    target: <g {...p}><circle cx="12" cy="12" r="9" /><circle cx="12" cy="12" r="4.5" /></g>,
    wallet: <g {...p}><rect x="3" y="6" width="18" height="14" rx="2" /><path d="M3 10h18M16 15h2" /></g>,
    dots: <g fill={color} stroke="none"><circle cx="5" cy="12" r="1.8" /><circle cx="12" cy="12" r="1.8" /><circle cx="19" cy="12" r="1.8" /></g>,
    bell: <g {...p}><path d="M6 9a6 6 0 0 1 12 0c0 5 2 6.5 2 6.5H4S6 14 6 9Z" /><path d="M10 19a2 2 0 0 0 4 0" /></g>,
    users: <g {...p}><circle cx="9" cy="8" r="3.5" /><path d="M2.5 20c.8-3.4 3.4-5 6.5-5s5.7 1.6 6.5 5" /><circle cx="17" cy="9" r="2.5" /><path d="M16.5 14.5c2.6.2 4.4 1.6 5 4.5" /></g>,
    book: <g {...p}><path d="M4 4.5A2.5 2.5 0 0 1 6.5 2H20v17.5H6.5A2.5 2.5 0 0 0 4 22V4.5Z" /><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20" /></g>,
    plus: <path {...p} d="M12 5v14M5 12h14" />,
    check: <path {...p} d="m4.5 12.5 5 5 10-11" />,
    x: <path {...p} d="M6 6l12 12M18 6L6 18" />,
    chevR: <path {...p} d="m9 5 7 7-7 7" />,
    chevL: <path {...p} d="m15 5-7 7 7 7" />,
    chevD: <path {...p} d="m5 9 7 7 7-7" />,
    edit: <g {...p}><path d="M12 20h9" /><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z" /></g>,
    link: <g {...p}><path d="M10 14a5 5 0 0 0 7.1 0l2.4-2.4a5 5 0 0 0-7.1-7.1L11 5.9" /><path d="M14 10a5 5 0 0 0-7.1 0l-2.4 2.4a5 5 0 0 0 7.1 7.1L13 18.1" /></g>,
    clock: <g {...p}><circle cx="12" cy="12" r="9" /><path d="M12 7v5l3 2" /></g>,
    pin: <g {...p}><path d="M12 21s-7-6.1-7-11a7 7 0 0 1 14 0c0 4.9-7 11-7 11Z" /><circle cx="12" cy="10" r="2.5" /></g>,
    doc: <g {...p}><path d="M6 2h9l5 5v15H6Z" /><path d="M15 2v5h5M9 12h7M9 16h7" /></g>,
    logout: <g {...p}><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" /><path d="m16 17 5-5-5-5M21 12H9" /></g>,
    trash: <g {...p}><path d="M4 7h16M9 7V4h6v3M6 7l1 14h10l1-14" /></g>,
    download: <g {...p}><path d="M12 3v12m0 0 4-4m-4 4-4-4" /><path d="M4 17v3h16v-3" /></g>,
    upload: <g {...p}><path d="M12 16V4m0 0 4 4m-4-4-4 4" /><path d="M4 17v3h16v-3" /></g>,
    shield: <g {...p}><path d="M12 2 4 5v6c0 5 3.5 8.8 8 11 4.5-2.2 8-6 8-11V5Z" /></g>,
    video: <g {...p}><rect x="2" y="6" width="13" height="12" rx="2" /><path d="m15 10 7-3v10l-7-3" /></g>,
    globe: <g {...p}><circle cx="12" cy="12" r="9" /><path d="M3 12h18M12 3c3 3.5 3 14 0 18-3-4-3-14.5 0-18Z" /></g>,
    flag: <g {...p}><path d="M5 21V4" /><path d="M5 4h12l-2.5 3.5L17 11H5" /></g>,
    lock: <g {...p}><rect x="5" y="11" width="14" height="9" rx="2" /><path d="M8 11V7a4 4 0 0 1 8 0v4" /></g>,
    msg: <g {...p}><path d="M21 12a8 8 0 0 1-8 8H4l2-3a8 8 0 1 1 15-5Z" /></g>,
    star: <path {...p} d="m12 3 2.7 5.6 6.1.8-4.5 4.2 1.1 6L12 16.8 6.6 19.6l1.1-6L3.2 9.4l6.1-.8Z" />,
    camera: <g {...p}><path d="M4 8h3l2-2h6l2 2h3v12H4Z" /><circle cx="12" cy="13" r="3.5" /></g>,
    repeat: <g {...p}><path d="M17 2 21 6l-4 4" /><path d="M3 11V9a3 3 0 0 1 3-3h15M7 22l-4-4 4-4" /><path d="M21 13v2a3 3 0 0 1-3 3H3" /></g>,
  };
  return <svg width={size} height={size} viewBox="0 0 24 24" aria-hidden="true">{paths[name] || null}</svg>;
}

function Wheel({ size = 120, color = "rgba(255,255,255,.14)" }) {
  const spokes = [], teeth = [];
  for (let i = 0; i < 6; i++) { const a = (i * 60 * Math.PI) / 180; spokes.push(<line key={i} x1={60 + 16 * Math.cos(a)} y1={60 + 16 * Math.sin(a)} x2={60 + 44 * Math.cos(a)} y2={60 + 44 * Math.sin(a)} stroke={color} strokeWidth="7" />); }
  for (let i = 0; i < 12; i++) { const a = (i * 30 * Math.PI) / 180; teeth.push(<line key={"t" + i} x1={60 + 50 * Math.cos(a)} y1={60 + 50 * Math.sin(a)} x2={60 + 57 * Math.cos(a)} y2={60 + 57 * Math.sin(a)} stroke={color} strokeWidth="8" />); }
  return <svg width={size} height={size} viewBox="0 0 120 120" aria-hidden="true"><circle cx="60" cy="60" r="46" fill="none" stroke={color} strokeWidth="9" /><circle cx="60" cy="60" r="14" fill="none" stroke={color} strokeWidth="7" />{spokes}{teeth}</svg>;
}

function Avatar({ m, size = 38 }) {
  const c = ROLE_COLOR[m?.role] || "#888";
  if (m?.photo) return <img src={m.photo} alt="" className="rounded-full object-cover flex-shrink-0" style={{ width: size, height: size, border: `2px solid ${c}66` }} />;
  return (
    <div className="flex items-center justify-center rounded-full font-bold flex-shrink-0"
      style={{ width: size, height: size, background: c + "1f", color: c, border: `2px solid ${c}55`, fontSize: size * 0.36, fontFamily: DISPLAY }}>
      {m ? initials(m.name) : "?"}
    </div>
  );
}
function Badge({ children, color = "#6B5A64", filled = false }) {
  return <span className="inline-flex items-center gap-1 rounded-full font-semibold" style={{ background: filled ? color : color + "18", color: filled ? "#fff" : color, fontSize: 11, padding: "3px 9px" }}>{children}</span>;
}
function Btn({ children, onClick, kind = "primary", small = false, disabled = false, style = {} }) {
  const base = {
    primary: { background: CRAN, color: "#fff", border: "none" },
    dark: { background: INK, color: "#fff", border: "none" },
    blue: { background: AZURE, color: "#fff", border: "none" },
    ghost: { background: "transparent", color: CRAN, border: `1.5px solid ${CRAN}55` },
    quiet: { background: "#fff", color: INK, border: `1.5px solid ${LINE}` },
    danger: { background: "#fff", color: BAD, border: `1.5px solid ${BAD}44` },
    ok: { background: OK, color: "#fff", border: "none" },
  }[kind];
  return (
    <button onClick={onClick} disabled={disabled} className="rounded-xl font-bold active:opacity-80 disabled:opacity-40 transition-opacity"
      style={{ ...base, padding: small ? "8px 14px" : "13px 18px", fontSize: small ? 13 : 15, fontFamily: DISPLAY, width: small ? "auto" : "100%", ...style }}>
      {children}
    </button>
  );
}
function Field({ label, children, hint }) {
  return (
    <label className="block mb-4">
      <div className="font-bold uppercase mb-1.5" style={{ fontSize: 11, letterSpacing: ".08em", color: "#8A7580", fontFamily: DISPLAY }}>{label}</div>
      {children}
      {hint ? <div className="mt-1" style={{ fontSize: 12, color: "#9A8B93" }}>{hint}</div> : null}
    </label>
  );
}
const inputStyle = { border: `1.5px solid ${LINE}`, borderRadius: 12, padding: "11px 13px", fontSize: 15, width: "100%", background: "#fff", color: INK, fontFamily: BODY };
function Input(props) { return <input {...props} style={{ ...inputStyle, ...(props.style || {}) }} />; }
function TextArea(props) { return <textarea rows={props.rows || 3} {...props} style={{ ...inputStyle, resize: "vertical", ...(props.style || {}) }} />; }
function Select(props) { return <select {...props} style={{ ...inputStyle, appearance: "none", ...(props.style || {}) }}>{props.children}</select>; }

function FilePick({ label, accept, onFile, hint }) {
  const ref = useRef(null);
  const { showToast } = useApp();
  return (
    <div className="mb-4">
      {label && <div className="font-bold uppercase mb-1.5" style={{ fontSize: 11, letterSpacing: ".08em", color: "#8A7580", fontFamily: DISPLAY }}>{label}</div>}
      <input ref={ref} type="file" accept={accept} className="hidden" onChange={(e) => {
        const f = e.target.files?.[0]; if (!f) return;
        readFileAsDataUrl(f, onFile, (msg) => showToast(msg));
        e.target.value = "";
      }} />
      <Btn kind="quiet" small onClick={() => ref.current?.click()}><span className="flex items-center gap-1.5"><Icon name="upload" size={15} /> Choose file</span></Btn>
      {hint && <div className="mt-1" style={{ fontSize: 12, color: "#9A8B93" }}>{hint}</div>}
    </div>
  );
}
function FileChip({ f, onRemove }) {
  return (
    <span className="inline-flex items-center gap-1.5 rounded-full px-3 py-1.5 mr-1.5 mb-1.5" style={{ background: "#fff", border: `1px solid ${LINE}`, fontSize: 12 }}>
      <Icon name="doc" size={13} color={CRAN} />
      <button onClick={() => downloadDataUrl(f.name, f.dataUrl)} className="font-semibold" style={{ color: AZURE, maxWidth: 140, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</button>
      {onRemove && <button onClick={onRemove} aria-label="Remove"><Icon name="x" size={12} color="#A896A0" /></button>}
    </span>
  );
}
function Sheet({ open, onClose, title, children, tall = false }) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center" role="dialog" aria-modal="true">
      <div className="absolute inset-0" style={{ background: "rgba(30,10,22,.5)" }} onClick={onClose} />
      <div className="relative w-full max-w-md rounded-t-3xl overflow-y-auto" style={{ background: PAPER, maxHeight: tall ? "93vh" : "80vh", boxShadow: "0 -12px 40px rgba(0,0,0,.25)" }}>
        <div className="sticky top-0 z-10 flex items-center justify-between px-5 pt-4 pb-3" style={{ background: PAPER }}>
          <h2 className="font-extrabold" style={{ fontFamily: DISPLAY, fontSize: 19 }}>{title}</h2>
          <button onClick={onClose} aria-label="Close" className="rounded-full p-2" style={{ background: "#fff", border: `1px solid ${LINE}` }}><Icon name="x" size={16} /></button>
        </div>
        <div className="px-5 pb-8">{children}</div>
      </div>
    </div>
  );
}
function Card({ children, onClick, style = {}, className = "" }) {
  return <div onClick={onClick} className={"rounded-2xl " + className + (onClick ? " cursor-pointer active:opacity-80" : "")} style={{ background: CARD, border: `1px solid ${LINE}`, ...style }}>{children}</div>;
}
function SectionTitle({ children, action }) {
  return (
    <div className="flex items-center justify-between mt-6 mb-2.5">
      <h3 className="font-extrabold uppercase" style={{ fontFamily: DISPLAY, fontSize: 12.5, letterSpacing: ".1em", color: "#7A6570" }}>{children}</h3>
      {action}
    </div>
  );
}
function Empty({ icon, title, text }) {
  return (
    <Card className="p-6 text-center">
      <div className="flex justify-center mb-2" style={{ color: "#C9B8C1" }}><Icon name={icon} size={30} /></div>
      <div className="font-bold" style={{ fontFamily: DISPLAY, fontSize: 15 }}>{title}</div>
      {text ? <div className="mt-1" style={{ fontSize: 13, color: "#8A7580" }}>{text}</div> : null}
    </Card>
  );
}
function Progress({ pct, color }) {
  return <div className="rounded-full overflow-hidden" style={{ height: 7, background: LINE }}><div className="h-full rounded-full transition-all" style={{ width: `${Math.min(100, pct)}%`, background: color || CRAN }} /></div>;
}
function Chips({ options, value, onChange }) {
  return (
    <div className="flex gap-1.5 overflow-x-auto pb-1">
      {options.map((o) => {
        const v = typeof o === "string" ? o : o.id, label = typeof o === "string" ? o : o.label;
        return (
          <button key={v} onClick={() => onChange(v)} className="rounded-full font-bold flex-shrink-0"
            style={{ fontSize: 12.5, padding: "6px 13px", fontFamily: DISPLAY, background: value === v ? INK : "#fff", color: value === v ? "#fff" : "#6B5A64", border: `1px solid ${value === v ? INK : LINE}` }}>
            {label}
          </button>
        );
      })}
    </div>
  );
}
function KV({ k, v }) {
  if (v === undefined || v === null || v === "") return null;
  return <div className="mb-2.5"><div className="font-bold" style={{ fontSize: 12, color: "#8A7580" }}>{k}</div><div style={{ fontSize: 14, whiteSpace: "pre-wrap" }}>{v}</div></div>;
}
function FullScreen({ title, onClose, children, accent, right }) {
  return (
    <div className="fixed inset-0 z-40 flex justify-center" style={{ background: "rgba(30,10,22,.4)" }}>
      <div className="w-full max-w-md overflow-y-auto" style={{ background: PAPER }}>
        <div className="sticky top-0 z-20 flex items-center gap-3 px-4 py-3.5" style={{ background: accent || CRAN }}>
          <button onClick={onClose} aria-label="Back" className="rounded-full p-1.5 flex-shrink-0" style={{ background: "rgba(255,255,255,.18)" }}><Icon name="chevL" size={19} color="#fff" /></button>
          <h2 className="font-extrabold text-white truncate flex-1" style={{ fontFamily: DISPLAY, fontSize: 18 }}>{title}</h2>
          {right}
        </div>
        <div className="px-4 pt-4" style={{ paddingBottom: 60 }}>{children}</div>
      </div>
    </div>
  );
}

/* ============ Root App ============ */
export default function Portal({ clubId = "demo", demo = false, userEmail = "", userName = "", clubName = "", joinCode = "", onSignOutAuth }) {
  const [db, setDb] = useState(null);
  const [meId, setMeId] = useState(null);
  const [booted, setBooted] = useState(false);
  const [tab, setTab] = useState("home");
  const [overlay, setOverlay] = useState(null);
  const [lastRead, setLastRead] = useState(0);
  const [prefs, setPrefs] = useState({ muted: {} });
  const [toast, setToast] = useState(null);

  useEffect(() => {
    let unsub = null;
    (async () => {
      CLUB_ID = clubId;
      let d = await loadShared("doc", null);
      if (!d || d.v !== 2) d = demo ? seedDb() : seedProductionDb(clubName, { name: userName, email: userEmail });
      d.meetings.forEach((m) => { if (!m.kind) m.kind = "general"; if (m.projectId === undefined) m.projectId = ""; });
      ensureObligations(d);
      await saveShared("doc", d);

      if (demo) {
        setDb(d);
        const sess = loadMine("session", null);
        if (sess && d.members.some((m) => m.id === sess.memberId)) {
          setMeId(sess.memberId); setLastRead(sess.lastRead || 0); setPrefs(sess.prefs || { muted: {} });
        }
      } else {
        // Link the signed-in account to a member record by email (retry around concurrent saves).
        const emailLc = (userEmail || "").toLowerCase();
        let resolved = null;
        for (let attempt = 0; attempt < 3 && !resolved; attempt++) {
          const cur = (await loadShared("doc", null)) || d;
          resolved = cur.members.find((m) => (m.email || "").toLowerCase() === emailLc) || null;
          if (!resolved) {
            const next = clone(cur);
            const m = { id: uid(), name: userName || userEmail, email: userEmail, phone: "", role: "Prospect", status: "Applied", joined: todayStr(), ...blankMemberExtras() };
            next.members.push(m);
            next.notifications.unshift({ id: uid(), ts: Date.now(), type: "membership", title: "New application", body: `${m.name} joined with the club code and is awaiting approval.` });
            next.audit.unshift({ id: uid(), ts: Date.now(), by: m.name, action: "Application submitted", detail: "Joined via invite code" });
            try { await saveClubDoc(CLUB_ID, next); resolved = m; d = next; } catch (e) { /* conflict — retry */ }
          } else d = cur;
        }
        setDb(d);
        if (resolved) setMeId(resolved.id);
        const sess = loadMine("session", null);
        if (sess) { setLastRead(sess.lastRead || 0); setPrefs(sess.prefs || { muted: {} }); }
        registerPush(clubId).catch(() => {});
      }
      unsub = subscribeClubDoc(clubId, (doc) => setDb(doc));
      setBooted(true);
    })();
    return () => { if (unsub) unsub(); };
  }, []);

  _freshHandler = (fresh) => { setDb(fresh); setToast("Someone else just updated the club — showing the latest. Please redo your last change."); setTimeout(() => setToast(null), 4000); };
  const persist = (next) => { setDb(next); saveShared("doc", next); };
  const patch = (fn) => { const next = clone(db); const r = fn(next); if (r === false) return false; persist(next); return true; };
  const saveSession = (over = {}) => saveMine("session", { memberId: meId, lastRead, prefs, ...over });

  const me = db && meId ? db.members.find((m) => m.id === meId) : null;
  const year = db ? db.years.find((y) => y.id === db.activeYearId) : null;
  const isEBOD = me && EBOD.includes(me.role);
  const isPres = me && me.role === "President";
  const canSchedule = me && (me.role === "President" || me.role === "Vice President");
  const isSecretary = me && me.role === "Secretary";
  const isTreasurer = me && me.role === "Treasurer";
  const canFinance = me && (isTreasurer || isPres);
  const memberById = (id) => db.members.find((m) => m.id === id);
  if (db) applyTheme(db.club.colors);

  const notify = (d, { type, title, body }) => {
    d.notifications.unshift({ id: uid(), ts: Date.now(), type, title, body });
    d.notifications = d.notifications.slice(0, 100);
    if (!demo) sendPush(clubId, { type, title, body }).catch(() => {});
    else { try { if (typeof Notification !== "undefined" && Notification.permission === "granted" && !prefs.muted[type]) new Notification(title, { body }); } catch (e) {} }
  };
  const audit = (d, action, detail) => { d.audit.unshift({ id: uid(), ts: Date.now(), by: me?.name || "system", action, detail }); };
  const showToast = (msg) => { setToast(msg); setTimeout(() => setToast(null), 3000); };
  const guardLock = (dateStr) => {
    if (isLocked(db, dateStr)) { showToast(`The ${dateStr.slice(0, 7)} financial period is locked.`); return false; }
    return true;
  };

  const signIn = (id) => { setMeId(id); saveMine("session", { memberId: id, lastRead, prefs }); };
  const signOut = () => {
    if (demo) { setMeId(null); setOverlay(null); setTab("home"); saveMine("session", null); }
    else if (onSignOutAuth) onSignOutAuth();
  };
  const openNotifications = () => {
    const t = Date.now(); setLastRead(t); saveSession({ lastRead: t });
    setOverlay({ type: "notifications" });
    try { if (typeof Notification !== "undefined" && Notification.permission === "default") Notification.requestPermission(); } catch (e) {}
  };
  const setPref = (cat, muted) => { const p = { ...prefs, muted: { ...prefs.muted, [cat]: muted } }; setPrefs(p); saveSession({ prefs: p }); };

  const ctx = { db, patch, me, year, isEBOD, isPres, canSchedule, isSecretary, isTreasurer, canFinance, memberById, notify, audit, showToast, guardLock, setOverlay, overlay, setTab, signOut, prefs, setPref, demo, joinCode };

  if (!booted) return <div className="min-h-screen flex items-center justify-center" style={{ background: CRAN }}><style>{fontCss()}</style><Wheel size={90} color="rgba(255,255,255,.5)" /></div>;

  return (
    <Ctx.Provider value={ctx}>
      <div className="min-h-screen" style={{ background: PAPER, color: INK, fontFamily: BODY }}>
        <style>{fontCss()}</style>
        <div className="max-w-md mx-auto min-h-screen relative" style={{ background: PAPER }}>
          {!me ? (demo ? <LoginScreen db={db} onSignIn={signIn} persist={persist} /> : (
            <div className="min-h-screen flex items-center justify-center"><Wheel size={80} color={CRAN + "55"} /></div>
          )) : (
            <>
              <Header unread={db.notifications.filter((n) => n.ts > lastRead && !prefs.muted[n.type]).length} onBell={openNotifications} />
              <main className="px-4" style={{ paddingBottom: 96 }}>
                {tab === "home" && <HomeTab />}
                {tab === "meetings" && <MeetingsTab />}
                {tab === "projects" && <ProjectsTab />}
                {tab === "finance" && <FinanceTab />}
                {tab === "more" && <MoreTab />}
              </main>
              <TabBar tab={tab} setTab={setTab} />
              <OverlayRouter />
              {toast && <div className="fixed left-1/2 z-50 rounded-full px-5 py-2.5 font-semibold" style={{ bottom: 100, transform: "translateX(-50%)", background: INK, color: "#fff", fontSize: 13.5, boxShadow: "0 8px 24px rgba(0,0,0,.3)", maxWidth: "88%" }}>{toast}</div>}
            </>
          )}
        </div>
      </div>
    </Ctx.Provider>
  );
}

function Header({ unread, onBell }) {
  const { db, me, setTab, year } = useApp();
  return (
    <header className="relative overflow-hidden" style={{ background: `linear-gradient(135deg, ${CRAN}, ${CRAN_DK})` }}>
      <div className="absolute" style={{ right: -26, top: -30 }}><Wheel size={130} /></div>
      <div className="relative flex items-center justify-between px-4 pt-4 pb-4">
        <div className="flex items-center gap-2.5 min-w-0">
          {db.club.logo && <img src={db.club.logo} alt="" className="rounded-full object-cover" style={{ width: 38, height: 38, border: "2px solid rgba(255,255,255,.5)" }} />}
          <div className="min-w-0">
            <div className="uppercase font-bold" style={{ fontFamily: DISPLAY, fontSize: 10.5, letterSpacing: ".18em", color: "rgba(255,255,255,.75)" }}>Rotary year {year?.label}</div>
            <h1 className="font-black text-white truncate" style={{ fontFamily: DISPLAY, fontSize: 21, letterSpacing: "-.01em" }}>{db.club.name}</h1>
          </div>
        </div>
        <div className="flex items-center gap-2.5 flex-shrink-0">
          <button onClick={onBell} aria-label="Notifications" className="relative rounded-full p-2.5" style={{ background: "rgba(255,255,255,.16)" }}>
            <Icon name="bell" size={19} color="#fff" />
            {unread > 0 && <span className="absolute flex items-center justify-center rounded-full font-bold" style={{ top: 2, right: 2, minWidth: 17, height: 17, background: GOLD, color: INK, fontSize: 10.5, padding: "0 4px" }}>{unread}</span>}
          </button>
          <button onClick={() => setTab("more")} aria-label="Profile"><Avatar m={me} size={38} /></button>
        </div>
      </div>
    </header>
  );
}

function TabBar({ tab, setTab }) {
  const tabs = [
    { id: "home", icon: "home", label: "Home" }, { id: "meetings", icon: "calendar", label: "Calendar" },
    { id: "projects", icon: "target", label: "Projects" }, { id: "finance", icon: "wallet", label: "Finance" },
    { id: "more", icon: "dots", label: "More" },
  ];
  return (
    <nav className="fixed bottom-0 left-1/2 z-40 w-full max-w-md" style={{ transform: "translateX(-50%)" }}>
      <div className="flex" style={{ background: "#fff", borderTop: `1px solid ${LINE}`, paddingBottom: "env(safe-area-inset-bottom)" }}>
        {tabs.map((t) => {
          const active = tab === t.id;
          return (
            <button key={t.id} onClick={() => setTab(t.id)} className="flex-1 flex flex-col items-center gap-0.5 pt-2.5 pb-2" aria-label={t.label}>
              <Icon name={t.icon} size={21} color={active ? CRAN : "#A896A0"} strokeWidth={active ? 2.4 : 2} />
              <span className="font-bold" style={{ fontSize: 10, fontFamily: DISPLAY, color: active ? CRAN : "#A896A0" }}>{t.label}</span>
              <span className="rounded-full" style={{ width: 16, height: 3, background: active ? CRAN : "transparent" }} />
            </button>
          );
        })}
      </div>
    </nav>
  );
}

function LoginScreen({ db, onSignIn, persist }) {
  const [joining, setJoining] = useState(false);
  const [name, setName] = useState(""); const [email, setEmail] = useState("");
  const join = () => {
    if (!name.trim()) return;
    const m = { id: uid(), name: name.trim(), email: email.trim(), phone: "", role: "Prospect", status: "Applied", joined: todayStr(), ...blankMemberExtras() };
    const next = clone(db);
    next.members.push(m);
    next.notifications.unshift({ id: uid(), ts: Date.now(), type: "membership", title: "New application", body: `${m.name} applied to join the club.` });
    next.audit.unshift({ id: uid(), ts: Date.now(), by: m.name, action: "Application submitted", detail: "" });
    persist(next); onSignIn(m.id);
  };
  return (
    <div className="min-h-screen flex flex-col">
      <div className="relative overflow-hidden px-6 pt-12 pb-10" style={{ background: `linear-gradient(150deg, ${CRAN}, ${CRAN_DK})` }}>
        <div className="absolute" style={{ right: -40, top: -40 }}><Wheel size={200} /></div>
        <div className="relative">
          <div className="uppercase font-bold mb-1" style={{ fontFamily: DISPLAY, fontSize: 11, letterSpacing: ".22em", color: "rgba(255,255,255,.8)" }}>Rotaract</div>
          <h1 className="font-black text-white" style={{ fontFamily: DISPLAY, fontSize: 34, lineHeight: 1.05 }}>Club Portal</h1>
          <p className="mt-2" style={{ color: "rgba(255,255,255,.85)", fontSize: 14.5, maxWidth: 290 }}>Meetings, projects, finances, and members — organized by Rotary year.</p>
        </div>
      </div>
      <div className="flex-1 px-4 pt-5 pb-10">
        {!joining ? (
          <>
            <SectionTitle>Sign in as</SectionTitle>
            <div className="flex flex-col gap-2">
              {db.members.filter((m) => !["Resigned", "Transferred"].includes(m.status)).map((m) => (
                <Card key={m.id} onClick={() => onSignIn(m.id)} className="flex items-center gap-3 p-3">
                  <Avatar m={m} />
                  <div className="flex-1 min-w-0">
                    <div className="font-bold truncate" style={{ fontSize: 15 }}>{m.name}</div>
                    <div style={{ fontSize: 12.5, color: ROLE_COLOR[m.role] }} className="font-semibold">{m.role}{m.status !== "Active" ? ` · ${m.status}` : ""}</div>
                  </div>
                  <Icon name="chevR" size={18} color="#C9B8C1" />
                </Card>
              ))}
            </div>
            <div className="mt-4"><Btn kind="ghost" onClick={() => setJoining(true)}>I'm new — apply to join</Btn></div>
            <p className="mt-4 text-center" style={{ fontSize: 12, color: "#9A8B93" }}>Demo sign-in: pick any member to explore that role's view.</p>
          </>
        ) : (
          <>
            <SectionTitle>Apply to join</SectionTitle>
            <Card className="p-4">
              <Field label="Full name"><Input value={name} onChange={(e) => setName(e.target.value)} /></Field>
              <Field label="Email"><Input type="email" value={email} onChange={(e) => setEmail(e.target.value)} /></Field>
              <Btn onClick={join} disabled={!name.trim()}>Submit application</Btn>
              <div className="mt-2"><Btn kind="quiet" onClick={() => setJoining(false)}>Back</Btn></div>
            </Card>
          </>
        )}
      </div>
    </div>
  );
}

function OverlayRouter() {
  const { overlay, setOverlay } = useApp();
  if (!overlay) return null;
  const close = () => setOverlay(null);
  const map = {
    notifications: <NotificationsSheet onClose={close} />,
    meeting: <MeetingDetail id={overlay.id} onClose={close} />,
    project: <ProjectDetail id={overlay.id} onClose={close} />,
    members: <MembersScreen onClose={close} />,
    memberProfile: <MemberProfile id={overlay.id} onClose={close} />,
    library: <LibraryScreen onClose={close} />,
    myAccount: <MemberAccountScreen memberId={overlay.id} onClose={close} />,
    controls: <FinancialControls onClose={close} />,
    clubSettings: <ClubSettings onClose={close} />,
    yearManager: <YearManager onClose={close} />,
    notifPrefs: <NotifPrefs onClose={close} />,
  };
  return map[overlay.type] || null;
}

function NotificationsSheet({ onClose }) {
  const { db, prefs, setOverlay } = useApp();
  const iconFor = { meetings: "calendar", minutes: "doc", dues: "wallet", projects: "target", tasks: "check", announcements: "bell", library: "book", membership: "users" };
  const list = db.notifications.filter((n) => !prefs.muted[n.type]);
  return (
    <Sheet open onClose={onClose} title="Notifications" tall>
      <div className="flex justify-end mb-2"><Btn small kind="quiet" onClick={() => setOverlay({ type: "notifPrefs" })}>Preferences</Btn></div>
      {list.length === 0 ? <Empty icon="bell" title="Nothing yet" text="Club activity lands here — meetings, dues, projects, and more." /> : (
        <div className="flex flex-col gap-2">
          {list.map((n) => (
            <Card key={n.id} className="flex gap-3 p-3.5">
              <div className="rounded-xl p-2 self-start" style={{ background: CRAN + "12", color: CRAN }}><Icon name={iconFor[n.type] || "bell"} size={18} /></div>
              <div className="flex-1 min-w-0">
                <div className="font-bold" style={{ fontSize: 14 }}>{n.title}</div>
                <div style={{ fontSize: 13, color: "#6B5A64" }}>{n.body}</div>
                <div className="mt-1" style={{ fontSize: 11.5, color: "#A896A0" }}>{timeAgo(n.ts)}</div>
              </div>
            </Card>
          ))}
        </div>
      )}
    </Sheet>
  );
}

function NotifPrefs({ onClose }) {
  const { prefs, setPref } = useApp();
  return (
    <Sheet open onClose={onClose} title="Notification preferences" tall>
      <p style={{ fontSize: 13, color: "#6B5A64" }} className="mb-3">Choose what reaches you. Critical club notices stay on for everyone.</p>
      <Card className="p-2">
        {Object.entries(NOTIF_CATS).map(([id, c], i) => {
          const muted = !!prefs.muted[id];
          return (
            <div key={id} className="flex items-center gap-3 px-2 py-3" style={{ borderTop: i ? `1px solid ${LINE}` : "none" }}>
              <div className="flex-1">
                <div className="font-bold" style={{ fontSize: 14 }}>{c.label}</div>
                {c.essential && <div style={{ fontSize: 11.5, color: "#8A7580" }}>Essential — always on</div>}
              </div>
              <button onClick={() => !c.essential && setPref(id, !muted)} disabled={c.essential} aria-label={c.label}
                className="rounded-full transition-all" style={{ width: 46, height: 26, background: c.essential || !muted ? OK : "#D8CCD3", position: "relative", opacity: c.essential ? 0.55 : 1 }}>
                <span className="absolute rounded-full bg-white transition-all" style={{ width: 20, height: 20, top: 3, left: c.essential || !muted ? 23 : 3 }} />
              </button>
            </div>
          );
        })}
      </Card>
    </Sheet>
  );
}

/* ============ Home ============ */
function HomeTab() {
  const { db, me, year, isEBOD, setTab, setOverlay, memberById } = useApp();
  const t = todayStr();
  const nextMeeting = db.meetings.filter((m) => m.status === "published" && m.date >= t).sort((a, b) => a.date.localeCompare(b.date))[0];
  const acct = memberAccount(db, me.id, year.id);
  const cur = db.duesConfig.currency;
  const pendingReview = isEBOD ? db.projects.filter((p) => ["Submitted", "Under Review"].includes(p.status) && !p.votes[me.id]).length : 0;
  const myTasks = db.projects.flatMap((p) => p.status === "Approved" ? p.tasks.filter((tk) => tk.assignee === me.id && tk.status !== "Completed").map((tk) => ({ ...tk, project: p })) : []);
  const myActions = db.meetings.flatMap((m) => (m.minutes?.actionItems || []).filter((a) => a.assignee === me.id && !a.done).map((a) => ({ ...a, meeting: m })));
  const hour = new Date().getHours();
  const greet = hour < 12 ? "Good morning" : hour < 18 ? "Good afternoon" : "Good evening";

  return (
    <div>
      <div className="mt-5 mb-1">
        <h2 className="font-black" style={{ fontFamily: DISPLAY, fontSize: 25 }}>{greet}, {me.name.split(" ")[0]}</h2>
        <div className="mt-1 flex gap-1.5"><Badge color={ROLE_COLOR[me.role]}>{me.role}</Badge>{me.status !== "Active" && <Badge color={GOLD}>{me.status}</Badge>}</div>
      </div>

      <SectionTitle>Next up</SectionTitle>
      {nextMeeting ? (() => { const nk = eventKind(nextMeeting); return (
        <Card onClick={() => setOverlay({ type: "meeting", id: nextMeeting.id })} className="relative overflow-hidden p-4" style={{ background: `linear-gradient(135deg, ${nk.color}, ${shade(nk.color, -0.28)})`, border: "none" }}>
          <div className="absolute" style={{ right: -20, bottom: -30 }}><Wheel size={110} /></div>
          <div className="relative text-white">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="font-black" style={{ fontFamily: DISPLAY, fontSize: 20 }}>{nextMeeting.title}</span>
              <span className="rounded-full font-semibold uppercase" style={{ fontSize: 10, padding: "2px 8px", background: "rgba(255,255,255,.22)" }}>{nk.short} · {nextMeeting.type}</span>
            </div>
            <div className="mt-1.5 flex items-center gap-1.5" style={{ fontSize: 13.5, color: "rgba(255,255,255,.9)" }}><Icon name="calendar" size={15} /> {fmtDate(nextMeeting.date)} · {nextMeeting.time}</div>
            <div className="mt-1 flex items-center gap-1.5" style={{ fontSize: 13.5, color: "rgba(255,255,255,.9)" }}><Icon name={nextMeeting.type === "virtual" ? "video" : "pin"} size={15} /> {nextMeeting.type === "virtual" ? "Online meeting" : nextMeeting.location}</div>
            <div className="mt-3"><span className="inline-flex rounded-full font-semibold" style={{ background: "rgba(255,255,255,.18)", color: "#fff", fontSize: 12, padding: "4px 11px" }}>{nextMeeting.rsvps?.[me.id] === "yes" ? "You're going ✓" : "Tap to view & RSVP →"}</span></div>
          </div>
        </Card>
      ); })() : <Empty icon="calendar" title="Nothing scheduled" text="Meetings, socials, fundraisers, and project dates appear here." />}

      <div className="grid grid-cols-2 gap-2.5 mt-4">
        <Card onClick={() => setOverlay({ type: "myAccount", id: me.id })} className="p-3.5">
          <div style={{ color: acct.balance > 0 ? (acct.overdue > 0 ? BAD : GOLD) : OK }}><Icon name="wallet" size={20} /></div>
          <div className="font-extrabold mt-1.5" style={{ fontFamily: DISPLAY, fontSize: 16 }}>{money(Math.max(0, acct.balance), cur)}</div>
          <div style={{ fontSize: 12, color: acct.overdue > 0 ? BAD : "#8A7580" }}>{acct.overdue > 0 ? `${money(acct.overdue, cur)} overdue` : acct.balance > 0 ? `due ${fmtShort(acct.nextDue)}` : "dues up to date"}</div>
        </Card>
        <Card onClick={() => setTab("projects")} className="p-3.5">
          <div style={{ color: AZURE }}><Icon name="target" size={20} /></div>
          <div className="font-extrabold mt-1.5" style={{ fontFamily: DISPLAY, fontSize: 16 }}>{db.projects.filter((p) => p.status === "Approved").length} active</div>
          <div style={{ fontSize: 12, color: "#8A7580" }}>{db.projects.filter((p) => ["Submitted", "Under Review"].includes(p.status)).length} in review</div>
        </Card>
      </div>

      {pendingReview > 0 && (
        <Card onClick={() => setTab("projects")} className="mt-2.5 flex items-center gap-3 p-3.5" style={{ borderColor: GOLD, background: GOLD + "14" }}>
          <div style={{ color: "#B07400" }}><Icon name="shield" size={20} /></div>
          <div className="flex-1"><div className="font-bold" style={{ fontSize: 14 }}>EBOD review needed</div><div style={{ fontSize: 12.5, color: "#7A6570" }}>{pendingReview} proposal{pendingReview === 1 ? "" : "s"} awaiting your vote</div></div>
          <Icon name="chevR" size={18} color="#B07400" />
        </Card>
      )}

      {(myTasks.length > 0 || myActions.length > 0) && (
        <>
          <SectionTitle>My open items</SectionTitle>
          <div className="flex flex-col gap-2">
            {myTasks.slice(0, 3).map((tk) => (
              <Card key={tk.id} onClick={() => setOverlay({ type: "project", id: tk.project.id })} className="flex items-center gap-3 p-3.5">
                <Badge color={PRIORITY_COLOR[tk.priority]}>{tk.priority}</Badge>
                <div className="flex-1 min-w-0">
                  <div className="font-bold truncate" style={{ fontSize: 14 }}>{tk.title}</div>
                  <div style={{ fontSize: 12, color: tk.deadline && tk.deadline < t ? BAD : "#8A7580" }}>{tk.project.title}{tk.deadline ? ` · due ${fmtShort(tk.deadline)}` : ""}</div>
                </div>
                <Badge color={TASK_COLOR[tk.status]}>{tk.status}</Badge>
              </Card>
            ))}
            {myActions.slice(0, 2).map((a) => (
              <Card key={a.id} onClick={() => setOverlay({ type: "meeting", id: a.meeting.id })} className="flex items-center gap-3 p-3.5">
                <div style={{ color: CRAN }}><Icon name="flag" size={17} /></div>
                <div className="flex-1 min-w-0">
                  <div className="font-bold truncate" style={{ fontSize: 14 }}>{a.text}</div>
                  <div style={{ fontSize: 12, color: "#8A7580" }}>Action item · {a.meeting.title}{a.due ? ` · due ${fmtShort(a.due)}` : ""}</div>
                </div>
              </Card>
            ))}
          </div>
        </>
      )}

      <SectionTitle action={<button onClick={() => setOverlay({ type: "library" })} className="font-bold" style={{ fontSize: 12.5, color: CRAN }}>Library →</button>}>Latest updates</SectionTitle>
      <div className="flex flex-col gap-2">
        {db.notifications.slice(0, 3).map((n) => (
          <Card key={n.id} className="p-3.5"><div className="font-bold" style={{ fontSize: 13.5 }}>{n.title}</div><div style={{ fontSize: 12.5, color: "#6B5A64" }}>{n.body}</div></Card>
        ))}
      </div>
    </div>
  );
}

/* ============ Meetings ============ */
function MeetingsTab() {
  const { db, canSchedule, me } = useApp();
  const [editing, setEditing] = useState(null); // null | 'new' | meeting
  const [kindFilter, setKindFilter] = useState("all");
  const [cursor, setCursor] = useState(() => todayStr().slice(0, 7)); // 'YYYY-MM'
  const [selDate, setSelDate] = useState(null);
  const t = todayStr();
  const vis = db.meetings.filter((m) => (m.status !== "draft" || canSchedule || m.createdBy === me.id) && (kindFilter === "all" || (m.kind || "general") === kindFilter));
  const byDate = {};
  vis.forEach((m) => { (byDate[m.date] = byDate[m.date] || []).push(m); });

  const [cy, cm] = cursor.split("-").map(Number);
  const daysInMonth = new Date(cy, cm, 0).getDate();
  const lead = new Date(cy, cm - 1, 1).getDay();
  const cells = [];
  for (let i = 0; i < lead; i++) cells.push(null);
  for (let dd = 1; dd <= daysInMonth; dd++) cells.push(`${cy}-${pad2(cm)}-${pad2(dd)}`);
  const moveMonth = (dir) => {
    let y = cy, m = cm + dir;
    if (m === 0) { m = 12; y--; } if (m === 13) { m = 1; y++; }
    setCursor(`${y}-${pad2(m)}`); setSelDate(null);
  };

  const upcoming = vis.filter((m) => m.date >= t && m.status !== "cancelled").sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time));
  const past = vis.filter((m) => m.date < t || m.status === "cancelled").sort((a, b) => b.date.localeCompare(a.date));
  const selEvents = selDate ? (byDate[selDate] || []).sort((a, b) => a.time.localeCompare(b.time)) : null;

  return (
    <div>
      <div className="flex items-center justify-between mt-5">
        <h2 className="font-black" style={{ fontFamily: DISPLAY, fontSize: 24 }}>Calendar</h2>
        {canSchedule && <Btn small onClick={() => setEditing("new")}><span className="flex items-center gap-1.5"><Icon name="plus" size={15} /> New event</span></Btn>}
      </div>

      <div className="flex gap-1.5 overflow-x-auto pb-1 mt-3">
        <button onClick={() => setKindFilter("all")} className="rounded-full font-bold flex-shrink-0" style={{ fontSize: 12, padding: "6px 12px", fontFamily: DISPLAY, background: kindFilter === "all" ? INK : "#fff", color: kindFilter === "all" ? "#fff" : "#6B5A64", border: `1px solid ${kindFilter === "all" ? INK : LINE}` }}>All</button>
        {Object.entries(EVENT_KINDS).map(([k, ek]) => (
          <button key={k} onClick={() => setKindFilter(kindFilter === k ? "all" : k)} className="rounded-full font-bold flex-shrink-0 flex items-center gap-1.5" style={{ fontSize: 12, padding: "6px 12px", fontFamily: DISPLAY, background: kindFilter === k ? ek.color : "#fff", color: kindFilter === k ? "#fff" : "#6B5A64", border: `1px solid ${kindFilter === k ? ek.color : LINE}` }}>
            <span className="rounded-full" style={{ width: 7, height: 7, background: kindFilter === k ? "#fff" : ek.color }} />{ek.short}
          </button>
        ))}
      </div>

      <Card className="p-3 mt-2.5">
        <div className="flex items-center justify-between mb-2">
          <button onClick={() => moveMonth(-1)} className="p-1.5 rounded-full" aria-label="Previous month" style={{ background: PAPER }}><Icon name="chevL" size={16} /></button>
          <div className="font-extrabold" style={{ fontFamily: DISPLAY, fontSize: 15.5 }}>{MONTHS[cm - 1]} {cy}</div>
          <button onClick={() => moveMonth(1)} className="p-1.5 rounded-full" aria-label="Next month" style={{ background: PAPER }}><Icon name="chevR" size={16} /></button>
        </div>
        <div className="grid grid-cols-7 mb-1">{["S", "M", "T", "W", "T", "F", "S"].map((d, i) => <div key={i} className="text-center font-bold" style={{ fontSize: 10, color: "#A896A0" }}>{d}</div>)}</div>
        <div className="grid grid-cols-7" style={{ rowGap: 2 }}>
          {cells.map((ds, i) => {
            if (!ds) return <div key={"e" + i} />;
            const evs = byDate[ds] || [];
            const isToday = ds === t, isSel = ds === selDate;
            return (
              <button key={ds} onClick={() => setSelDate(isSel ? null : ds)} className="flex flex-col items-center rounded-xl pt-1" style={{ background: isSel ? INK : isToday ? CRAN + "14" : "transparent", paddingBottom: 3 }} aria-label={fmtDate(ds)}>
                <span className="font-semibold" style={{ fontSize: 12.5, color: isSel ? "#fff" : isToday ? CRAN : INK }}>{Number(ds.slice(8))}</span>
                <span className="flex gap-0.5" style={{ height: 5 }}>
                  {evs.slice(0, 3).map((ev) => <span key={ev.id} className="rounded-full" style={{ width: 4.5, height: 4.5, background: ev.status === "cancelled" ? "#C9B8C1" : eventKind(ev).color }} />)}
                </span>
              </button>
            );
          })}
        </div>
      </Card>

      {selDate ? (
        <>
          <SectionTitle action={<button onClick={() => setSelDate(null)} className="font-bold" style={{ fontSize: 12.5, color: CRAN }}>Clear</button>}>{fmtDate(selDate)}</SectionTitle>
          {selEvents.length ? <div className="flex flex-col gap-2">{selEvents.map((m) => <EventCard key={m.id} m={m} />)}</div>
            : <Empty icon="calendar" title="Nothing on this date" text={canSchedule ? "Tap New event to schedule something." : ""} />}
        </>
      ) : (
        <>
          <SectionTitle>Upcoming</SectionTitle>
          {upcoming.length ? <div className="flex flex-col gap-2">{upcoming.map((m) => <EventCard key={m.id} m={m} />)}</div> : <Empty icon="calendar" title="Nothing scheduled" text="Meetings, socials, fundraisers, and project dates all live here." />}
          <SectionTitle>Past & cancelled</SectionTitle>
          {past.length ? <div className="flex flex-col gap-2">{past.slice(0, 12).map((m) => <EventCard key={m.id} m={m} />)}</div> : <Empty icon="clock" title="No past events yet" />}
        </>
      )}
      {editing && <MeetingForm meeting={editing === "new" ? null : editing} onClose={() => setEditing(null)} />}
    </div>
  );
}

function EventCard({ m }) {
  const { setOverlay } = useApp();
  const t = todayStr();
  const d = new Date(m.date + "T12:00:00");
  const ek = eventKind(m);
  const TYPE_ICON = { physical: "pin", virtual: "video", hybrid: "globe" };
  const present = Object.values(m.attendance || {}).filter((v) => ["present", "late", "virtual"].includes(v)).length;
  const going = Object.values(m.rsvps || {}).filter((r) => r === "yes").length;
  const minutesPending = !m.minutes && m.date < t && m.status === "published" && ["general", "professional", "orientation"].includes(m.kind || "general");
  return (
    <Card onClick={() => setOverlay({ type: "meeting", id: m.id })} className="flex gap-3 p-3.5" style={m.status === "cancelled" ? { opacity: 0.6 } : {}}>
      <div className="flex flex-col items-center justify-center rounded-xl flex-shrink-0" style={{ width: 52, height: 56, background: ek.color + "12" }}>
        <div className="font-black" style={{ fontFamily: DISPLAY, fontSize: 19, color: ek.color, lineHeight: 1 }}>{d.getDate()}</div>
        <div className="font-bold uppercase" style={{ fontSize: 10, color: ek.color, letterSpacing: ".08em" }}>{MONTHS[d.getMonth()]}</div>
      </div>
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-1.5"><span className="font-extrabold truncate" style={{ fontFamily: DISPLAY, fontSize: 15.5 }}>{m.title}</span><Icon name={TYPE_ICON[m.type]} size={14} color={AZURE} /></div>
        <div style={{ fontSize: 12.5, color: "#8A7580" }}>{m.time} · {m.type === "virtual" ? "online" : m.location || ek.label}</div>
        <div className="mt-1.5 flex gap-1.5 flex-wrap">
          <Badge color={ek.color}>{ek.short}</Badge>
          {m.status === "draft" && <Badge color={GOLD}>Draft</Badge>}
          {m.status === "cancelled" && <Badge color={BAD}>Cancelled</Badge>}
          {m.minutes ? <Badge color={OK}>Minutes ✓</Badge> : minutesPending ? <Badge color={GOLD}>Minutes pending</Badge> : null}
          {present > 0 && <Badge>{present} attended</Badge>}
          {m.date >= t && m.status === "published" && going > 0 && <Badge color={AZURE}>{going} going</Badge>}
        </div>
      </div>
      <Icon name="chevR" size={18} color="#C9B8C1" />
    </Card>
  );
}

function MeetingForm({ meeting, onClose }) {
  const { db, patch, me, year, notify, audit, showToast, memberById } = useApp();
  const [f, setF] = useState(meeting ? { kind: "general", projectId: "", ...clone(meeting) } : {
    kind: "general", projectId: "", title: "General Meeting", date: "", time: "19:00", type: "physical", location: "", link: "",
    agenda: [""], presenters: [], rsvpDeadline: "", attachments: [],
  });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const activeMembers = db.members.filter((m) => ACTIVE_LIKE.includes(m.status));
  const togglePresenter = (id) => setF({ ...f, presenters: f.presenters.includes(id) ? f.presenters.filter((x) => x !== id) : [...f.presenters, id] });
  const save = (publish) => {
    if (!f.title.trim() || !f.date) return;
    const agenda = f.agenda.map((a) => a.trim()).filter(Boolean);
    const kl = (EVENT_KINDS[f.kind] || EVENT_KINDS.general).label;
    const rescheduled = meeting && meeting.status === "published" && (meeting.date !== f.date || meeting.time !== f.time);
    patch((d) => {
      if (meeting) {
        const m = d.meetings.find((x) => x.id === meeting.id);
        Object.assign(m, { ...f, agenda });
        if (publish && m.status === "draft") { m.status = "published"; m.publishedAt = Date.now(); }
        if (m.status === "published") {
          notify(d, { type: "meetings", title: rescheduled ? `${kl} rescheduled` : `${kl} updated`, body: `${m.title} — ${fmtDate(m.date)} at ${m.time}. ${agenda.length} agenda items.` });
        }
        audit(d, rescheduled ? "Event rescheduled" : "Event updated", m.title);
      } else {
        const m = { id: uid(), yearId: year.id, ...f, agenda, status: publish ? "published" : "draft", rsvps: {}, attendance: {}, guests: [], minutes: null, createdBy: me.id, publishedAt: publish ? Date.now() : null };
        d.meetings.push(m);
        if (publish) notify(d, { type: "meetings", title: `${kl} published`, body: `${m.title} — ${fmtDate(m.date)} at ${m.time}. Please RSVP${f.rsvpDeadline ? ` by ${fmtShort(f.rsvpDeadline)}` : ""}.` });
        audit(d, publish ? "Event published" : "Event drafted", m.title);
      }
    });
    showToast(publish ? "Published — members notified" : "Saved as draft");
    onClose();
  };
  return (
    <Sheet open onClose={onClose} title={meeting ? "Edit event" : "New event"} tall>
      <Field label="Event type">
        <div className="flex gap-1.5 flex-wrap">
          {Object.entries(EVENT_KINDS).map(([k, ek]) => (
            <button key={k} onClick={() => setF((prev) => ({ ...prev, kind: k, title: !prev.title.trim() || Object.values(EVENT_KINDS).some((x) => x.label === prev.title) ? ek.label : prev.title }))}
              className="rounded-full font-bold" style={{ fontSize: 12, padding: "6px 12px", fontFamily: DISPLAY, background: f.kind === k ? ek.color : "#fff", color: f.kind === k ? "#fff" : "#6B5A64", border: `1px solid ${f.kind === k ? ek.color : LINE}` }}>{ek.short}</button>
          ))}
        </div>
      </Field>
      <Field label="Title"><Input value={f.title} onChange={set("title")} /></Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Date"><Input type="date" value={f.date} onChange={set("date")} /></Field>
        <Field label="Time"><Input type="time" value={f.time} onChange={set("time")} /></Field>
      </div>
      <Field label="Format">
        <div className="flex gap-2">
          {["physical", "virtual", "hybrid"].map((tp) => (
            <button key={tp} onClick={() => setF({ ...f, type: tp })} className="flex-1 rounded-xl font-bold py-2.5 capitalize"
              style={{ fontFamily: DISPLAY, fontSize: 13, background: f.type === tp ? AZURE : "#fff", color: f.type === tp ? "#fff" : "#6B5A64", border: `1.5px solid ${f.type === tp ? "transparent" : LINE}` }}>{tp}</button>
          ))}
        </div>
      </Field>
      {f.type !== "virtual" && <Field label="Location"><Input value={f.location} onChange={set("location")} placeholder="e.g. Community Center, Room B" /></Field>}
      {f.type !== "physical" && <Field label="Meeting link"><Input value={f.link} onChange={set("link")} placeholder="https://…" /></Field>}
      <Field label="RSVP deadline"><Input type="date" value={f.rsvpDeadline} onChange={set("rsvpDeadline")} /></Field>
      {["project", "fundraiser"].includes(f.kind) && (
        <Field label="Related project" hint="Links this date to the project workspace.">
          <Select value={f.projectId} onChange={set("projectId")}>
            <option value="">None</option>
            {db.projects.filter((p) => ["Approved", "Pending Close"].includes(p.status)).map((p) => <option key={p.id} value={p.id}>{p.title}</option>)}
          </Select>
        </Field>
      )}
      <Field label="Agenda">
        <div className="flex flex-col gap-2">
          {f.agenda.map((a, i) => (
            <div key={i} className="flex gap-2">
              <Input value={a} onChange={(e) => setF({ ...f, agenda: f.agenda.map((x, j) => j === i ? e.target.value : x) })} placeholder={`Item ${i + 1}`} />
              {f.agenda.length > 1 && <button onClick={() => setF({ ...f, agenda: f.agenda.filter((_, j) => j !== i) })} aria-label="Remove"><Icon name="x" size={16} color="#A896A0" /></button>}
            </div>
          ))}
        </div>
        <div className="mt-2"><Btn small kind="ghost" onClick={() => setF({ ...f, agenda: [...f.agenda, ""] })}>+ Item</Btn></div>
      </Field>
      <Field label="Expected presenters">
        <div className="flex flex-wrap gap-1.5">
          {activeMembers.map((m) => (
            <button key={m.id} onClick={() => togglePresenter(m.id)} className="rounded-full font-semibold"
              style={{ fontSize: 12, padding: "5px 11px", background: f.presenters.includes(m.id) ? AZURE : "#fff", color: f.presenters.includes(m.id) ? "#fff" : "#6B5A64", border: `1px solid ${f.presenters.includes(m.id) ? AZURE : LINE}` }}>
              {m.name.split(" ")[0]}
            </button>
          ))}
        </div>
      </Field>
      <FilePick label="Attach documents" accept="*/*" hint={`Agenda packets, reports — under ${MAX_FILE_KB} KB each.`} onFile={(file) => setF({ ...f, attachments: [...f.attachments, file] })} />
      {f.attachments.length > 0 && <div className="mb-4">{f.attachments.map((a) => <FileChip key={a.id} f={a} onRemove={() => setF({ ...f, attachments: f.attachments.filter((x) => x.id !== a.id) })} />)}</div>}
      <div className="flex flex-col gap-2">
        <Btn onClick={() => save(true)} disabled={!f.title.trim() || !f.date}>{meeting?.status === "published" ? "Save & notify members" : "Publish & notify members"}</Btn>
        {(!meeting || meeting.status === "draft") && <Btn kind="quiet" onClick={() => save(false)} disabled={!f.title.trim() || !f.date}>Save as draft</Btn>}
      </div>
    </Sheet>
  );
}

function MeetingDetail({ id, onClose }) {
  const { db, me, patch, isSecretary, canSchedule, memberById, notify, audit, showToast, setOverlay } = useApp();
  const m = db.meetings.find((x) => x.id === id);
  const [editAtt, setEditAtt] = useState(false);
  const [editMeeting, setEditMeeting] = useState(false);
  const [minutesOpen, setMinutesOpen] = useState(false);
  const [guestName, setGuestName] = useState("");
  const [guestKind, setGuestKind] = useState("guest");
  if (!m) return null;
  const t = todayStr();
  const activeMembers = db.members.filter((x) => ACTIVE_LIKE.includes(x.status));
  const att = m.attendance || {};
  const counts = {}; ATT_STATUSES.forEach((s) => (counts[s] = 0)); Object.values(att).forEach((v) => { if (counts[v] !== undefined) counts[v]++; });
  const rsvpOpen = m.status === "published" && m.date >= t && (!m.rsvpDeadline || m.rsvpDeadline >= t);
  const myRsvp = m.rsvps?.[me.id];

  const setAtt = (mid, val) => patch((d) => {
    const mt = d.meetings.find((x) => x.id === id);
    if (val) mt.attendance[mid] = val; else delete mt.attendance[mid];
  });
  const rsvp = (v) => patch((d) => { const mt = d.meetings.find((x) => x.id === id); mt.rsvps = mt.rsvps || {}; mt.rsvps[me.id] = v; });
  const cancelMeeting = () => {
    patch((d) => {
      const mt = d.meetings.find((x) => x.id === id);
      mt.status = "cancelled";
      notify(d, { type: "meetings", title: `${eventKind(mt).label} cancelled`, body: `${mt.title} on ${fmtDate(mt.date)} has been cancelled.` });
      audit(d, "Event cancelled", mt.title);
    });
    showToast("Cancelled — members notified");
  };
  const addGuest = () => {
    if (!guestName.trim()) return;
    patch((d) => { d.meetings.find((x) => x.id === id).guests.push({ id: uid(), name: guestName.trim(), kind: guestKind }); });
    setGuestName("");
  };
  const toggleAction = (aid) => patch((d) => { const a = d.meetings.find((x) => x.id === id).minutes.actionItems.find((x) => x.id === aid); a.done = !a.done; });

  return (
    <FullScreen title={m.title} onClose={onClose}
      right={canSchedule && m.status !== "cancelled" ? <button onClick={() => setEditMeeting(true)} aria-label="Edit" className="rounded-full p-1.5" style={{ background: "rgba(255,255,255,.18)" }}><Icon name="edit" size={17} color="#fff" /></button> : null}>
      {m.status === "cancelled" && <Card className="p-3 mb-3 text-center font-bold" style={{ background: BAD + "12", borderColor: BAD + "44", color: BAD }}>This meeting was cancelled.</Card>}
      {m.status === "draft" && <Card className="p-3 mb-3 text-center font-bold" style={{ background: GOLD + "14", borderColor: GOLD, color: "#8A5A00" }}>Draft — not yet visible to members.</Card>}
      <Card className="p-4">
        <div className="flex items-center gap-2" style={{ fontSize: 14 }}><Icon name="calendar" size={16} color={CRAN} /> <b>{fmtDate(m.date)}</b> · {m.time}</div>
        {m.type !== "virtual" && <div className="flex items-center gap-2 mt-1.5" style={{ fontSize: 14 }}><Icon name="pin" size={16} color={CRAN} /> {m.location || "Location TBA"}</div>}
        {m.type !== "physical" && m.link && <div className="flex items-center gap-2 mt-1.5" style={{ fontSize: 14 }}><Icon name="video" size={16} color={AZURE} /> <a href={m.link} target="_blank" rel="noreferrer" style={{ color: AZURE, fontWeight: 600 }}>Join online</a></div>}
        <div className="mt-2 flex gap-1.5 flex-wrap items-center">
          <Badge color={eventKind(m).color} filled>{eventKind(m).label}</Badge>
          <Badge color={AZURE}>{m.type}</Badge>
          {m.rsvpDeadline && <Badge color={GOLD}>RSVP by {fmtShort(m.rsvpDeadline)}</Badge>}
          {m.presenters?.length > 0 && <Badge>Presenters: {m.presenters.map((p) => memberById(p)?.name.split(" ")[0]).join(", ")}</Badge>}
          {m.projectId && db.projects.some((p) => p.id === m.projectId) && (
            <button onClick={() => setOverlay({ type: "project", id: m.projectId })}><Badge color="#B34700">↗ {db.projects.find((p) => p.id === m.projectId).title}</Badge></button>
          )}
        </div>
        {m.attachments?.length > 0 && <div className="mt-3 pt-3" style={{ borderTop: `1px solid ${LINE}` }}>{m.attachments.map((a) => <FileChip key={a.id} f={a} />)}</div>}
      </Card>

      {rsvpOpen && (
        <>
          <SectionTitle>RSVP · {Object.values(m.rsvps || {}).filter((r) => r === "yes").length} going</SectionTitle>
          <div className="flex gap-2">
            <Btn kind={myRsvp === "yes" ? "ok" : "quiet"} onClick={() => rsvp("yes")} style={{ flex: 1 }}>Going{myRsvp === "yes" ? " ✓" : ""}</Btn>
            <Btn kind={myRsvp === "no" ? "danger" : "quiet"} onClick={() => rsvp("no")} style={{ flex: 1 }}>Can't make it{myRsvp === "no" ? " ✓" : ""}</Btn>
          </div>
        </>
      )}

      <SectionTitle>Agenda</SectionTitle>
      <Card className="p-4">
        {m.agenda.length ? m.agenda.map((a, i) => (
          <div key={i} className="flex gap-3 py-1.5" style={{ borderTop: i ? `1px solid ${LINE}` : "none" }}>
            <span className="font-black" style={{ fontFamily: DISPLAY, color: CRAN, fontSize: 14, minWidth: 22 }}>{String(i + 1).padStart(2, "0")}</span>
            <span style={{ fontSize: 14.5 }}>{a}</span>
          </div>
        )) : <span style={{ fontSize: 13.5, color: "#8A7580" }}>No agenda items.</span>}
      </Card>

      <SectionTitle action={isSecretary && m.status === "published" ? <Btn small kind={editAtt ? "dark" : "ghost"} onClick={() => setEditAtt(!editAtt)}>{editAtt ? "Done" : "Record"}</Btn> : null}>Attendance</SectionTitle>
      {Object.keys(att).length > 0 || editAtt ? (
        <>
          <div className="flex gap-1.5 mb-2 flex-wrap">{ATT_STATUSES.map((s) => counts[s] > 0 && <Badge key={s} color={ATT_COLOR[s]}>{counts[s]} {s}</Badge>)}</div>
          <Card className="p-2">
            {activeMembers.map((mb, i) => (
              <div key={mb.id} className="px-2 py-2" style={{ borderTop: i ? `1px solid ${LINE}` : "none" }}>
                <div className="flex items-center gap-2.5">
                  <Avatar m={mb} size={30} />
                  <div className="flex-1 font-semibold" style={{ fontSize: 13.5 }}>{mb.name}</div>
                  {!editAtt && (att[mb.id] ? <Badge color={ATT_COLOR[att[mb.id]]}>{att[mb.id]}</Badge> : <span style={{ fontSize: 12, color: "#C9B8C1" }}>—</span>)}
                </div>
                {editAtt && (
                  <div className="flex gap-1 mt-1.5 ml-10 flex-wrap">
                    {ATT_STATUSES.map((s) => (
                      <button key={s} onClick={() => setAtt(mb.id, att[mb.id] === s ? null : s)} className="rounded-full font-bold capitalize"
                        style={{ fontSize: 10.5, padding: "4px 9px", background: att[mb.id] === s ? ATT_COLOR[s] : "#F0EAEE", color: att[mb.id] === s ? "#fff" : "#8A7580" }}>{s}</button>
                    ))}
                  </div>
                )}
              </div>
            ))}
          </Card>
          <SectionTitle>Guests & prospects present</SectionTitle>
          <Card className="p-3">
            {m.guests.length ? m.guests.map((g) => (
              <div key={g.id} className="flex items-center gap-2 py-1" style={{ fontSize: 13.5 }}>
                <Icon name="users" size={14} color="#8A7580" /> {g.name} <Badge color={g.kind === "prospect" ? GOLD : AZURE}>{g.kind}</Badge>
              </div>
            )) : <span style={{ fontSize: 13, color: "#9A8B93" }}>No guests recorded.</span>}
            {isSecretary && (
              <div className="flex gap-2 mt-2.5">
                <Input value={guestName} onChange={(e) => setGuestName(e.target.value)} placeholder="Guest name" style={{ flex: 1 }} />
                <Select value={guestKind} onChange={(e) => setGuestKind(e.target.value)} style={{ width: 110 }}><option value="guest">Guest</option><option value="prospect">Prospect</option></Select>
                <Btn small onClick={addGuest}>Add</Btn>
              </div>
            )}
          </Card>
        </>
      ) : <Empty icon="users" title="Not recorded yet" text={isSecretary ? "Tap Record when the meeting starts." : "The secretary records attendance."} />}

      <SectionTitle action={isSecretary ? <Btn small kind="ghost" onClick={() => setMinutesOpen(true)}>{m.minutes ? "Revise / actions" : "Submit minutes"}</Btn> : null}>Minutes</SectionTitle>
      {m.minutes ? (
        <>
          {(() => { const v = m.minutes.versions[m.minutes.versions.length - 1]; return (
            <Card className="p-4">
              {v.dataUrl ? (
                <button onClick={() => downloadDataUrl(v.fileName, v.dataUrl)} className="flex items-center gap-2 font-bold" style={{ color: AZURE, fontSize: 14.5 }}>
                  <Icon name="doc" size={18} /> {v.fileName} <Icon name="download" size={15} />
                </button>
              ) : <div style={{ fontSize: 14.5, whiteSpace: "pre-wrap", lineHeight: 1.55 }}>{v.text}</div>}
              <div className="mt-3 pt-3" style={{ borderTop: `1px solid ${LINE}`, fontSize: 12.5, color: "#8A7580" }}>
                Version {m.minutes.versions.length} · by {memberById(v.by)?.name} · {timeAgo(v.at)}{v.note ? ` · "${v.note}"` : ""}
              </div>
              {m.minutes.versions.length > 1 && (
                <details className="mt-2"><summary className="font-bold cursor-pointer" style={{ fontSize: 12.5, color: CRAN }}>Revision history ({m.minutes.versions.length} versions)</summary>
                  {m.minutes.versions.slice(0, -1).reverse().map((pv, i) => (
                    <div key={pv.id} className="mt-2 pt-2" style={{ borderTop: `1px dashed ${LINE}`, fontSize: 12.5, color: "#6B5A64" }}>
                      <b>v{m.minutes.versions.length - 1 - i}</b> · {memberById(pv.by)?.name} · {timeAgo(pv.at)}{pv.note ? ` · ${pv.note}` : ""}
                      {pv.dataUrl ? <button onClick={() => downloadDataUrl(pv.fileName, pv.dataUrl)} className="ml-2 font-semibold" style={{ color: AZURE }}>download</button> : <div className="mt-1" style={{ whiteSpace: "pre-wrap" }}>{pv.text}</div>}
                    </div>
                  ))}
                </details>
              )}
            </Card>
          ); })()}
          {m.minutes.actionItems?.length > 0 && (
            <>
              <SectionTitle>Action items</SectionTitle>
              <Card className="p-2">
                {m.minutes.actionItems.map((a, i) => (
                  <div key={a.id} className="flex items-center gap-3 px-2 py-2" style={{ borderTop: i ? `1px solid ${LINE}` : "none" }}>
                    <button onClick={() => (isSecretary || a.assignee === me.id) && toggleAction(a.id)} className="rounded-full flex items-center justify-center flex-shrink-0"
                      style={{ width: 24, height: 24, border: `2px solid ${a.done ? OK : "#C9B8C1"}`, background: a.done ? OK : "transparent" }}>
                      {a.done && <Icon name="check" size={13} color="#fff" />}
                    </button>
                    <div className="flex-1">
                      <div className="font-semibold" style={{ fontSize: 13.5, textDecoration: a.done ? "line-through" : "none", color: a.done ? "#A896A0" : INK }}>{a.text}</div>
                      <div style={{ fontSize: 11.5, color: "#8A7580" }}>{memberById(a.assignee)?.name}{a.due ? ` · due ${fmtShort(a.due)}` : ""}</div>
                    </div>
                  </div>
                ))}
              </Card>
            </>
          )}
        </>
      ) : <Empty icon="doc" title="Minutes not posted" text="The secretary uploads the minutes PDF after the meeting; corrections keep a revision history." />}

      {canSchedule && m.status === "published" && m.date >= t && (
        <div className="mt-5"><Btn kind="danger" onClick={cancelMeeting}>Cancel this event & notify</Btn></div>
      )}
      {editMeeting && <MeetingForm meeting={m} onClose={() => setEditMeeting(false)} />}
      {minutesOpen && <MinutesSheet meeting={m} onClose={() => setMinutesOpen(false)} />}
    </FullScreen>
  );
}

function MinutesSheet({ meeting, onClose }) {
  const { db, patch, me, notify, audit, showToast } = useApp();
  const [file, setFile] = useState(null);
  const [text, setText] = useState("");
  const [note, setNote] = useState(meeting.minutes ? "" : "Original");
  const [items, setItems] = useState(meeting.minutes?.actionItems ? clone(meeting.minutes.actionItems) : []);
  const [ni, setNi] = useState({ text: "", assignee: "", due: "" });
  const activeMembers = db.members.filter((x) => ACTIVE_LIKE.includes(x.status));
  const isRevision = !!meeting.minutes;

  const addItem = () => { if (!ni.text.trim() || !ni.assignee) return; setItems([...items, { id: uid(), ...ni, text: ni.text.trim(), done: false }]); setNi({ text: "", assignee: "", due: "" }); };
  const submit = () => {
    if (!file && !text.trim() && !isRevision) return;
    patch((d) => {
      const mt = d.meetings.find((x) => x.id === meeting.id);
      if (!mt.minutes) mt.minutes = { versions: [], actionItems: [] };
      if (file || text.trim()) {
        mt.minutes.versions.push({ id: uid(), fileName: file ? file.name : "Minutes (text)", dataUrl: file ? file.dataUrl : null, text: file ? "" : text.trim(), by: me.id, at: Date.now(), note: note.trim() || (isRevision ? "Revision" : "Original") });
      }
      const newAssignees = items.filter((it) => !(meeting.minutes?.actionItems || []).some((old) => old.id === it.id));
      mt.minutes.actionItems = items;
      if (!isRevision) notify(d, { type: "minutes", title: "Minutes published", body: `Minutes for ${mt.title} (${fmtShort(mt.date)}) are available to all members.` });
      else if (file || text.trim()) notify(d, { type: "minutes", title: "Minutes revised", body: `A correction to ${mt.title} minutes was posted — the original remains in the revision history.` });
      newAssignees.forEach((it) => notify(d, { type: "tasks", title: "Action item assigned", body: `${d.members.find((x) => x.id === it.assignee)?.name}: "${it.text}"${it.due ? `, due ${fmtShort(it.due)}` : ""}.` }));
      audit(d, isRevision ? "Minutes revised" : "Minutes published", mt.title);
    });
    showToast(isRevision ? "Revision saved — history preserved" : "Minutes published");
    onClose();
  };
  return (
    <Sheet open onClose={onClose} title={isRevision ? "Revise minutes / action items" : "Submit minutes"} tall>
      {isRevision && <p className="mb-3" style={{ fontSize: 13, color: "#6B5A64" }}>Uploading a new file creates a new version. The original stays in the revision history — nothing is silently changed.</p>}
      <FilePick label="Minutes PDF" accept="application/pdf" hint={`PDF preferred, under ${MAX_FILE_KB} KB.`} onFile={setFile} />
      {file && <div className="mb-3"><FileChip f={file} onRemove={() => setFile(null)} /></div>}
      <Field label="Or paste minutes as text"><TextArea rows={5} value={text} onChange={(e) => setText(e.target.value)} placeholder="Call to order, motions, decisions, adjournment…" /></Field>
      {isRevision && <Field label="Reason for revision"><Input value={note} onChange={(e) => setNote(e.target.value)} placeholder="e.g. Corrected treasurer's figure" /></Field>}
      <SectionTitle>Action items</SectionTitle>
      {items.map((it) => (
        <div key={it.id} className="flex items-center gap-2 mb-1.5" style={{ fontSize: 13 }}>
          <Icon name="flag" size={14} color={CRAN} />
          <span className="flex-1">{it.text} — <b>{db.members.find((x) => x.id === it.assignee)?.name}</b>{it.due ? `, ${fmtShort(it.due)}` : ""}</span>
          <button onClick={() => setItems(items.filter((x) => x.id !== it.id))}><Icon name="x" size={14} color="#A896A0" /></button>
        </div>
      ))}
      <Card className="p-3 mt-1">
        <Input value={ni.text} onChange={(e) => setNi({ ...ni, text: e.target.value })} placeholder="Action item…" />
        <div className="grid grid-cols-2 gap-2 mt-2">
          <Select value={ni.assignee} onChange={(e) => setNi({ ...ni, assignee: e.target.value })}><option value="">Assign to…</option>{activeMembers.map((x) => <option key={x.id} value={x.id}>{x.name}</option>)}</Select>
          <Input type="date" value={ni.due} onChange={(e) => setNi({ ...ni, due: e.target.value })} />
        </div>
        <div className="mt-2"><Btn small kind="ghost" onClick={addItem}>+ Add action item</Btn></div>
      </Card>
      <div className="mt-4"><Btn onClick={submit}>{isRevision ? "Save revision & actions" : "Publish minutes to the club"}</Btn></div>
    </Sheet>
  );
}

/* ============ Finance ============ */
function receiptHtml(db, pay, member) {
  const cur = db.duesConfig.currency;
  return docShell(`Receipt ${pay.receiptNo}`, `
    <h1>${db.club.name}</h1><div class="muted">Official payment receipt</div>
    <h2>Receipt ${pay.receiptNo}</h2>
    <table>
      <tr><th>Received from</th><td>${member?.name || ""}</td></tr>
      <tr><th>Amount</th><td class="big">${money(pay.amount, cur)}</td></tr>
      <tr><th>Date</th><td>${fmtDate(pay.date)}</td></tr>
      <tr><th>Method</th><td>${pay.method}</td></tr>
      <tr><th>Reference</th><td>${pay.ref || "—"}</td></tr>
      <tr><th>Applied to</th><td>${pay.appliesTo || "Member dues account"}</td></tr>
      <tr><th>Notes</th><td>${pay.notes || "—"}</td></tr>
      <tr><th>Recorded by</th><td>${pay.byName || ""}</td></tr>
    </table>`, CRAN);
}

function FinanceTab() {
  const { db, me, year, canFinance, isTreasurer, setOverlay } = useApp();
  const [sub, setSub] = useState("overview");
  const cur = db.duesConfig.currency;
  const totals = clubTotals(db, year.id);
  const subs = canFinance
    ? [["overview", "Overview"], ["dues", "Dues"], ["ledger", "Ledger"], ["reports", "Reports"]]
    : [["overview", "Overview"], ["account", "My account"]];
  return (
    <div>
      <div className="flex items-center justify-between mt-5">
        <h2 className="font-black" style={{ fontFamily: DISPLAY, fontSize: 24 }}>Finance</h2>
        <div className="flex gap-1.5">
          {canFinance && <Btn small kind="quiet" onClick={() => setOverlay({ type: "controls" })}><span className="flex items-center gap-1.5"><Icon name="shield" size={14} /> Controls</span></Btn>}
        </div>
      </div>
      <div className="mt-3"><Chips options={subs.map(([id, label]) => ({ id, label }))} value={sub} onChange={setSub} /></div>
      {sub === "overview" && <FinanceOverview totals={totals} cur={cur} />}
      {sub === "account" && <MemberAccountInline memberId={me.id} />}
      {sub === "dues" && canFinance && <DuesManager />}
      {sub === "ledger" && canFinance && <Ledger />}
      {sub === "reports" && canFinance && <FinanceReports />}
    </div>
  );
}

function FinanceOverview({ totals, cur }) {
  const { db, year, canFinance, me, setOverlay } = useApp();
  const tx = db.transactions.filter((t) => t.yearId === year.id && !t.reversed).sort((a, b) => b.date.localeCompare(a.date)).slice(0, 6);
  const arrears = db.members.filter((m) => ACTIVE_LIKE.includes(m.status)).map((m) => ({ m, a: memberAccount(db, m.id, year.id) })).filter((x) => x.a.overdue > 0);
  const missingReceipts = db.transactions.filter((t) => t.yearId === year.id && t.type === "expense" && !t.reversed && !t.receipt).length;
  return (
    <div>
      <Card className="relative overflow-hidden mt-4 p-5" style={{ background: `linear-gradient(135deg, ${INK}, #47253A)`, border: "none" }}>
        <div className="absolute" style={{ right: -24, top: -24 }}><Wheel size={110} color="rgba(255,255,255,.08)" /></div>
        <div className="relative">
          <div className="uppercase font-bold" style={{ fontSize: 10.5, letterSpacing: ".16em", color: "rgba(255,255,255,.6)", fontFamily: DISPLAY }}>Club balance · {year.label}</div>
          <div className="font-black text-white" style={{ fontFamily: DISPLAY, fontSize: 34 }}>{money(totals.balance, cur)}</div>
          <div className="flex gap-4 mt-2">
            <span style={{ fontSize: 13, color: "#7FE0AE" }}>▲ {money(totals.income, cur)} in</span>
            <span style={{ fontSize: 13, color: "#FF9C9C" }}>▼ {money(totals.expense, cur)} out</span>
          </div>
        </div>
      </Card>
      {canFinance && missingReceipts > 0 && (
        <Card className="mt-2.5 p-3 flex items-center gap-2.5" style={{ background: GOLD + "14", borderColor: GOLD }}>
          <Icon name="doc" size={18} color="#8A5A00" />
          <span style={{ fontSize: 13, color: "#7A5A00" }}><b>{missingReceipts}</b> expense{missingReceipts === 1 ? "" : "s"} missing a receipt or invoice.</span>
        </Card>
      )}
      {canFinance && arrears.length > 0 && (
        <>
          <SectionTitle>Members in arrears</SectionTitle>
          <Card className="p-2">
            {arrears.map(({ m, a }, i) => (
              <button key={m.id} onClick={() => setOverlay({ type: "myAccount", id: m.id })} className="w-full flex items-center gap-3 px-2 py-2 text-left" style={{ borderTop: i ? `1px solid ${LINE}` : "none" }}>
                <Avatar m={m} size={30} />
                <span className="flex-1 font-semibold" style={{ fontSize: 13.5 }}>{m.name}</span>
                <span className="font-bold" style={{ fontSize: 13.5, color: BAD }}>{money(a.overdue, cur)}</span>
              </button>
            ))}
          </Card>
        </>
      )}
      <SectionTitle>Recent activity</SectionTitle>
      {tx.length ? (
        <Card className="p-2">
          {tx.map((t, i) => (
            <div key={t.id} className="flex items-center gap-3 px-2 py-2.5" style={{ borderTop: i ? `1px solid ${LINE}` : "none" }}>
              <div className="rounded-xl p-2" style={{ background: (t.type === "income" ? OK : BAD) + "14", color: t.type === "income" ? OK : BAD }}><Icon name="wallet" size={16} /></div>
              <div className="flex-1 min-w-0">
                <div className="font-semibold truncate" style={{ fontSize: 13.5 }}>{t.desc}</div>
                <div style={{ fontSize: 11.5, color: "#8A7580" }}>{t.category} · {fmtShort(t.date)}{t.approval === "pending" ? " · awaiting approval" : ""}</div>
              </div>
              <div className="font-bold" style={{ fontSize: 14, color: t.type === "income" ? OK : BAD }}>{t.type === "income" ? "+" : "−"}{money(t.amount, cur)}</div>
            </div>
          ))}
        </Card>
      ) : <Empty icon="wallet" title="No transactions yet" />}
      {!canFinance && <div className="mt-4"><Btn kind="ghost" onClick={() => setOverlay({ type: "myAccount", id: me.id })}>View my dues account</Btn></div>}
    </div>
  );
}

function MemberAccountInline({ memberId }) { return <MemberAccountBody memberId={memberId} />; }
function MemberAccountScreen({ memberId, onClose }) {
  const { memberById } = useApp();
  return <FullScreen title={`${memberById(memberId)?.name} — account`} onClose={onClose}><MemberAccountBody memberId={memberId} /></FullScreen>;
}
function MemberAccountBody({ memberId }) {
  const { db, year, me, canFinance, isTreasurer, memberById, showToast } = useApp();
  const [payOpen, setPayOpen] = useState(false);
  const [chargeOpen, setChargeOpen] = useState(false);
  const acct = memberAccount(db, memberId, year.id);
  const cur = db.duesConfig.currency;
  const member = memberById(memberId);
  const ledger = [
    ...acct.charges.map((c) => ({ kind: "charge", label: c.label, amount: c.amount, date: c.dueDate, id: c.id })),
    ...acct.payments.map((p) => ({ kind: "payment", label: `Payment · ${p.method}${p.ref ? ` · ${p.ref}` : ""} · ${p.receiptNo}`, amount: -p.amount, date: p.date, id: p.id, pay: p })),
    ...acct.credits.map((c) => ({ kind: "credit", label: `Credit · ${c.reason}`, amount: -c.amount, date: c.date, id: c.id })),
  ].sort((a, b) => b.date.localeCompare(a.date));
  return (
    <div>
      <Card className="p-4 mt-2" style={{ background: acct.overdue > 0 ? BAD + "0d" : acct.balance > 0 ? GOLD + "10" : OK + "0d", borderColor: acct.overdue > 0 ? BAD + "44" : acct.balance > 0 ? GOLD : OK + "44" }}>
        <div className="grid grid-cols-2 gap-3">
          <div><div style={{ fontSize: 11.5, color: "#8A7580" }} className="font-bold uppercase">Current due</div><div className="font-black" style={{ fontFamily: DISPLAY, fontSize: 24 }}>{money(Math.max(0, acct.balance), cur)}</div></div>
          <div><div style={{ fontSize: 11.5, color: "#8A7580" }} className="font-bold uppercase">Overdue</div><div className="font-black" style={{ fontFamily: DISPLAY, fontSize: 24, color: acct.overdue > 0 ? BAD : OK }}>{money(acct.overdue, cur)}</div></div>
        </div>
        <div className="mt-2 flex gap-3 flex-wrap" style={{ fontSize: 12.5, color: "#6B5A64" }}>
          <span>Charged {money(acct.charged, cur)}</span><span>Paid {money(acct.paid, cur)}</span><span>Credits {money(acct.credited, cur)}</span>
        </div>
        {acct.nextDue && acct.balance > 0 && <div className="mt-1.5" style={{ fontSize: 12.5 }}><b>Next payment:</b> {fmtDate(acct.nextDue)}{acct.arrangement ? " (per arrangement)" : ""}</div>}
        {acct.balance < 0 && <div className="mt-1.5 font-bold" style={{ fontSize: 12.5, color: OK }}>Account in credit by {money(-acct.balance, cur)}</div>}
      </Card>
      {acct.arrangement && (
        <Card className="mt-2.5 p-3.5" style={{ borderColor: AZURE + "55" }}>
          <div className="font-bold flex items-center gap-2" style={{ fontSize: 13.5, color: AZURE }}><Icon name="repeat" size={15} /> Payment arrangement</div>
          <div style={{ fontSize: 13, color: "#4A3A44" }} className="mt-1">{acct.arrangement.note} — {money(acct.arrangement.installment, cur)} next on {fmtShort(acct.arrangement.nextDate)}</div>
        </Card>
      )}
      {isTreasurer && (
        <div className="flex gap-2 mt-3">
          <Btn small onClick={() => setPayOpen(true)}><span className="flex items-center gap-1.5"><Icon name="plus" size={14} /> Record payment</span></Btn>
          <Btn small kind="quiet" onClick={() => setChargeOpen(true)}>Charge / credit</Btn>
        </div>
      )}
      <SectionTitle>Account history · {year.label}</SectionTitle>
      {ledger.length ? (
        <Card className="p-2">
          {ledger.map((l, i) => (
            <div key={l.id} className="flex items-center gap-3 px-2 py-2.5" style={{ borderTop: i ? `1px solid ${LINE}` : "none" }}>
              <div className="rounded-xl p-1.5" style={{ background: (l.kind === "charge" ? GOLD : OK) + "18", color: l.kind === "charge" ? "#8A5A00" : OK }}>
                <Icon name={l.kind === "charge" ? "doc" : "check"} size={15} />
              </div>
              <div className="flex-1 min-w-0">
                <div className="font-semibold" style={{ fontSize: 13 }}>{l.label}</div>
                <div style={{ fontSize: 11.5, color: "#8A7580" }}>{fmtShort(l.date)}</div>
              </div>
              <div className="font-bold" style={{ fontSize: 13.5, color: l.amount > 0 ? "#8A5A00" : OK }}>{l.amount > 0 ? "+" : "−"}{money(Math.abs(l.amount), cur)}</div>
              {l.pay && <button onClick={() => printHtml(`Receipt-${l.pay.receiptNo}`, receiptHtml(db, l.pay, member))} aria-label="Save receipt as PDF" style={{ color: AZURE }}><Icon name="download" size={16} /></button>}
            </div>
          ))}
        </Card>
      ) : <Empty icon="wallet" title="No account activity" text="Charges appear when dues obligations are generated." />}
      {payOpen && <PaymentSheet memberId={memberId} onClose={() => setPayOpen(false)} />}
      {chargeOpen && <ChargeCreditSheet memberId={memberId} onClose={() => setChargeOpen(false)} />}
    </div>
  );
}

function PaymentSheet({ memberId, onClose }) {
  const { db, patch, me, year, memberById, notify, audit, guardLock, showToast } = useApp();
  const [f, setF] = useState({ amount: "", date: todayStr(), method: "Cash", ref: "", notes: "", doc: null });
  const member = memberById(memberId);
  const cur = db.duesConfig.currency;
  const save = () => {
    if (!Number(f.amount)) return;
    if (!guardLock(f.date)) return;
    let receiptNo = "";
    let payRec = null;
    patch((d) => {
      d.seq.receipt += 1;
      receiptNo = `R-${String(d.seq.receipt).padStart(4, "0")}`;
      payRec = { id: uid(), yearId: year.id, memberId, amount: Number(f.amount), date: f.date, method: f.method, ref: f.ref.trim(), notes: f.notes.trim(), doc: f.doc, receiptNo, by: me.id, byName: me.name, reversed: false, appliesTo: "Member dues account" };
      d.payments.push(payRec);
      d.transactions.push({ id: uid(), yearId: year.id, type: "income", category: "Dues", desc: `Dues payment — ${member.name} (${receiptNo})`, payee: member.name, amount: Number(f.amount), date: f.date, by: me.id, approval: "approved", projectId: "", receipt: f.doc, reversed: false, reversalOf: null, payId: payRec.id });
      notify(d, { type: "dues", title: "Payment recorded", body: `${member.name}: ${money(Number(f.amount), d.duesConfig.currency)} received (${receiptNo}).` });
      audit(d, "Payment recorded", `${member.name} · ${money(Number(f.amount), d.duesConfig.currency)} · ${receiptNo}`);
    });
    showToast(`Recorded — receipt ${receiptNo}`);
    printHtml(`Receipt-${receiptNo}`, receiptHtml(db, { ...payRec }, member));
    onClose();
  };
  return (
    <Sheet open onClose={onClose} title={`Record payment — ${member.name}`} tall>
      <div className="grid grid-cols-2 gap-3">
        <Field label={`Amount (${cur})`}><Input type="number" min="0" step="0.01" value={f.amount} onChange={(e) => setF({ ...f, amount: e.target.value })} /></Field>
        <Field label="Date"><Input type="date" value={f.date} onChange={(e) => setF({ ...f, date: e.target.value })} /></Field>
      </div>
      <Field label="Method"><Select value={f.method} onChange={(e) => setF({ ...f, method: e.target.value })}>{PAY_METHODS.map((m) => <option key={m}>{m}</option>)}</Select></Field>
      <Field label="Reference number"><Input value={f.ref} onChange={(e) => setF({ ...f, ref: e.target.value })} placeholder="Transfer / cheque no." /></Field>
      <Field label="Notes"><Input value={f.notes} onChange={(e) => setF({ ...f, notes: e.target.value })} /></Field>
      <FilePick label="Supporting document" accept="image/*,application/pdf" hint="Deposit slip or proof of payment (optional)." onFile={(file) => setF({ ...f, doc: file })} />
      {f.doc && <div className="mb-3"><FileChip f={f.doc} onRemove={() => setF({ ...f, doc: null })} /></div>}
      <Btn onClick={save} disabled={!Number(f.amount)}>Save & download receipt</Btn>
      <p className="mt-2" style={{ fontSize: 12, color: "#9A8B93" }}>Receipts are numbered sequentially and the payment posts to the club ledger automatically.</p>
    </Sheet>
  );
}

function ChargeCreditSheet({ memberId, onClose }) {
  const { db, patch, me, year, memberById, notify, audit, guardLock, showToast } = useApp();
  const [mode, setMode] = useState("charge");
  const [f, setF] = useState({ kind: "penalty", label: "", amount: "", date: todayStr() });
  const member = memberById(memberId);
  const save = () => {
    if (!Number(f.amount)) return;
    if (!guardLock(f.date)) return;
    patch((d) => {
      if (mode === "charge") {
        d.charges.push({ id: uid(), yearId: year.id, memberId, kind: f.kind, label: f.label.trim() || CHARGE_KINDS[f.kind], period: f.date.slice(0, 7), amount: Number(f.amount), dueDate: f.date, at: Date.now(), by: me.id, reversed: false });
        notify(d, { type: "dues", title: "Charge added", body: `${member.name}: ${f.label.trim() || CHARGE_KINDS[f.kind]} — ${money(Number(f.amount), d.duesConfig.currency)}.` });
        audit(d, "Charge added", `${member.name} · ${CHARGE_KINDS[f.kind]} · ${money(Number(f.amount), d.duesConfig.currency)}`);
      } else {
        d.credits.push({ id: uid(), yearId: year.id, memberId, amount: Number(f.amount), date: f.date, reason: f.label.trim() || "Credit adjustment", by: me.id, reversed: false });
        audit(d, "Credit issued", `${member.name} · ${money(Number(f.amount), d.duesConfig.currency)} · ${f.label.trim()}`);
      }
    });
    showToast(mode === "charge" ? "Charge added" : "Credit issued");
    onClose();
  };
  return (
    <Sheet open onClose={onClose} title={`${member.name} — adjust account`}>
      <div className="flex gap-2 mb-4">
        {["charge", "credit"].map((m) => (
          <button key={m} onClick={() => setMode(m)} className="flex-1 rounded-xl font-bold py-2.5 capitalize" style={{ fontFamily: DISPLAY, fontSize: 14, background: mode === m ? (m === "charge" ? "#8A5A00" : OK) : "#fff", color: mode === m ? "#fff" : "#6B5A64", border: `1.5px solid ${mode === m ? "transparent" : LINE}` }}>{m}</button>
        ))}
      </div>
      {mode === "charge" && (
        <Field label="Type"><Select value={f.kind} onChange={(e) => setF({ ...f, kind: e.target.value })}>
          {["penalty", "happy", "late", "other"].map((k) => <option key={k} value={k}>{CHARGE_KINDS[k]}</option>)}
        </Select></Field>
      )}
      <Field label={mode === "charge" ? "Description" : "Reason"}><Input value={f.label} onChange={(e) => setF({ ...f, label: e.target.value })} placeholder={mode === "charge" ? "e.g. Happy Dollars — good news!" : "e.g. Waived June dues (student)"} /></Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Amount"><Input type="number" min="0" step="0.01" value={f.amount} onChange={(e) => setF({ ...f, amount: e.target.value })} /></Field>
        <Field label="Date"><Input type="date" value={f.date} onChange={(e) => setF({ ...f, date: e.target.value })} /></Field>
      </div>
      <Btn onClick={save} disabled={!Number(f.amount)}>{mode === "charge" ? "Add charge" : "Issue credit"}</Btn>
    </Sheet>
  );
}

function DuesManager() {
  const { db, patch, me, year, memberById, notify, audit, guardLock, showToast, setOverlay } = useApp();
  const [cfgOpen, setCfgOpen] = useState(false);
  const [arrOpen, setArrOpen] = useState(null);
  const [cfg, setCfg] = useState({ ...db.duesConfig });
  const cur = db.duesConfig.currency;
  const members = db.members.filter((m) => ACTIVE_LIKE.includes(m.status));
  const saveCfg = () => {
    patch((d) => {
      d.duesConfig = { ...d.duesConfig, monthly: +cfg.monthly || 0, district: +cfg.district || 0, ri: +cfg.ri || 0, currency: cfg.currency || "$", dueDay: Math.min(28, Math.max(1, +cfg.dueDay || 5)), graceDays: +cfg.graceDays || 0, lateFee: +cfg.lateFee || 0, lateFeeOn: !!cfg.lateFeeOn };
      ensureObligations(d);
      audit(d, "Dues configuration changed", `Monthly ${money(+cfg.monthly, cfg.currency)}, due day ${cfg.dueDay}, grace ${cfg.graceDays}d`);
    });
    setCfgOpen(false); showToast("Dues configuration saved");
  };
  const regen = () => { patch((d) => { const c = ensureObligations(d); if (c) audit(d, "Obligations generated", "Monthly/district/RI charges created"); }); showToast("Obligations up to date"); };
  const applyLateFees = () => {
    if (!db.duesConfig.lateFee) { showToast("Set a late charge amount first"); return; }
    let n = 0;
    patch((d) => {
      const grace = d.duesConfig.graceDays || 0; const t = todayStr();
      members.forEach((m) => {
        const acct = memberAccount(d, m.id, year.id);
        if (acct.overdue > 0) {
          const period = t.slice(0, 7);
          if (!d.charges.some((c) => c.memberId === m.id && c.kind === "late" && c.period === period && !c.reversed)) {
            d.charges.push({ id: uid(), yearId: year.id, memberId: m.id, kind: "late", label: `Late charge ${period}`, period, amount: d.duesConfig.lateFee, dueDate: t, at: Date.now(), by: me.id, reversed: false });
            notify(d, { type: "dues", title: "Overdue dues", body: `${m.name}: account is overdue — a late charge of ${money(d.duesConfig.lateFee, cur)} was applied.` });
            n++;
          }
        }
      });
      if (n) audit(d, "Late charges applied", `${n} member(s)`);
    });
    showToast(n ? `Late charge applied to ${n} member(s)` : "No overdue accounts without a late charge this month");
  };
  const toggleExempt = (mid) => {
    patch((d) => {
      const ex = d.exemptions.find((e) => e.memberId === mid && e.active);
      if (ex) { ex.active = false; audit(d, "Exemption ended", memberById(mid)?.name); }
      else { d.exemptions.push({ id: uid(), memberId: mid, active: true, note: "Exempt from dues", at: Date.now(), by: me.id }); audit(d, "Exemption granted", memberById(mid)?.name); }
    });
  };
  const saveArrangement = (mid, note, installment, nextDate) => {
    patch((d) => {
      d.arrangements.forEach((a) => { if (a.memberId === mid) a.active = false; });
      if (note || installment) {
        d.arrangements.push({ id: uid(), memberId: mid, note, installment: +installment || 0, nextDate, active: true, at: Date.now(), by: me.id });
        audit(d, "Payment arrangement set", `${memberById(mid)?.name} · ${money(+installment || 0, cur)} next ${nextDate}`);
      } else audit(d, "Payment arrangement cleared", memberById(mid)?.name);
    });
    setArrOpen(null); showToast("Arrangement saved");
  };
  return (
    <div>
      <div className="flex gap-1.5 mt-4 flex-wrap">
        <Btn small kind="quiet" onClick={() => { setCfg({ ...db.duesConfig }); setCfgOpen(true); }}><span className="flex items-center gap-1.5"><Icon name="edit" size={14} /> Configure</span></Btn>
        <Btn small kind="quiet" onClick={regen}>Generate obligations</Btn>
        <Btn small kind="quiet" onClick={applyLateFees}>Apply late charges</Btn>
      </div>
      <p className="mt-2 mb-2" style={{ fontSize: 12, color: "#9A8B93" }}>Obligations (monthly, district, RI) are auto-created for each active, non-exempt member. Due day {db.duesConfig.dueDay}, grace {db.duesConfig.graceDays} days.</p>
      <div className="flex flex-col gap-2">
        {members.map((m) => {
          const a = memberAccount(db, m.id, year.id);
          const exempt = db.exemptions.some((e) => e.memberId === m.id && e.active);
          return (
            <Card key={m.id} className="p-3.5">
              <div className="flex items-center gap-2.5">
                <button onClick={() => setOverlay({ type: "myAccount", id: m.id })}><Avatar m={m} size={34} /></button>
                <div className="flex-1 min-w-0">
                  <button onClick={() => setOverlay({ type: "myAccount", id: m.id })} className="font-bold truncate block text-left" style={{ fontSize: 14 }}>{m.name}</button>
                  <div style={{ fontSize: 11.5, color: a.overdue > 0 ? BAD : a.balance > 0 ? "#8A5A00" : OK }}>
                    {exempt ? "Exempt from dues" : a.overdue > 0 ? `${money(a.overdue, cur)} overdue` : a.balance > 0 ? `${money(a.balance, cur)} due ${a.nextDue ? fmtShort(a.nextDue) : ""}` : "Up to date"}
                  </div>
                </div>
                <div className="flex gap-1">
                  <button onClick={() => setArrOpen(m)} title="Arrangement" className="rounded-full p-1.5" style={{ background: a.arrangement ? AZURE + "18" : "#F0EAEE", color: a.arrangement ? AZURE : "#8A7580" }}><Icon name="repeat" size={15} /></button>
                  <button onClick={() => toggleExempt(m.id)} title="Exemption" className="rounded-full p-1.5" style={{ background: exempt ? "#7A4BA618" : "#F0EAEE", color: exempt ? "#7A4BA6" : "#8A7580" }}><Icon name="star" size={15} /></button>
                </div>
              </div>
            </Card>
          );
        })}
      </div>
      <Sheet open={cfgOpen} onClose={() => setCfgOpen(false)} title="Dues configuration" tall>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Monthly club dues"><Input type="number" min="0" value={cfg.monthly} onChange={(e) => setCfg({ ...cfg, monthly: e.target.value })} /></Field>
          <Field label="District dues (annual)"><Input type="number" min="0" value={cfg.district} onChange={(e) => setCfg({ ...cfg, district: e.target.value })} /></Field>
          <Field label="RI charges (annual)"><Input type="number" min="0" value={cfg.ri} onChange={(e) => setCfg({ ...cfg, ri: e.target.value })} /></Field>
          <Field label="Currency symbol"><Input value={cfg.currency} onChange={(e) => setCfg({ ...cfg, currency: e.target.value })} /></Field>
          <Field label="Due day of month"><Input type="number" min="1" max="28" value={cfg.dueDay} onChange={(e) => setCfg({ ...cfg, dueDay: e.target.value })} /></Field>
          <Field label="Grace period (days)"><Input type="number" min="0" value={cfg.graceDays} onChange={(e) => setCfg({ ...cfg, graceDays: e.target.value })} /></Field>
          <Field label="Late charge"><Input type="number" min="0" value={cfg.lateFee} onChange={(e) => setCfg({ ...cfg, lateFee: e.target.value })} /></Field>
        </div>
        <Btn onClick={saveCfg}>Save configuration</Btn>
        <p className="mt-2" style={{ fontSize: 12, color: "#9A8B93" }}>Penalties and Happy Dollars are added per member from their account page.</p>
      </Sheet>
      {arrOpen && <ArrangementSheet member={arrOpen} existing={db.arrangements.find((a) => a.memberId === arrOpen.id && a.active)} onSave={saveArrangement} onClose={() => setArrOpen(null)} />}
    </div>
  );
}

function ArrangementSheet({ member, existing, onSave, onClose }) {
  const [note, setNote] = useState(existing?.note || "");
  const [inst, setInst] = useState(existing?.installment || "");
  const [next, setNext] = useState(existing?.nextDate || "");
  return (
    <Sheet open onClose={onClose} title={`Arrangement — ${member.name}`}>
      <Field label="Terms" hint="e.g. Pays half now, half at month end."><Input value={note} onChange={(e) => setNote(e.target.value)} /></Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Installment amount"><Input type="number" min="0" value={inst} onChange={(e) => setInst(e.target.value)} /></Field>
        <Field label="Next payment date"><Input type="date" value={next} onChange={(e) => setNext(e.target.value)} /></Field>
      </div>
      <Btn onClick={() => onSave(member.id, note.trim(), inst, next)}>Save arrangement</Btn>
      {existing && <div className="mt-2"><Btn kind="danger" onClick={() => onSave(member.id, "", "", "")}>End arrangement</Btn></div>}
    </Sheet>
  );
}

function Ledger() {
  const { db, patch, me, year, isPres, isTreasurer, memberById, audit, notify, guardLock, showToast } = useApp();
  const [adding, setAdding] = useState(false);
  const [f, setF] = useState({ type: "expense", category: TX_CATS[6], payee: "", desc: "", amount: "", date: todayStr(), projectId: "", receipt: null });
  const cur = db.duesConfig.currency;
  const list = db.transactions.filter((t) => t.yearId === year.id).sort((a, b) => b.date.localeCompare(a.date));
  const add = () => {
    if (!f.desc.trim() || !Number(f.amount)) return;
    if (!guardLock(f.date)) return;
    patch((d) => {
      d.transactions.push({ id: uid(), yearId: year.id, type: f.type, category: f.category, payee: f.payee.trim(), desc: f.desc.trim(), amount: Number(f.amount), date: f.date, by: me.id, approval: f.type === "expense" ? "pending" : "approved", projectId: f.projectId, receipt: f.receipt, reversed: false, reversalOf: null });
      audit(d, `${f.type === "income" ? "Income" : "Expense"} recorded`, `${f.desc.trim()} · ${money(Number(f.amount), cur)}`);
    });
    showToast(f.type === "expense" ? "Expense recorded — awaiting approval" : "Income recorded");
    setAdding(false); setF({ type: "expense", category: TX_CATS[6], payee: "", desc: "", amount: "", date: todayStr(), projectId: "", receipt: null });
  };
  const approve = (id, ok) => patch((d) => {
    const t = d.transactions.find((x) => x.id === id);
    t.approval = ok ? "approved" : "rejected";
    audit(d, `Expense ${ok ? "approved" : "rejected"}`, `${t.desc} · ${money(t.amount, cur)}`);
  });
  const reverse = (id) => {
    const t0 = db.transactions.find((x) => x.id === id);
    if (!guardLock(t0.date)) return;
    patch((d) => {
      const t = d.transactions.find((x) => x.id === id);
      t.reversed = true;
      d.transactions.push({ ...t, id: uid(), desc: `REVERSAL — ${t.desc}`, reversed: false, reversalOf: t.id, type: t.type === "income" ? "expense" : "income", approval: "approved", date: todayStr(), by: me.id, receipt: null, payId: null });
      if (t.payId) { const p = d.payments.find((x) => x.id === t.payId); if (p) p.reversed = true; }
      audit(d, "Transaction reversed", `${t.desc} · ${money(t.amount, cur)} — reversing entry posted`);
    });
    showToast("Reversal posted — original preserved");
  };
  return (
    <div>
      <div className="flex justify-between items-center mt-4">
        <p style={{ fontSize: 12, color: "#9A8B93" }}>Entries are never deleted — reversals keep the audit trail intact.</p>
        <Btn small onClick={() => setAdding(true)}><span className="flex items-center gap-1.5"><Icon name="plus" size={14} /> Entry</span></Btn>
      </div>
      <div className="flex flex-col gap-2 mt-3">
        {list.length ? list.map((t) => (
          <Card key={t.id} className="p-3.5" style={t.reversed ? { opacity: 0.55 } : {}}>
            <div className="flex items-center gap-3">
              <div className="rounded-xl p-2" style={{ background: (t.type === "income" ? OK : BAD) + "14", color: t.type === "income" ? OK : BAD }}><Icon name="wallet" size={16} /></div>
              <div className="flex-1 min-w-0">
                <div className="font-semibold" style={{ fontSize: 13.5, textDecoration: t.reversed ? "line-through" : "none" }}>{t.desc}</div>
                <div style={{ fontSize: 11.5, color: "#8A7580" }}>
                  {t.category}{t.payee ? ` · ${t.payee}` : ""} · {fmtShort(t.date)} · by {memberById(t.by)?.name?.split(" ")[0] || "system"}
                  {t.projectId ? ` · ${db.projects.find((p) => p.id === t.projectId)?.title || "project"}` : ""}
                </div>
                <div className="mt-1 flex gap-1.5 flex-wrap">
                  {t.approval === "pending" && <Badge color={GOLD}>Awaiting approval</Badge>}
                  {t.approval === "rejected" && <Badge color={BAD}>Rejected</Badge>}
                  {t.reversed && <Badge color={BAD}>Reversed</Badge>}
                  {t.reversalOf && <Badge color="#7A4BA6">Reversal entry</Badge>}
                  {t.type === "expense" && !t.receipt && !t.reversed && !t.reversalOf && <Badge color={GOLD}>No receipt</Badge>}
                  {t.receipt && <button onClick={() => downloadDataUrl(t.receipt.name, t.receipt.dataUrl)}><Badge color={AZURE}>Receipt ↓</Badge></button>}
                </div>
              </div>
              <div className="text-right">
                <div className="font-bold" style={{ fontSize: 14, color: t.type === "income" ? OK : BAD }}>{t.type === "income" ? "+" : "−"}{money(t.amount, cur)}</div>
                {!t.reversed && !t.reversalOf && isTreasurer && <button onClick={() => reverse(t.id)} className="font-bold" style={{ fontSize: 11, color: "#A896A0" }}>reverse</button>}
              </div>
            </div>
            {t.approval === "pending" && isPres && (
              <div className="flex gap-2 mt-2.5">
                <Btn small kind="ok" onClick={() => approve(t.id, true)}>Approve</Btn>
                <Btn small kind="danger" onClick={() => approve(t.id, false)}>Reject</Btn>
              </div>
            )}
          </Card>
        )) : <Empty icon="wallet" title="Ledger is empty" />}
      </div>
      <Sheet open={adding} onClose={() => setAdding(false)} title="Record a transaction" tall>
        <div className="flex gap-2 mb-4">
          {["income", "expense"].map((tp) => (
            <button key={tp} onClick={() => setF({ ...f, type: tp })} className="flex-1 rounded-xl font-bold py-2.5 capitalize" style={{ fontFamily: DISPLAY, fontSize: 14, background: f.type === tp ? (tp === "income" ? OK : BAD) : "#fff", color: f.type === tp ? "#fff" : "#6B5A64", border: `1.5px solid ${f.type === tp ? "transparent" : LINE}` }}>{tp}</button>
          ))}
        </div>
        <Field label="Category"><Select value={f.category} onChange={(e) => setF({ ...f, category: e.target.value })}>{TX_CATS.map((c) => <option key={c}>{c}</option>)}</Select></Field>
        {f.type === "expense" && <Field label="Payee"><Input value={f.payee} onChange={(e) => setF({ ...f, payee: e.target.value })} placeholder="Who was paid" /></Field>}
        <Field label="Description"><Input value={f.desc} onChange={(e) => setF({ ...f, desc: e.target.value })} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label={`Amount (${cur})`}><Input type="number" min="0" step="0.01" value={f.amount} onChange={(e) => setF({ ...f, amount: e.target.value })} /></Field>
          <Field label="Date"><Input type="date" value={f.date} onChange={(e) => setF({ ...f, date: e.target.value })} /></Field>
        </div>
        <Field label="Project (optional)"><Select value={f.projectId} onChange={(e) => setF({ ...f, projectId: e.target.value })}><option value="">Not project-related</option>{db.projects.filter((p) => ["Approved", "Pending Close", "Completed"].includes(p.status)).map((p) => <option key={p.id} value={p.id}>{p.title}</option>)}</Select></Field>
        <FilePick label="Receipt / invoice" accept="image/*,application/pdf" onFile={(file) => setF({ ...f, receipt: file })} />
        {f.receipt && <div className="mb-3"><FileChip f={f.receipt} onRemove={() => setF({ ...f, receipt: null })} /></div>}
        <Btn onClick={add} disabled={!f.desc.trim() || !Number(f.amount)}>Save entry</Btn>
        {f.type === "expense" && <p className="mt-2" style={{ fontSize: 12, color: "#9A8B93" }}>Expenses await the president's approval before counting toward the balance.</p>}
      </Sheet>
    </div>
  );
}

function FinanceReports() {
  const { db, year, memberById } = useApp();
  const cur = db.duesConfig.currency;
  const [month, setMonth] = useState(todayStr().slice(0, 7));
  const months = ryMonths(year.startYear).filter((p) => `${p}-01` <= todayStr());

  const monthTx = db.transactions.filter((t) => t.yearId === year.id && !t.reversed && t.date.slice(0, 7) === month && (t.type === "income" || t.approval === "approved"));
  const mIncome = monthTx.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const mExpense = monthTx.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);

  const arrearsRows = () => {
    const rows = [["Member", "Charged", "Paid", "Credits", "Balance", "Overdue", "Next due", "Arrangement"]];
    db.members.filter((m) => ACTIVE_LIKE.includes(m.status)).forEach((m) => {
      const a = memberAccount(db, m.id, year.id);
      rows.push([m.name, a.charged.toFixed(2), a.paid.toFixed(2), a.credited.toFixed(2), a.balance.toFixed(2), a.overdue.toFixed(2), a.nextDue || "", a.arrangement ? a.arrangement.note : ""]);
    });
    return rows;
  };
  const stmtRows = () => {
    const rows = [["Date", "Type", "Category", "Description", "Payee", "Amount", "Entered by", "Status"]];
    db.transactions.filter((t) => t.yearId === year.id && !t.reversed).sort((a, b) => a.date.localeCompare(b.date))
      .forEach((t) => rows.push([t.date, t.type, t.category, t.desc, t.payee || "", (t.type === "income" ? 1 : -1) * t.amount, memberById(t.by)?.name || "system", t.approval]));
    const totals = clubTotals(db, year.id);
    rows.push([], ["", "", "", "Total income", "", totals.income.toFixed(2)], ["", "", "", "Total expense", "", totals.expense.toFixed(2)], ["", "", "", "Net balance", "", totals.balance.toFixed(2)]);
    return rows;
  };
  const projectRows = () => {
    const rows = [["Project", "Status", "Budget", "Expected income", "Actual income", "Actual expense", "Net"]];
    db.projects.filter((p) => !["Draft", "Denied"].includes(p.status)).forEach((p) => {
      const tx = db.transactions.filter((t) => t.projectId === p.id && !t.reversed && (t.type === "income" || t.approval === "approved"));
      const inc = tx.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
      const exp = tx.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
      rows.push([p.title, p.status, p.budget || 0, p.expectedIncome || 0, inc.toFixed(2), exp.toFixed(2), (inc - exp).toFixed(2)]);
    });
    return rows;
  };
  const monthlyHtml = () => {
    const rowsHtml = monthTx.sort((a, b) => a.date.localeCompare(b.date)).map((t) => `<tr><td>${t.date}</td><td>${t.type}</td><td>${t.category}</td><td>${t.desc}</td><td style="text-align:right">${t.type === "income" ? "" : "−"}${money(t.amount, cur)}</td></tr>`).join("");
    const arrears = db.members.filter((m) => ACTIVE_LIKE.includes(m.status)).map((m) => ({ m, a: memberAccount(db, m.id, year.id) })).filter((x) => x.a.overdue > 0);
    return docShell(`Treasurer's report ${month}`, `
      <h1>${db.club.name}</h1><div class="muted">Treasurer's monthly report — ${month} · Rotary year ${year.label}</div>
      <h2>Summary</h2>
      <table><tr><th>Income</th><td>${money(mIncome, cur)}</td></tr><tr><th>Expenses</th><td>${money(mExpense, cur)}</td></tr><tr><th>Net for month</th><td class="big">${money(mIncome - mExpense, cur)}</td></tr><tr><th>Club balance (year to date)</th><td>${money(clubTotals(db, year.id).balance, cur)}</td></tr></table>
      <h2>Transactions</h2><table><tr><th>Date</th><th>Type</th><th>Category</th><th>Description</th><th>Amount</th></tr>${rowsHtml || "<tr><td colspan=5>None</td></tr>"}</table>
      <h2>Members in arrears</h2><table><tr><th>Member</th><th>Overdue</th></tr>${arrears.map((x) => `<tr><td>${x.m.name}</td><td>${money(x.a.overdue, cur)}</td></tr>`).join("") || "<tr><td colspan=2>None 🎉</td></tr>"}</table>`, CRAN);
  };

  const ReportRow = ({ title, desc, rows, htmlFn, name }) => (
    <Card className="p-4">
      <div className="font-extrabold" style={{ fontFamily: DISPLAY, fontSize: 15 }}>{title}</div>
      <div style={{ fontSize: 12.5, color: "#8A7580" }}>{desc}</div>
      <div className="flex gap-1.5 mt-2.5 flex-wrap">
        {htmlFn && <Btn small kind="quiet" onClick={() => printHtml(name, htmlFn())}>Save as PDF</Btn>}
        {rows && <Btn small kind="quiet" onClick={() => exportCsv(name, rows())}>CSV</Btn>}
        {rows && <Btn small kind="quiet" onClick={() => exportXlsx(name, rows())}>Excel</Btn>}
      </div>
    </Card>
  );
  return (
    <div className="flex flex-col gap-2.5 mt-4">
      <Card className="p-4">
        <div className="font-extrabold mb-2" style={{ fontFamily: DISPLAY, fontSize: 15 }}>Treasurer's monthly report</div>
        <Select value={month} onChange={(e) => setMonth(e.target.value)}>{months.map((p) => <option key={p} value={p}>{p}</option>)}</Select>
        <div style={{ fontSize: 13, color: "#6B5A64" }} className="mt-2">Income {money(mIncome, cur)} · Expenses {money(mExpense, cur)} · Net {money(mIncome - mExpense, cur)}</div>
        <div className="flex gap-1.5 mt-2.5"><Btn small kind="quiet" onClick={() => printHtml(`Treasurer-report-${month}`, monthlyHtml())}>Save as PDF</Btn></div>
      </Card>
      <ReportRow title="Member arrears report" desc="Every member's charges, payments, balance, and overdue amount." rows={arrearsRows} name={`Arrears-${year.id}`} />
      <ReportRow title="Income & expense statement" desc="Full-year statement with running totals." rows={stmtRows} name={`Statement-${year.id}`} />
      <ReportRow title="Project financial report" desc="Budget vs actual income and spend per project." rows={projectRows} name={`Projects-financial-${year.id}`} />
      <p style={{ fontSize: 12, color: "#9A8B93" }}>"Save as PDF" opens your device's print dialog — choose "Save as PDF" as the destination. CSV/Excel open in any spreadsheet.</p>
    </div>
  );
}

function FinancialControls({ onClose }) {
  const { db, patch, me, year, isTreasurer, audit, showToast, memberById } = useApp();
  const [tab2, setTab2] = useState("audit");
  const months = ryMonths(year.startYear).filter((p) => `${p}-01` <= todayStr());
  const [note, setNote] = useState("");
  const toggleLock = (p) => {
    patch((d) => {
      if (d.locks[p]) { delete d.locks[p]; audit(d, "Period unlocked", p); }
      else { d.locks[p] = { by: me.name, at: Date.now() }; audit(d, "Period locked", p); }
    });
  };
  const reconcile = (p) => {
    patch((d) => { d.reconciliations[p] = { by: me.name, at: Date.now(), note: note.trim() }; audit(d, "Month reconciled", `${p}${note ? ` — ${note}` : ""}`); });
    setNote(""); showToast(`${p} marked reconciled`);
  };
  return (
    <FullScreen title="Financial controls" onClose={onClose} accent={INK}>
      <Chips options={[{ id: "audit", label: "Audit log" }, { id: "locks", label: "Period locks" }, { id: "recon", label: "Reconciliation" }]} value={tab2} onChange={setTab2} />
      {tab2 === "audit" && (
        <>
          <p className="mt-3" style={{ fontSize: 12.5, color: "#8A7580" }}>Append-only record of every significant action. Entries cannot be edited or deleted.</p>
          <Card className="p-2 mt-2">
            {db.audit.slice(0, 60).map((a, i) => (
              <div key={a.id} className="px-2 py-2" style={{ borderTop: i ? `1px solid ${LINE}` : "none" }}>
                <div className="font-semibold" style={{ fontSize: 13 }}>{a.action}</div>
                <div style={{ fontSize: 11.5, color: "#8A7580" }}>{a.detail ? a.detail + " · " : ""}{a.by} · {timeAgo(a.ts)}</div>
              </div>
            ))}
          </Card>
        </>
      )}
      {tab2 === "locks" && (
        <>
          <p className="mt-3" style={{ fontSize: 12.5, color: "#8A7580" }}>Locked months reject new entries, edits, and reversals — close a period once it's reconciled.</p>
          <Card className="p-2 mt-2">
            {months.map((p, i) => (
              <div key={p} className="flex items-center gap-3 px-2 py-2.5" style={{ borderTop: i ? `1px solid ${LINE}` : "none" }}>
                <Icon name="lock" size={16} color={db.locks[p] ? BAD : "#C9B8C1"} />
                <div className="flex-1"><div className="font-bold" style={{ fontSize: 14 }}>{p}</div>
                  {db.locks[p] && <div style={{ fontSize: 11.5, color: "#8A7580" }}>Locked by {db.locks[p].by} · {timeAgo(db.locks[p].at)}</div>}
                </div>
                {isTreasurer && <Btn small kind={db.locks[p] ? "quiet" : "dark"} onClick={() => toggleLock(p)}>{db.locks[p] ? "Unlock" : "Lock"}</Btn>}
              </div>
            ))}
          </Card>
        </>
      )}
      {tab2 === "recon" && (
        <>
          <p className="mt-3" style={{ fontSize: 12.5, color: "#8A7580" }}>Confirm each month's ledger matches the bank/cash count, then lock the period.</p>
          {isTreasurer && <div className="mt-2"><Input value={note} onChange={(e) => setNote(e.target.value)} placeholder="Reconciliation note (e.g. matches bank stmt #6)" /></div>}
          <Card className="p-2 mt-2">
            {months.map((p, i) => {
              const r = db.reconciliations[p];
              const tx = db.transactions.filter((t) => t.yearId === year.id && !t.reversed && t.date.slice(0, 7) === p && (t.type === "income" || t.approval === "approved"));
              const net = tx.reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);
              return (
                <div key={p} className="flex items-center gap-3 px-2 py-2.5" style={{ borderTop: i ? `1px solid ${LINE}` : "none" }}>
                  <Icon name="check" size={16} color={r ? OK : "#C9B8C1"} />
                  <div className="flex-1">
                    <div className="font-bold" style={{ fontSize: 14 }}>{p} <span style={{ fontWeight: 400, color: "#8A7580", fontSize: 12 }}>net {money(net, db.duesConfig.currency)}</span></div>
                    {r && <div style={{ fontSize: 11.5, color: "#8A7580" }}>Reconciled by {r.by} · {timeAgo(r.at)}{r.note ? ` · ${r.note}` : ""}</div>}
                  </div>
                  {isTreasurer && !r && <Btn small kind="ok" onClick={() => reconcile(p)}>Reconcile</Btn>}
                </div>
              );
            })}
          </Card>
        </>
      )}
    </FullScreen>
  );
}

/* ============ Projects ============ */
const PROJ_STATUS_COLOR = { Draft: "#8A7580", Submitted: GOLD, "Under Review": "#B07400", Approved: OK, Denied: BAD, "Returned": "#7A4BA6", "Pending Close": AZURE, Completed: "#0E8F8F" };

function ProjectsTab() {
  const { db, me, setOverlay } = useApp();
  const [propose, setPropose] = useState(null); // null | {} | draft project
  const [filter, setFilter] = useState("all");
  const canPropose = !["Applied", "Invited", "Resigned", "Transferred"].includes(me.status) || me.status === "Prospect";
  const visible = db.projects.filter((p) => p.status !== "Draft" || p.submittedBy === me.id);
  const filtered = visible.filter((p) => filter === "all" ? true : filter === "active" ? p.status === "Approved" : filter === "review" ? ["Submitted", "Under Review", "Returned", "Pending Close"].includes(p.status) : filter === "drafts" ? p.status === "Draft" : ["Completed", "Denied"].includes(p.status));
  const sorted = [...filtered].sort((a, b) => b.at - a.at);
  return (
    <div>
      <div className="flex items-center justify-between mt-5">
        <h2 className="font-black" style={{ fontFamily: DISPLAY, fontSize: 24 }}>Projects</h2>
        {canPropose && <Btn small onClick={() => setPropose({})}><span className="flex items-center gap-1.5"><Icon name="plus" size={15} /> Propose</span></Btn>}
      </div>
      <div className="mt-3"><Chips options={[{ id: "all", label: "All" }, { id: "active", label: "Active" }, { id: "review", label: "In review" }, { id: "drafts", label: "My drafts" }, { id: "done", label: "Closed" }]} value={filter} onChange={setFilter} /></div>
      <div className="flex flex-col gap-2 mt-3">
        {sorted.length ? sorted.map((p) => {
          const done = p.tasks.filter((t) => t.status === "Completed").length;
          return (
            <Card key={p.id} onClick={() => p.status === "Draft" ? setPropose(p) : setOverlay({ type: "project", id: p.id })} className="p-4">
              <div className="flex items-start justify-between gap-2">
                <div className="font-extrabold" style={{ fontFamily: DISPLAY, fontSize: 16 }}>{p.title}</div>
                <Badge color={PROJ_STATUS_COLOR[p.status]}>{p.status}</Badge>
              </div>
              <div className="mt-0.5" style={{ fontSize: 12.5, color: "#8A7580" }}>{p.area} · led by {db.members.find((m) => m.id === p.lead)?.name || "—"}</div>
              <p className="mt-1.5" style={{ fontSize: 13, color: "#4A3A44", display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical", overflow: "hidden" }}>{p.description}</p>
              {p.status === "Approved" && p.tasks.length > 0 && (
                <div className="mt-2.5"><div className="flex justify-between mb-1" style={{ fontSize: 11.5, color: "#8A7580" }}><span>{done}/{p.tasks.length} tasks</span><span>{Math.round((done / p.tasks.length) * 100)}%</span></div><Progress pct={(done / p.tasks.length) * 100} color={OK} /></div>
              )}
            </Card>
          );
        }) : <Empty icon="target" title="Nothing here" text="Propose a service project — every member's ideas are welcome." />}
      </div>
      {propose && <ProposeSheet draft={propose.id ? propose : null} onClose={() => setPropose(null)} />}
    </div>
  );
}

function ProposeSheet({ draft, onClose }) {
  const { db, patch, me, year, notify, audit, showToast } = useApp();
  const [step, setStep] = useState(0);
  const [f, setF] = useState(draft ? clone(draft) : {
    title: "", area: AREAS[0], description: "", problem: "", objectives: "", goals: "", impact: "", beneficiaries: "",
    location: "", dates: "", volunteersNeeded: "", partners: "", budget: "", expectedIncome: "", materials: "", risks: "", success: "", sustainability: "", files: [],
  });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const steps = ["Basics", "Plan", "Resources"];
  const save = (submit) => {
    if (!f.title.trim() || !f.description.trim()) return;
    patch((d) => {
      const base = { ...f, title: f.title.trim(), volunteersNeeded: +f.volunteersNeeded || 0, budget: +f.budget || 0, expectedIncome: +f.expectedIncome || 0 };
      if (draft) {
        const p = d.projects.find((x) => x.id === draft.id);
        Object.assign(p, base);
        if (submit) { p.status = "Submitted"; p.at = Date.now(); p.decisions.push({ id: uid(), by: me.id, action: "Submitted", reason: "", at: Date.now() }); }
      } else {
        d.projects.push({
          id: uid(), yearId: year.id, ...base, lead: me.id, submittedBy: me.id, at: Date.now(),
          status: submit ? "Submitted" : "Draft",
          decisions: submit ? [{ id: uid(), by: me.id, action: "Submitted", reason: "", at: Date.now() }] : [],
          votes: {}, questions: [], ebodNotes: [], recommendations: "",
          tasks: [], milestones: [], volunteerRoles: [], updates: [], discussion: [], serviceHours: [], risksLog: [], finalReport: null, completedAt: null,
        });
      }
      if (submit) notify(d, { type: "projects", title: "New proposal", body: `"${f.title.trim()}" submitted by ${me.name} — EBOD review requested.` });
      audit(d, submit ? "Proposal submitted" : "Proposal draft saved", f.title.trim());
    });
    showToast(submit ? "Submitted for EBOD review" : "Draft saved — find it under My drafts");
    onClose();
  };
  return (
    <Sheet open onClose={onClose} title={draft ? "Edit proposal" : "Propose a project"} tall>
      <div className="flex gap-1.5 mb-4">
        {steps.map((s, i) => (
          <button key={s} onClick={() => setStep(i)} className="flex-1 rounded-full font-bold py-1.5" style={{ fontFamily: DISPLAY, fontSize: 12, background: step === i ? CRAN : step > i ? CRAN + "22" : "#fff", color: step === i ? "#fff" : CRAN, border: `1px solid ${step >= i ? "transparent" : LINE}` }}>{i + 1}. {s}</button>
        ))}
      </div>
      {step === 0 && (<>
        <Field label="Project name"><Input value={f.title} onChange={set("title")} /></Field>
        <Field label="Rotary area of focus"><Select value={f.area} onChange={set("area")}>{AREAS.map((a) => <option key={a}>{a}</option>)}</Select></Field>
        <Field label="Description"><TextArea value={f.description} onChange={set("description")} placeholder="What will the club do?" /></Field>
        <Field label="Problem / need being addressed"><TextArea rows={2} value={f.problem} onChange={set("problem")} /></Field>
        <Field label="Objectives"><TextArea rows={2} value={f.objectives} onChange={set("objectives")} placeholder="Measurable objectives" /></Field>
        <Field label="Goals"><TextArea rows={2} value={f.goals} onChange={set("goals")} /></Field>
      </>)}
      {step === 1 && (<>
        <Field label="Expected impact"><TextArea rows={2} value={f.impact} onChange={set("impact")} /></Field>
        <Field label="Target beneficiaries"><Input value={f.beneficiaries} onChange={set("beneficiaries")} placeholder="Who benefits, and roughly how many?" /></Field>
        <Field label="Proposed location"><Input value={f.location} onChange={set("location")} /></Field>
        <Field label="Proposed dates"><Input value={f.dates} onChange={set("dates")} placeholder="e.g. Two Saturdays in September" /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Volunteers needed"><Input type="number" min="0" value={f.volunteersNeeded} onChange={set("volunteersNeeded")} /></Field>
          <Field label="Partner organizations"><Input value={f.partners} onChange={set("partners")} /></Field>
        </div>
        <Field label="Success indicators"><TextArea rows={2} value={f.success} onChange={set("success")} placeholder="How will we know it worked?" /></Field>
        <Field label="Sustainability plan"><TextArea rows={2} value={f.sustainability} onChange={set("sustainability")} /></Field>
      </>)}
      {step === 2 && (<>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Estimated budget"><Input type="number" min="0" value={f.budget} onChange={set("budget")} /></Field>
          <Field label="Expected income / sponsorship"><Input type="number" min="0" value={f.expectedIncome} onChange={set("expectedIncome")} /></Field>
        </div>
        <Field label="Required materials"><TextArea rows={2} value={f.materials} onChange={set("materials")} /></Field>
        <Field label="Risks"><TextArea rows={2} value={f.risks} onChange={set("risks")} /></Field>
        <FilePick label="Supporting documents & images" accept="*/*" onFile={(file) => setF({ ...f, files: [...(f.files || []), file] })} />
        {(f.files || []).length > 0 && <div className="mb-3">{f.files.map((a) => <FileChip key={a.id} f={a} onRemove={() => setF({ ...f, files: f.files.filter((x) => x.id !== a.id) })} />)}</div>}
      </>)}
      <div className="flex gap-2 mt-2">
        {step > 0 && <Btn kind="quiet" onClick={() => setStep(step - 1)} style={{ flex: 1 }}>Back</Btn>}
        {step < 2 ? <Btn onClick={() => setStep(step + 1)} style={{ flex: 2 }} disabled={step === 0 && (!f.title.trim() || !f.description.trim())}>Next</Btn>
          : (<>
            <Btn kind="quiet" onClick={() => save(false)} style={{ flex: 1 }} disabled={!f.title.trim()}>Save draft</Btn>
            <Btn onClick={() => save(true)} style={{ flex: 2 }} disabled={!f.title.trim() || !f.description.trim()}>Submit to EBOD</Btn>
          </>)}
      </div>
    </Sheet>
  );
}

function ProjectDetail({ id, onClose }) {
  const { db, me, patch, isEBOD, memberById, notify, audit, showToast, setOverlay } = useApp();
  const p = db.projects.find((x) => x.id === id);
  const [wtab, setWtab] = useState("overview");
  const [taskFilter, setTaskFilter] = useState("all");
  const [voteOpen, setVoteOpen] = useState(false);
  const [taskOpen, setTaskOpen] = useState(null); // 'new' | task
  const [reportOpen, setReportOpen] = useState(false);
  const [qText, setQText] = useState(""); const [noteText, setNoteText] = useState(""); const [msgText, setMsgText] = useState("");
  const [updText, setUpdText] = useState(""); const [hrs, setHrs] = useState(""); const [riskText, setRiskText] = useState("");
  const [msT, setMsT] = useState(""); const [msD, setMsD] = useState("");
  if (!p) return null;
  const isLead = p.lead === me.id;
  const inReview = ["Submitted", "Under Review"].includes(p.status);
  const cur = db.duesConfig.currency;
  const up = (fn) => patch((d) => fn(d.projects.find((x) => x.id === id), d));

  /* --- EBOD decision logic --- */
  const castVote = (v, comment) => {
    up((pr, d) => {
      pr.votes[me.id] = { v, comment };
      if (pr.status === "Submitted") pr.status = "Under Review";
      const approvals = Object.values(pr.votes).filter((x) => x.v === "approve").length;
      const denials = Object.values(pr.votes).filter((x) => x.v === "deny").length;
      if (approvals >= 3) {
        pr.status = "Approved";
        pr.decisions.push({ id: uid(), by: me.id, action: "Approved", reason: comment || "Majority of EBOD approved.", at: Date.now() });
        notify(d, { type: "projects", title: "Project approved 🎉", body: `"${pr.title}" is approved — ${d.members.find((m) => m.id === pr.lead)?.name} leads the workspace.` });
      } else if (denials >= 3) {
        pr.status = "Denied";
        pr.decisions.push({ id: uid(), by: me.id, action: "Denied", reason: comment || "Majority of EBOD declined.", at: Date.now() });
        notify(d, { type: "projects", title: "Proposal decision", body: `"${pr.title}" was not approved this time.` });
      }
      audit(d, `EBOD vote: ${v}`, pr.title);
    });
    setVoteOpen(false);
  };
  const returnForRevision = (reason) => {
    up((pr, d) => {
      pr.status = "Returned"; pr.votes = {};
      pr.decisions.push({ id: uid(), by: me.id, action: "Returned for revision", reason, at: Date.now() });
      notify(d, { type: "projects", title: "Revision requested", body: `"${pr.title}" was returned to ${d.members.find((m) => m.id === pr.submittedBy)?.name}: ${reason}` });
      audit(d, "Proposal returned", pr.title);
    });
    setVoteOpen(false); showToast("Returned to the proposer");
  };
  const resubmit = () => up((pr, d) => {
    pr.status = "Submitted"; pr.at = Date.now();
    pr.decisions.push({ id: uid(), by: me.id, action: "Resubmitted", reason: "", at: Date.now() });
    notify(d, { type: "projects", title: "Proposal resubmitted", body: `"${pr.title}" is back for EBOD review.` });
  });
  const reassignLead = (mid) => up((pr, d) => {
    pr.lead = mid;
    notify(d, { type: "projects", title: "Project lead assigned", body: `${d.members.find((m) => m.id === mid)?.name} now leads "${pr.title}".` });
    audit(d, "Project lead reassigned", pr.title);
  });
  const closeProject = () => {
    up((pr, d) => {
      pr.status = "Completed"; pr.completedAt = Date.now();
      pr.decisions.push({ id: uid(), by: me.id, action: "Closed", reason: "Final report reviewed by EBOD.", at: Date.now() });
      notify(d, { type: "projects", title: "Project closed", body: `"${pr.title}" is complete — thanks to everyone who served! 🎉` });
      audit(d, "Project closed", pr.title);
    });
  };

  const team = [...new Set([p.lead, ...p.tasks.map((t) => t.assignee), ...p.volunteerRoles.flatMap((r) => r.filled)].filter(Boolean))];
  const projTx = db.transactions.filter((t) => t.projectId === p.id && !t.reversed && (t.type === "income" || t.approval === "approved"));
  const spent = projTx.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
  const earned = projTx.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const totalHours = p.serviceHours.reduce((s, h) => s + Number(h.hours || 0), 0);
  const myVote = p.votes[me.id];
  const approvals = Object.values(p.votes).filter((x) => x.v === "approve").length;

  const askQuestion = () => { if (!qText.trim()) return; up((pr, d) => { pr.questions.push({ id: uid(), by: me.id, text: qText.trim(), at: Date.now(), answer: "" }); notify(d, { type: "projects", title: "EBOD question", body: `On "${pr.title}": ${qText.trim()}` }); }); setQText(""); };
  const answerQ = (qid, ans) => up((pr) => { const q = pr.questions.find((x) => x.id === qid); q.answer = ans; });
  const addNote = () => { if (!noteText.trim()) return; up((pr) => pr.ebodNotes.push({ id: uid(), by: me.id, text: noteText.trim(), at: Date.now() })); setNoteText(""); };
  const addMsg = () => { if (!msgText.trim()) return; up((pr) => pr.discussion.push({ id: uid(), by: me.id, text: msgText.trim(), at: Date.now() })); setMsgText(""); };
  const addUpdate = () => { if (!updText.trim()) return; up((pr, d) => { pr.updates.unshift({ id: uid(), by: me.id, text: updText.trim(), at: Date.now() }); notify(d, { type: "projects", title: "Project update", body: `${p.title}: ${updText.trim().slice(0, 90)}` }); }); setUpdText(""); };
  const logHours = () => { if (!Number(hrs)) return; up((pr) => pr.serviceHours.push({ id: uid(), memberId: me.id, hours: Number(hrs), date: todayStr() })); setHrs(""); showToast("Service hours logged"); };
  const addRisk = () => { if (!riskText.trim()) return; up((pr) => pr.risksLog.push({ id: uid(), by: me.id, text: riskText.trim(), at: Date.now(), resolved: false })); setRiskText(""); };
  const addMilestone = () => { if (!msT.trim()) return; up((pr) => pr.milestones.push({ id: uid(), title: msT.trim(), date: msD, done: false })); setMsT(""); setMsD(""); };
  const joinRole = (rid) => up((pr, d) => {
    const r = pr.volunteerRoles.find((x) => x.id === rid);
    if (r.filled.includes(me.id)) r.filled = r.filled.filter((x) => x !== me.id);
    else if (r.filled.length < r.slots) { r.filled.push(me.id); notify(d, { type: "projects", title: "Volunteer signed up", body: `${me.name} took "${r.title}" on ${pr.title}.` }); }
  });

  return (
    <FullScreen title={p.title} onClose={onClose} accent={p.status === "Approved" ? AZURE : CRAN}>
      <div className="flex gap-1.5 flex-wrap items-center">
        <Badge color={PROJ_STATUS_COLOR[p.status]} filled>{p.status}</Badge>
        <Badge color={AZURE}>{p.area}</Badge>
        <Badge>Lead: {memberById(p.lead)?.name || "—"}</Badge>
      </div>

      {["Approved", "Pending Close", "Completed"].includes(p.status) && (
        <div className="grid grid-cols-4 gap-2 mt-3">
          {[
            ["Tasks", `${p.tasks.filter((t) => t.status === "Completed").length}/${p.tasks.length}`, OK],
            ["Hours", totalHours, AZURE],
            ["Spent", `${cur}${Math.round(spent)}`, spent > p.budget && p.budget ? BAD : "#B07400"],
            ["Team", team.length, CRAN],
          ].map(([k, v, c]) => (
            <Card key={k} className="p-2 text-center">
              <div className="font-black" style={{ fontFamily: DISPLAY, fontSize: 16, color: c }}>{v}</div>
              <div className="font-bold uppercase" style={{ fontSize: 9.5, color: "#8A7580", letterSpacing: ".06em" }}>{k}</div>
            </Card>
          ))}
        </div>
      )}

      {/* ---- Review phase ---- */}
      {(inReview || p.status === "Returned" || p.status === "Denied") && (
        <>
          {inReview && (
            <Card className="mt-3 p-4" style={{ borderColor: GOLD, background: GOLD + "10" }}>
              <div className="font-bold" style={{ fontSize: 14 }}>EBOD review · {approvals}/3 approvals</div>
              <Progress pct={(approvals / 3) * 100} color={GOLD} />
              <div className="mt-2 flex flex-col gap-1">
                {Object.entries(p.votes).map(([mid, v]) => (
                  <div key={mid} style={{ fontSize: 12.5 }}><b>{memberById(mid)?.name}</b> — {v.v === "approve" ? "✅ approve" : "❌ deny"}{v.comment ? `: "${v.comment}"` : ""}</div>
                ))}
              </div>
              {isEBOD && <div className="mt-3"><Btn small onClick={() => setVoteOpen(true)}>{myVote ? "Change my decision" : "Record my decision"}</Btn></div>}
            </Card>
          )}
          {p.status === "Returned" && p.submittedBy === me.id && (
            <Card className="mt-3 p-4" style={{ borderColor: "#7A4BA6" }}>
              <div className="font-bold" style={{ fontSize: 14, color: "#7A4BA6" }}>Returned for revision</div>
              <div style={{ fontSize: 13 }} className="mt-1">{p.decisions.filter((x) => x.action === "Returned for revision").slice(-1)[0]?.reason}</div>
              <div className="mt-3 flex gap-2"><ProposeReopen p={p} /><Btn small onClick={resubmit}>Resubmit as-is</Btn></div>
            </Card>
          )}
        </>
      )}

      {/* ---- Workspace tabs when Approved+ ---- */}
      {["Approved", "Pending Close", "Completed"].includes(p.status) && (
        <div className="mt-3"><Chips options={[{ id: "overview", label: "Overview" }, { id: "tasks", label: `Tasks (${p.tasks.length})` }, { id: "budget", label: "Budget" }, { id: "team", label: "Team" }, { id: "more2", label: "More" }]} value={wtab} onChange={setWtab} /></div>
      )}

      {(wtab === "overview" || !["Approved", "Pending Close", "Completed"].includes(p.status)) && (
        <>
          <SectionTitle>Proposal</SectionTitle>
          <Card className="p-4">
            <KV k="Description" v={p.description} /><KV k="Problem / need" v={p.problem} /><KV k="Objectives" v={p.objectives} /><KV k="Goals" v={p.goals} />
            <KV k="Expected impact" v={p.impact} /><KV k="Beneficiaries" v={p.beneficiaries} /><KV k="Location" v={p.location} /><KV k="Dates" v={p.dates} />
            <KV k="Volunteers needed" v={p.volunteersNeeded ? String(p.volunteersNeeded) : ""} /><KV k="Partners" v={p.partners} />
            <KV k="Budget" v={p.budget ? money(p.budget, cur) : ""} /><KV k="Expected income" v={p.expectedIncome ? money(p.expectedIncome, cur) : ""} />
            <KV k="Materials" v={p.materials} /><KV k="Risks" v={p.risks} /><KV k="Success indicators" v={p.success} /><KV k="Sustainability" v={p.sustainability} />
            {(p.files || []).length > 0 && <div className="mt-1">{p.files.map((a) => <FileChip key={a.id} f={a} />)}</div>}
            <div className="pt-2 mt-1" style={{ borderTop: `1px solid ${LINE}`, fontSize: 12, color: "#8A7580" }}>Proposed by {memberById(p.submittedBy)?.name} · {timeAgo(p.at)}</div>
          </Card>

          {(() => { const evs = db.meetings.filter((x) => x.projectId === p.id && x.status !== "cancelled").sort((a, b) => a.date.localeCompare(b.date)); return evs.length ? (<>
            <SectionTitle>Scheduled dates</SectionTitle>
            <Card className="p-3">
              {evs.map((ev) => (
                <button key={ev.id} onClick={() => setOverlay({ type: "meeting", id: ev.id })} className="w-full flex items-center gap-2.5 py-1.5 text-left" style={{ fontSize: 13.5 }}>
                  <Icon name="calendar" size={15} color={eventKind(ev).color} />
                  <span className="flex-1"><b>{ev.title}</b> · {fmtShort(ev.date)}, {ev.time}</span>
                  <Badge color={eventKind(ev).color}>{eventKind(ev).short}</Badge>
                </button>
              ))}
            </Card>
          </>) : null; })()}

          {isEBOD && (
            <>
              <SectionTitle>EBOD workspace</SectionTitle>
              <Card className="p-4">
                <div className="font-bold" style={{ fontSize: 13 }}>Questions to the proposer</div>
                {p.questions.map((q) => (
                  <div key={q.id} className="mt-2" style={{ fontSize: 13 }}>
                    <b>{memberById(q.by)?.name}:</b> {q.text}
                    {q.answer ? <div style={{ color: "#4A3A44" }} className="mt-0.5">↳ {q.answer}</div>
                      : (p.submittedBy === me.id || isLead) ? <AnswerBox onSave={(a) => answerQ(q.id, a)} /> : <span style={{ color: "#A896A0" }}> — awaiting answer</span>}
                  </div>
                ))}
                <div className="flex gap-2 mt-2"><Input value={qText} onChange={(e) => setQText(e.target.value)} placeholder="Ask a question…" /><Btn small onClick={askQuestion}>Ask</Btn></div>
                <div className="font-bold mt-4" style={{ fontSize: 13 }}>Private EBOD comments <span style={{ fontWeight: 400, color: "#A896A0" }}>(EBOD only)</span></div>
                {p.ebodNotes.map((n) => <div key={n.id} className="mt-1.5" style={{ fontSize: 13 }}><b>{memberById(n.by)?.name}:</b> {n.text}</div>)}
                <div className="flex gap-2 mt-2"><Input value={noteText} onChange={(e) => setNoteText(e.target.value)} placeholder="Private note…" /><Btn small kind="dark" onClick={addNote}>Add</Btn></div>
                <div className="font-bold mt-4" style={{ fontSize: 13 }}>Recommendations</div>
                <RecommendBox value={p.recommendations} onSave={(v) => up((pr) => { pr.recommendations = v; })} />
                {["Approved", "Under Review", "Submitted"].includes(p.status) && (
                  <Field label="Project lead"><Select value={p.lead} onChange={(e) => reassignLead(e.target.value)}>{db.members.filter((m) => ACTIVE_LIKE.includes(m.status)).map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}</Select></Field>
                )}
              </Card>
            </>
          )}
          {p.recommendations && !isEBOD && (
            <Card className="mt-2.5 p-3.5" style={{ borderColor: AZURE + "55" }}><div className="font-bold" style={{ fontSize: 12.5, color: AZURE }}>EBOD recommendations</div><div style={{ fontSize: 13 }}>{p.recommendations}</div></Card>
          )}
          <SectionTitle>Decision history</SectionTitle>
          <Card className="p-3">
            {p.decisions.length ? p.decisions.slice().reverse().map((dd) => (
              <div key={dd.id} className="py-1.5" style={{ fontSize: 12.5, borderBottom: `1px dashed ${LINE}` }}>
                <b>{dd.action}</b> — {memberById(dd.by)?.name} · {timeAgo(dd.at)}{dd.reason ? <div style={{ color: "#6B5A64" }}>"{dd.reason}"</div> : null}
              </div>
            )) : <span style={{ fontSize: 13, color: "#9A8B93" }}>No decisions yet.</span>}
          </Card>
        </>
      )}

      {["Approved", "Pending Close", "Completed"].includes(p.status) && wtab === "tasks" && (
        <>
          <div className="flex items-center gap-2 mt-4">
            <div className="flex-1 min-w-0"><Chips options={[{ id: "all", label: "All" }, { id: "mine", label: "Mine" }, { id: "open", label: "Open" }, { id: "done", label: "Done" }]} value={taskFilter} onChange={setTaskFilter} /></div>
            {isLead && p.status === "Approved" && <Btn small onClick={() => setTaskOpen("new")}><span className="flex items-center gap-1.5"><Icon name="plus" size={14} /> Task</span></Btn>}
          </div>
          {p.tasks.length > 0 && (
            <Card className="p-3 mt-2.5">
              <div className="flex justify-between mb-1.5" style={{ fontSize: 12, color: "#8A7580" }}>
                <span className="font-bold">{p.tasks.filter((t) => t.status === "Completed").length} of {p.tasks.length} tasks completed</span>
                <span>{Math.round((p.tasks.filter((t) => t.status === "Completed").length / p.tasks.length) * 100)}%</span>
              </div>
              <Progress pct={(p.tasks.filter((t) => t.status === "Completed").length / p.tasks.length) * 100} color={OK} />
            </Card>
          )}
          <p className="mt-2 mb-1 flex items-center gap-1.5" style={{ fontSize: 11.5, color: "#9A8B93" }}><Icon name="shield" size={13} /> {isLead ? "You assign tasks; each assignee updates their own." : `${memberById(p.lead)?.name?.split(" ")[0]} (lead) assigns tasks — you can update the ones assigned to you.`}</p>
          <div className="flex flex-col gap-2 mt-1.5">
            {(() => {
              const list = p.tasks
                .filter((tk) => taskFilter === "all" ? true : taskFilter === "mine" ? tk.assignee === me.id : taskFilter === "open" ? tk.status !== "Completed" : tk.status === "Completed")
                .slice().sort((x, y) => (x.status === "Completed") - (y.status === "Completed") || (x.deadline || "9999").localeCompare(y.deadline || "9999"));
              return list.length ? list.map((tk) => {
                const mine = tk.assignee === me.id;
                const overdue = tk.deadline && tk.deadline < todayStr() && tk.status !== "Completed";
                const ckDone = tk.checklist.filter((c) => c.done).length;
                const blocked = tk.deps.some((did) => { const dt = p.tasks.find((x) => x.id === did); return dt && dt.status !== "Completed"; });
                return (
                  <Card key={tk.id} onClick={() => setTaskOpen(tk)} className="p-3" style={{ borderLeft: `4px solid ${overdue ? BAD : TASK_COLOR[tk.status]}`, opacity: tk.status === "Completed" ? 0.72 : 1 }}>
                    <div className="flex items-center gap-2.5">
                      <Avatar m={memberById(tk.assignee)} size={32} />
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-1.5">
                          <span className="font-bold truncate" style={{ fontSize: 14, textDecoration: tk.status === "Completed" ? "line-through" : "none" }}>{tk.title}</span>
                          {mine && tk.status !== "Completed" && <Badge color={AZURE}>Yours</Badge>}
                        </div>
                        <div className="flex items-center gap-1 mt-0.5 flex-wrap" style={{ fontSize: 11.5, color: "#8A7580" }}>
                          <span style={{ color: PRIORITY_COLOR[tk.priority], fontWeight: 700 }}>{tk.priority}</span>
                          <span>· {memberById(tk.assignee)?.name?.split(" ")[0] || "Unassigned"}</span>
                          {tk.deadline && <span style={{ color: overdue ? BAD : "#8A7580", fontWeight: overdue ? 700 : 400 }}>· {overdue ? "overdue" : "due"} {fmtShort(tk.deadline)}</span>}
                          {blocked && tk.status !== "Completed" && <span style={{ color: BAD, fontWeight: 700 }}>· blocked</span>}
                        </div>
                        {tk.checklist.length > 0 && (
                          <div className="mt-1.5 flex items-center gap-2">
                            <div className="flex-1"><Progress pct={(ckDone / tk.checklist.length) * 100} color={OK} /></div>
                            <span style={{ fontSize: 10.5, color: "#8A7580" }}>{ckDone}/{tk.checklist.length}</span>
                          </div>
                        )}
                      </div>
                      <Badge color={TASK_COLOR[tk.status]}>{tk.status === "Not Started" ? "To do" : tk.status === "Awaiting Review" ? "Review" : tk.status === "In Progress" ? "Doing" : tk.status}</Badge>
                    </div>
                  </Card>
                );
              }) : <Empty icon="check" title={taskFilter === "mine" ? "Nothing assigned to you" : "No tasks here"} text={isLead && p.status === "Approved" ? "Break the project into tasks and assign them to volunteers." : ""} />;
            })()}
          </div>
        </>
      )}

      {["Approved", "Pending Close", "Completed"].includes(p.status) && wtab === "budget" && (
        <>
          <SectionTitle>Budget vs actual</SectionTitle>
          <Card className="p-4">
            <div className="grid grid-cols-2 gap-3 text-center">
              <div><div style={{ fontSize: 11.5, color: "#8A7580" }} className="font-bold uppercase">Spent</div><div className="font-black" style={{ fontFamily: DISPLAY, fontSize: 22, color: spent > p.budget && p.budget ? BAD : INK }}>{money(spent, cur)}</div><div style={{ fontSize: 11.5, color: "#8A7580" }}>of {money(p.budget || 0, cur)} budget</div></div>
              <div><div style={{ fontSize: 11.5, color: "#8A7580" }} className="font-bold uppercase">Income</div><div className="font-black" style={{ fontFamily: DISPLAY, fontSize: 22, color: OK }}>{money(earned, cur)}</div><div style={{ fontSize: 11.5, color: "#8A7580" }}>expected {money(p.expectedIncome || 0, cur)}</div></div>
            </div>
            {p.budget > 0 && <div className="mt-3"><Progress pct={(spent / p.budget) * 100} color={spent > p.budget ? BAD : OK} /></div>}
          </Card>
          <SectionTitle>Project transactions</SectionTitle>
          {projTx.length ? (
            <Card className="p-2">{projTx.map((t, i) => (
              <div key={t.id} className="flex items-center gap-3 px-2 py-2" style={{ borderTop: i ? `1px solid ${LINE}` : "none" }}>
                <div className="flex-1"><div className="font-semibold" style={{ fontSize: 13 }}>{t.desc}</div><div style={{ fontSize: 11.5, color: "#8A7580" }}>{t.category} · {fmtShort(t.date)}</div></div>
                <div className="font-bold" style={{ fontSize: 13.5, color: t.type === "income" ? OK : BAD }}>{t.type === "income" ? "+" : "−"}{money(t.amount, cur)}</div>
              </div>
            ))}</Card>
          ) : <Empty icon="wallet" title="No transactions" text="The treasurer links income and expenses to this project from the Ledger." />}
        </>
      )}

      {["Approved", "Pending Close", "Completed"].includes(p.status) && wtab === "team" && (
        <>
          <SectionTitle>Team</SectionTitle>
          <Card className="p-2">
            {team.map((mid, i) => { const mm = memberById(mid); return mm && (
              <div key={mid} className="flex items-center gap-3 px-2 py-2" style={{ borderTop: i ? `1px solid ${LINE}` : "none" }}>
                <Avatar m={mm} size={30} /><span className="flex-1 font-semibold" style={{ fontSize: 13.5 }}>{mm.name}</span>
                {mid === p.lead && <Badge color={CRAN}>Lead</Badge>}
              </div>); })}
          </Card>
          <SectionTitle>Volunteer positions</SectionTitle>
          {p.volunteerRoles.length ? p.volunteerRoles.map((r) => (
            <Card key={r.id} className="p-3.5 mb-2">
              <div className="flex items-center justify-between">
                <div><div className="font-bold" style={{ fontSize: 14 }}>{r.title}</div><div style={{ fontSize: 12, color: "#8A7580" }}>{r.filled.length}/{r.slots} filled{r.filled.length ? " — " + r.filled.map((x) => memberById(x)?.name.split(" ")[0]).join(", ") : ""}</div></div>
                {p.status === "Approved" && <Btn small kind={r.filled.includes(me.id) ? "quiet" : "blue"} onClick={() => joinRole(r.id)} disabled={!r.filled.includes(me.id) && r.filled.length >= r.slots}>{r.filled.includes(me.id) ? "Leave" : "Sign up"}</Btn>}
              </div>
            </Card>
          )) : <Empty icon="users" title="No positions posted" />}
          {isLead && p.status === "Approved" && <AddRoleBox onAdd={(title, slots) => up((pr) => pr.volunteerRoles.push({ id: uid(), title, slots, filled: [] }))} />}
          <SectionTitle>Milestones</SectionTitle>
          {p.milestones.map((ms) => (
            <div key={ms.id} className="flex items-center gap-2.5 mb-1.5" style={{ fontSize: 13.5 }}>
              <button onClick={() => isLead && up((pr) => { const x = pr.milestones.find((y) => y.id === ms.id); x.done = !x.done; })} className="rounded-full flex items-center justify-center" style={{ width: 22, height: 22, border: `2px solid ${ms.done ? OK : "#C9B8C1"}`, background: ms.done ? OK : "transparent" }}>{ms.done && <Icon name="check" size={12} color="#fff" />}</button>
              <span style={{ textDecoration: ms.done ? "line-through" : "none" }}>{ms.title}{ms.date ? ` — ${fmtShort(ms.date)}` : ""}</span>
            </div>
          ))}
          {isLead && p.status === "Approved" && (
            <div className="flex gap-2 mt-1"><Input value={msT} onChange={(e) => setMsT(e.target.value)} placeholder="Milestone" /><Input type="date" value={msD} onChange={(e) => setMsD(e.target.value)} style={{ width: 140 }} /><Btn small onClick={addMilestone}>Add</Btn></div>
          )}
        </>
      )}

      {["Approved", "Pending Close", "Completed"].includes(p.status) && wtab === "more2" && (
        <>
          <SectionTitle>Progress updates</SectionTitle>
          {p.status === "Approved" && <div className="flex gap-2 mb-2"><Input value={updText} onChange={(e) => setUpdText(e.target.value)} placeholder="Share an update with the club…" /><Btn small onClick={addUpdate}>Post</Btn></div>}
          {p.updates.map((u) => <Card key={u.id} className="p-3 mb-1.5"><div style={{ fontSize: 13.5 }}>{u.text}</div><div style={{ fontSize: 11.5, color: "#8A7580" }} className="mt-1">{memberById(u.by)?.name} · {timeAgo(u.at)}</div></Card>)}
          <SectionTitle>Discussion</SectionTitle>
          <Card className="p-3">
            {p.discussion.map((mg) => <div key={mg.id} className="mb-2" style={{ fontSize: 13.5 }}><b>{memberById(mg.by)?.name.split(" ")[0]}:</b> {mg.text} <span style={{ fontSize: 11, color: "#A896A0" }}>{timeAgo(mg.at)}</span></div>)}
            <div className="flex gap-2"><Input value={msgText} onChange={(e) => setMsgText(e.target.value)} placeholder="Message the team…" /><Btn small onClick={addMsg}>Send</Btn></div>
          </Card>
          <SectionTitle>Service hours · {totalHours} h total</SectionTitle>
          <Card className="p-3">
            {p.serviceHours.slice(-6).map((h) => <div key={h.id} style={{ fontSize: 13 }}>{memberById(h.memberId)?.name} — {h.hours} h · {fmtShort(h.date)}</div>)}
            {p.status === "Approved" && <div className="flex gap-2 mt-2"><Input type="number" min="0" step="0.5" value={hrs} onChange={(e) => setHrs(e.target.value)} placeholder="Hours served" /><Btn small kind="ok" onClick={logHours}>Log mine</Btn></div>}
          </Card>
          <SectionTitle>Risks & issues</SectionTitle>
          <Card className="p-3">
            {p.risksLog.map((r) => (
              <div key={r.id} className="flex items-center gap-2 mb-1.5" style={{ fontSize: 13 }}>
                <button onClick={() => up((pr) => { const x = pr.risksLog.find((y) => y.id === r.id); x.resolved = !x.resolved; })}><Icon name={r.resolved ? "check" : "flag"} size={14} color={r.resolved ? OK : BAD} /></button>
                <span style={{ textDecoration: r.resolved ? "line-through" : "none", flex: 1 }}>{r.text}</span>
              </div>
            ))}
            <div className="flex gap-2"><Input value={riskText} onChange={(e) => setRiskText(e.target.value)} placeholder="Raise a risk or issue…" /><Btn small kind="quiet" onClick={addRisk}>Raise</Btn></div>
          </Card>
          <SectionTitle>Files</SectionTitle>
          <Card className="p-3">
            {(p.files || []).length ? p.files.map((a) => <FileChip key={a.id} f={a} onRemove={isLead ? () => up((pr) => { pr.files = pr.files.filter((x) => x.id !== a.id); }) : null} />) : <span style={{ fontSize: 13, color: "#9A8B93" }}>No files.</span>}
            {p.status === "Approved" && <FilePick accept="*/*" onFile={(file) => up((pr) => pr.files.push(file))} />}
          </Card>
        </>
      )}

      {/* ---- Completion ---- */}
      {p.status === "Approved" && isLead && (
        <div className="mt-5"><Btn kind="dark" onClick={() => setReportOpen(true)}>Complete final report → request closure</Btn></div>
      )}
      {p.status === "Pending Close" && (
        <>
          <SectionTitle>Final report</SectionTitle>
          <FinalReportView p={p} />
          {isEBOD && <div className="mt-3"><Btn kind="ok" onClick={closeProject}>EBOD: accept report & close project</Btn></div>}
          {!isEBOD && <p className="mt-2" style={{ fontSize: 13, color: "#8A7580" }}>Awaiting EBOD review of the final report.</p>}
        </>
      )}
      {p.status === "Completed" && p.finalReport && (<><SectionTitle>Final report</SectionTitle><FinalReportView p={p} /></>)}

      {voteOpen && <VoteSheet onVote={castVote} onReturn={returnForRevision} onClose={() => setVoteOpen(false)} />}
      {taskOpen && <TaskSheet project={p} task={taskOpen === "new" ? null : taskOpen} onClose={() => setTaskOpen(null)} />}
      {reportOpen && <FinalReportSheet project={p} onClose={() => setReportOpen(false)} />}
    </FullScreen>
  );
}

function RecommendBox({ value, onSave }) {
  const [v, setV] = useState(value || "");
  return (<div className="mt-1.5">
    <TextArea rows={2} value={v} onChange={(e) => setV(e.target.value)} placeholder="Conditions or advice to the project team…" />
    {v !== (value || "") && <div className="mt-1.5"><Btn small kind="quiet" onClick={() => onSave(v.trim())}>Save recommendations</Btn></div>}
  </div>);
}
function AnswerBox({ onSave }) {
  const [v, setV] = useState("");
  return <div className="flex gap-2 mt-1"><Input value={v} onChange={(e) => setV(e.target.value)} placeholder="Answer…" /><Btn small kind="quiet" onClick={() => { if (v.trim()) { onSave(v.trim()); setV(""); } }}>Reply</Btn></div>;
}
function AddRoleBox({ onAdd }) {
  const [t, setT] = useState(""); const [s, setS] = useState(2);
  return <div className="flex gap-2 mb-2"><Input value={t} onChange={(e) => setT(e.target.value)} placeholder="Position title" /><Input type="number" min="1" value={s} onChange={(e) => setS(e.target.value)} style={{ width: 70 }} /><Btn small onClick={() => { if (t.trim()) { onAdd(t.trim(), +s || 1); setT(""); } }}>Add</Btn></div>;
}
function VoteSheet({ onVote, onReturn, onClose }) {
  const [comment, setComment] = useState("");
  return (
    <Sheet open onClose={onClose} title="EBOD decision">
      <Field label="Reason / comment" hint="Recorded in the decision history."><TextArea value={comment} onChange={(e) => setComment(e.target.value)} /></Field>
      <div className="flex flex-col gap-2">
        <Btn kind="ok" onClick={() => onVote("approve", comment.trim())}>Approve</Btn>
        <Btn kind="danger" onClick={() => onVote("deny", comment.trim())}>Deny</Btn>
        <Btn kind="quiet" onClick={() => { if (comment.trim()) onReturn(comment.trim()); }} disabled={!comment.trim()}>Return for revision (reason required)</Btn>
      </div>
      <p className="mt-3" style={{ fontSize: 12.5, color: "#9A8B93" }}>Three approvals pass a proposal; three denials decline it.</p>
    </Sheet>
  );
}

function TaskSheet({ project, task, onClose }) {
  const { db, patch, me, memberById, notify, audit, showToast } = useApp();
  const isLead = project.lead === me.id;
  const [f, setF] = useState(task ? clone(task) : { title: "", assignee: project.lead, deadline: "", priority: "Medium", status: "Not Started", checklist: [], comments: [], deps: [], attachments: [] });
  const [ck, setCk] = useState(""); const [cm, setCm] = useState("");
  const others = project.tasks.filter((t) => !task || t.id !== task.id);
  const blockedBy = f.deps.map((d) => project.tasks.find((t) => t.id === d)).filter((t) => t && t.status !== "Completed");
  const canEdit = isLead;                          // only the project lead assigns & structures tasks
  const canWork = isLead || f.assignee === me.id;  // only the assignee (or lead) progresses the task
  const team = [...new Set([project.lead, ...project.volunteerRoles.flatMap((r) => r.filled), ...project.tasks.map((t) => t.assignee)])].filter(Boolean);
  const assignables = db.members.filter((m) => ACTIVE_LIKE.includes(m.status)).slice().sort((x, y) => (team.includes(y.id) ? 1 : 0) - (team.includes(x.id) ? 1 : 0));

  const save = () => {
    if (!f.title.trim() || (!canWork && !canEdit)) return;
    patch((d) => {
      const pr = d.projects.find((x) => x.id === project.id);
      const assigneeChanged = !task || task.assignee !== f.assignee;
      const deadlineChanged = task && task.deadline !== f.deadline;
      if (task) {
        const t = pr.tasks.find((x) => x.id === task.id);
        const known = new Set(f.comments.map((c) => c.id));
        Object.assign(t, { ...f, comments: [...f.comments, ...t.comments.filter((c) => !known.has(c.id))] });
      } else pr.tasks.push({ ...f, id: uid(), title: f.title.trim() });
      if (assigneeChanged && f.assignee && canEdit) notify(d, { type: "tasks", title: "Task assigned", body: `${memberById(f.assignee)?.name}: "${f.title.trim()}" on ${pr.title}${f.deadline ? `, due ${fmtShort(f.deadline)}` : ""}.` });
      else if (deadlineChanged) notify(d, { type: "tasks", title: "Deadline changed", body: `"${f.title.trim()}" on ${pr.title} is now due ${f.deadline ? fmtShort(f.deadline) : "—"}.` });
      audit(d, task ? "Task updated" : "Task created", `${pr.title} · ${f.title.trim()}`);
    });
    showToast(task ? "Task updated" : "Task created & assignee notified"); onClose();
  };
  const setStatus = (s) => {
    if (!canWork) return;
    if (s !== "Blocked" && s !== "Not Started" && blockedBy.length && s !== f.status) { showToast(`Finish first: ${blockedBy.map((t) => t.title).join(", ")}`); return; }
    setF({ ...f, status: s });
  };
  const postComment = () => {
    if (!cm.trim()) return;
    const c = { id: uid(), by: me.id, text: cm.trim(), at: Date.now() };
    if (task) patch((d) => { const pr = d.projects.find((x) => x.id === project.id); pr.tasks.find((x) => x.id === task.id).comments.push(c); });
    setF({ ...f, comments: [...f.comments, c] }); setCm("");
  };

  return (
    <Sheet open onClose={onClose} title={task ? "Task" : "New task"} tall>
      {task && !isLead && (
        <Card className="p-3 mb-4 flex items-start gap-2.5" style={{ background: (canWork ? AZURE : "#8A7580") + "0e", borderColor: canWork ? AZURE + "55" : LINE }}>
          <div style={{ color: canWork ? AZURE : "#8A7580" }} className="mt-0.5"><Icon name={canWork ? "check" : "lock"} size={16} /></div>
          <div style={{ fontSize: 12.5, color: canWork ? "#134B7A" : "#6B5A64" }}>
            {canWork ? <><b>This task is assigned to you.</b> Update the status and checklist as you go — the lead handles reassignment and deadlines.</>
              : <><b>View only.</b> {memberById(project.lead)?.name?.split(" ")[0]} (lead) assigns tasks, and only {memberById(f.assignee)?.name?.split(" ")[0] || "the assignee"} can update this one. You're welcome to comment.</>}
          </div>
        </Card>
      )}

      <Field label="Title"><Input value={f.title} onChange={(e) => canEdit && setF({ ...f, title: e.target.value })} disabled={!canEdit} /></Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Assignee" hint={canEdit ? "Team members listed first." : ""}>
          <Select value={f.assignee} onChange={(e) => setF({ ...f, assignee: e.target.value })} disabled={!canEdit}>
            {assignables.map((m) => <option key={m.id} value={m.id}>{m.name}{team.includes(m.id) ? " ★" : ""}</option>)}
          </Select>
        </Field>
        <Field label="Deadline"><Input type="date" value={f.deadline} onChange={(e) => setF({ ...f, deadline: e.target.value })} disabled={!canEdit} /></Field>
      </div>
      <Field label="Priority">
        <div className="flex gap-2">{PRIORITIES.map((pr) => (
          <button key={pr} disabled={!canEdit} onClick={() => setF({ ...f, priority: pr })} className="flex-1 rounded-xl font-bold py-2" style={{ fontFamily: DISPLAY, fontSize: 13, background: f.priority === pr ? PRIORITY_COLOR[pr] : "#fff", color: f.priority === pr ? "#fff" : "#6B5A64", border: `1.5px solid ${f.priority === pr ? "transparent" : LINE}`, opacity: canEdit ? 1 : 0.65 }}>{pr}</button>
        ))}</div>
      </Field>

      <Field label="Status" hint={blockedBy.length ? `Waiting on: ${blockedBy.map((t) => t.title).join(", ")}` : ""}>
        <div className="flex flex-col gap-1.5">
          {TASK_STATUSES.map((s, i) => {
            const active = f.status === s;
            return (
              <button key={s} disabled={!canWork} onClick={() => setStatus(s)} className="flex items-center gap-2.5 rounded-xl px-3 py-2 text-left"
                style={{ background: active ? TASK_COLOR[s] + "12" : "#fff", border: `1.5px solid ${active ? TASK_COLOR[s] : LINE}`, opacity: canWork ? 1 : 0.6 }}>
                <span className="rounded-full flex items-center justify-center font-bold flex-shrink-0" style={{ width: 22, height: 22, background: active ? TASK_COLOR[s] : "#F0EAEE", color: active ? "#fff" : "#8A7580", fontSize: 11 }}>{i + 1}</span>
                <span className="font-bold" style={{ fontSize: 13.5, color: active ? TASK_COLOR[s] : "#6B5A64" }}>{s}</span>
                {active && <span className="ml-auto"><Icon name="check" size={15} color={TASK_COLOR[s]} /></span>}
              </button>
            );
          })}
        </div>
      </Field>

      {canEdit && others.length > 0 && (
        <Field label="Depends on" hint="This task can't progress until these are completed.">
          <div className="flex gap-1.5 flex-wrap">{others.map((t) => (
            <button key={t.id} onClick={() => setF({ ...f, deps: f.deps.includes(t.id) ? f.deps.filter((x) => x !== t.id) : [...f.deps, t.id] })} className="rounded-full font-semibold" style={{ fontSize: 11.5, padding: "5px 10px", background: f.deps.includes(t.id) ? INK : "#fff", color: f.deps.includes(t.id) ? "#fff" : "#6B5A64", border: `1px solid ${LINE}` }}>{t.title}</button>
          ))}</div>
        </Field>
      )}
      {!canEdit && f.deps.length > 0 && (
        <Field label="Depends on"><div className="flex gap-1.5 flex-wrap">{f.deps.map((did) => { const dt = project.tasks.find((x) => x.id === did); return dt && <Badge key={did} color={dt.status === "Completed" ? OK : BAD}>{dt.title}</Badge>; })}</div></Field>
      )}

      <SectionTitle>Checklist{f.checklist.length ? ` · ${f.checklist.filter((c) => c.done).length}/${f.checklist.length}` : ""}</SectionTitle>
      {f.checklist.map((c) => (
        <div key={c.id} className="flex items-center gap-2.5 mb-1.5">
          <button disabled={!canWork} onClick={() => setF({ ...f, checklist: f.checklist.map((x) => x.id === c.id ? { ...x, done: !x.done } : x) })} className="rounded-full flex items-center justify-center" style={{ width: 22, height: 22, border: `2px solid ${c.done ? OK : "#C9B8C1"}`, background: c.done ? OK : "transparent", opacity: canWork ? 1 : 0.6 }}>{c.done && <Icon name="check" size={12} color="#fff" />}</button>
          <span style={{ fontSize: 13.5, textDecoration: c.done ? "line-through" : "none", flex: 1, color: c.done ? "#A896A0" : INK }}>{c.text}</span>
          {canEdit && <button onClick={() => setF({ ...f, checklist: f.checklist.filter((x) => x.id !== c.id) })} aria-label="Remove"><Icon name="x" size={13} color="#A896A0" /></button>}
        </div>
      ))}
      {canWork && <div className="flex gap-2"><Input value={ck} onChange={(e) => setCk(e.target.value)} placeholder="Checklist item" /><Btn small kind="quiet" onClick={() => { if (ck.trim()) { setF({ ...f, checklist: [...f.checklist, { id: uid(), text: ck.trim(), done: false }] }); setCk(""); } }}>Add</Btn></div>}

      <SectionTitle>Comments</SectionTitle>
      {f.comments.length === 0 && <p style={{ fontSize: 12.5, color: "#9A8B93" }} className="mb-1">No comments yet.</p>}
      {f.comments.map((c) => (
        <div key={c.id} className="flex gap-2 mb-2">
          <Avatar m={memberById(c.by)} size={26} />
          <div className="rounded-2xl px-3 py-1.5" style={{ background: "#fff", border: `1px solid ${LINE}`, fontSize: 13 }}>
            <b>{memberById(c.by)?.name.split(" ")[0]}</b> {c.text} <span style={{ fontSize: 10.5, color: "#A896A0" }}>· {timeAgo(c.at)}</span>
          </div>
        </div>
      ))}
      <div className="flex gap-2"><Input value={cm} onChange={(e) => setCm(e.target.value)} placeholder="Comment…" /><Btn small kind="quiet" onClick={postComment}>Post</Btn></div>

      {canWork && (
        <>
          <FilePick label="Attachments" accept="*/*" onFile={(file) => setF({ ...f, attachments: [...f.attachments, file] })} />
          {f.attachments.length > 0 && <div className="mb-2">{f.attachments.map((a) => <FileChip key={a.id} f={a} onRemove={canEdit ? () => setF({ ...f, attachments: f.attachments.filter((x) => x.id !== a.id) }) : null} />)}</div>}
        </>
      )}
      {!canWork && f.attachments.length > 0 && <div className="mb-2 mt-3">{f.attachments.map((a) => <FileChip key={a.id} f={a} />)}</div>}

      {(canWork || canEdit) && <div className="mt-3"><Btn onClick={save} disabled={!f.title.trim()}>{task ? "Save changes" : "Create & assign task"}</Btn></div>}
      {task && canEdit && <div className="mt-2"><Btn kind="danger" onClick={() => { patch((d) => { const pr = d.projects.find((x) => x.id === project.id); pr.tasks = pr.tasks.filter((t) => t.id !== task.id); }); onClose(); }}>Delete task</Btn></div>}
    </Sheet>
  );
}

function FinalReportSheet({ project, onClose }) {
  const { patch, me, notify, audit, showToast } = useApp();
  const [f, setF] = useState({ activities: "", objectivesAchieved: "", beneficiaries: "", volunteers: "", serviceHours: String(project.serviceHours.reduce((s, h) => s + Number(h.hours || 0), 0)), expenditure: "", challenges: "", lessons: "", recommendations: "", communityImpact: "", photos: [] });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const submit = () => {
    patch((d) => {
      const pr = d.projects.find((x) => x.id === project.id);
      pr.finalReport = { ...f, by: me.id, at: Date.now() };
      pr.status = "Pending Close";
      pr.decisions.push({ id: uid(), by: me.id, action: "Final report submitted", reason: "", at: Date.now() });
      notify(d, { type: "projects", title: "Final report submitted", body: `"${pr.title}" is ready for EBOD closure review.` });
      audit(d, "Final report submitted", pr.title);
    });
    showToast("Sent to EBOD for closure"); onClose();
  };
  return (
    <Sheet open onClose={onClose} title="Final report" tall>
      <Field label="Activities completed"><TextArea value={f.activities} onChange={set("activities")} /></Field>
      <Field label="Objectives achieved"><TextArea rows={2} value={f.objectivesAchieved} onChange={set("objectivesAchieved")} /></Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Number of beneficiaries"><Input type="number" min="0" value={f.beneficiaries} onChange={set("beneficiaries")} /></Field>
        <Field label="Volunteers who took part"><Input type="number" min="0" value={f.volunteers} onChange={set("volunteers")} /></Field>
        <Field label="Total service hours"><Input type="number" min="0" value={f.serviceHours} onChange={set("serviceHours")} /></Field>
        <Field label="Final expenditure"><Input type="number" min="0" value={f.expenditure} onChange={set("expenditure")} /></Field>
      </div>
      <Field label="Community impact"><TextArea rows={2} value={f.communityImpact} onChange={set("communityImpact")} /></Field>
      <Field label="Challenges"><TextArea rows={2} value={f.challenges} onChange={set("challenges")} /></Field>
      <Field label="Lessons learned"><TextArea rows={2} value={f.lessons} onChange={set("lessons")} /></Field>
      <Field label="Recommendations for next time"><TextArea rows={2} value={f.recommendations} onChange={set("recommendations")} /></Field>
      <FilePick label="Photos" accept="image/*" onFile={(file) => setF({ ...f, photos: [...f.photos, file] })} />
      {f.photos.length > 0 && <div className="mb-3 flex flex-wrap gap-2">{f.photos.map((ph) => <img key={ph.id} src={ph.dataUrl} alt="" className="rounded-xl object-cover" style={{ width: 74, height: 74 }} />)}</div>}
      <Btn onClick={submit}>Submit report — request EBOD closure</Btn>
    </Sheet>
  );
}
function FinalReportView({ p }) {
  const { memberById } = useApp();
  const r = p.finalReport;
  if (!r) return null;
  return (
    <Card className="p-4">
      <KV k="Activities completed" v={r.activities} /><KV k="Objectives achieved" v={r.objectivesAchieved} />
      <div className="grid grid-cols-2 gap-2 my-2">
        {[["Beneficiaries", r.beneficiaries], ["Volunteers", r.volunteers], ["Service hours", r.serviceHours], ["Expenditure", r.expenditure]].map(([k, v]) => v !== "" && (
          <div key={k} className="rounded-xl p-2.5 text-center" style={{ background: PAPER }}><div className="font-black" style={{ fontFamily: DISPLAY, fontSize: 18 }}>{v}</div><div style={{ fontSize: 11, color: "#8A7580" }}>{k}</div></div>
        ))}
      </div>
      <KV k="Community impact" v={r.communityImpact} /><KV k="Challenges" v={r.challenges} /><KV k="Lessons learned" v={r.lessons} /><KV k="Recommendations" v={r.recommendations} />
      {r.photos?.length > 0 && <div className="flex flex-wrap gap-2 mt-1">{r.photos.map((ph) => <img key={ph.id} src={ph.dataUrl} alt="" className="rounded-xl object-cover" style={{ width: 74, height: 74 }} />)}</div>}
      <div className="pt-2 mt-2" style={{ borderTop: `1px solid ${LINE}`, fontSize: 12, color: "#8A7580" }}>Filed by {memberById(r.by)?.name} · {timeAgo(r.at)}</div>
    </Card>
  );
}
function ProposeReopen({ p }) {
  const [open, setOpen] = useState(false);
  return (<>
    <Btn small kind="quiet" onClick={() => setOpen(true)}>Edit & revise</Btn>
    {open && <ProposeSheet draft={p} onClose={() => setOpen(false)} />}
  </>);
}

/* ============ Members ============ */
function MembersScreen({ onClose }) {
  const { db, year, isEBOD, patch, me, notify, audit, showToast, setOverlay, memberById } = useApp();
  const [inviteOpen, setInviteOpen] = useState(false);
  const [iv, setIv] = useState({ name: "", email: "" });
  const groups = [
    ["Executive board", db.members.filter((m) => EBOD.includes(m.role) && ACTIVE_LIKE.includes(m.status))],
    ["Members", db.members.filter((m) => !EBOD.includes(m.role) && ACTIVE_LIKE.includes(m.status))],
    ["Prospects & applications", db.members.filter((m) => ["Prospect", "Applied", "Invited"].includes(m.status))],
    ["Former / inactive", db.members.filter((m) => ["Transferred", "Resigned", "Alumni", "Inactive"].includes(m.status))],
  ];
  const invite = () => {
    if (!iv.name.trim()) return;
    patch((d) => {
      d.members.push({ id: uid(), name: iv.name.trim(), email: iv.email.trim(), phone: "", role: "Prospect", status: "Invited", joined: todayStr(), ...blankMemberExtras() });
      notify(d, { type: "membership", title: "Prospect invited", body: `${iv.name.trim()} was invited to visit the club.` });
      audit(d, "Prospect invited", iv.name.trim());
    });
    setIv({ name: "", email: "" }); setInviteOpen(false); showToast("Invitation recorded");
  };
  const exportReport = (fmt) => {
    const rows = [["Name", "Role", "Status", "Email", "Phone", "Joined", "Rotary ID", "Attendance %", "Service hours", "Dues balance", "Overdue", "Committees"]];
    db.members.forEach((m) => {
      const att = attendanceStats(db, m.id, year.id);
      const a = memberAccount(db, m.id, year.id);
      const comms = (year.committees || []).filter((c) => c.members.includes(m.id)).map((c) => c.name).join("; ");
      rows.push([m.name, m.role, m.status, m.email, m.phone, m.joined, m.rotaryId, att.pct ?? "", serviceHoursOf(db, m.id), a.balance.toFixed(2), a.overdue.toFixed(2), comms]);
    });
    if (fmt === "csv") exportCsv(`Membership-${year.id}`, rows); else exportXlsx(`Membership-${year.id}`, rows);
  };
  return (
    <FullScreen title="Members" onClose={onClose}
      right={isEBOD ? <button onClick={() => setInviteOpen(true)} className="rounded-full p-1.5" style={{ background: "rgba(255,255,255,.18)" }} aria-label="Invite"><Icon name="plus" size={18} color="#fff" /></button> : null}>
      {isEBOD && <div className="flex gap-1.5 mb-1"><Btn small kind="quiet" onClick={() => exportReport("csv")}>Export CSV</Btn><Btn small kind="quiet" onClick={() => exportReport("xlsx")}>Export Excel</Btn></div>}
      {groups.map(([label, list]) => list.length > 0 && (
        <div key={label}>
          <SectionTitle>{label} · {list.length}</SectionTitle>
          <div className="flex flex-col gap-2">
            {list.map((m) => {
              const att = attendanceStats(db, m.id, year.id);
              return (
                <Card key={m.id} onClick={() => setOverlay({ type: "memberProfile", id: m.id })} className="flex items-center gap-3 p-3">
                  <Avatar m={m} />
                  <div className="flex-1 min-w-0">
                    <div className="font-bold truncate" style={{ fontSize: 14.5 }}>{m.name}</div>
                    <div style={{ fontSize: 12 }} className="font-semibold"><span style={{ color: ROLE_COLOR[m.role] }}>{m.role}</span>{m.status !== "Active" && <span style={{ color: "#8A7580" }}> · {m.status}</span>}{att.pct !== null && <span style={{ color: "#8A7580" }}> · {att.pct}% att.</span>}</div>
                  </div>
                  <Icon name="chevR" size={18} color="#C9B8C1" />
                </Card>
              );
            })}
          </div>
        </div>
      ))}
      <Sheet open={inviteOpen} onClose={() => setInviteOpen(false)} title="Invite a prospect">
        <Field label="Full name"><Input value={iv.name} onChange={(e) => setIv({ ...iv, name: e.target.value })} /></Field>
        <Field label="Email"><Input value={iv.email} onChange={(e) => setIv({ ...iv, email: e.target.value })} /></Field>
        <Btn onClick={invite} disabled={!iv.name.trim()}>Record invitation</Btn>
      </Sheet>
    </FullScreen>
  );
}

function MemberProfile({ id, onClose }) {
  const { db, year, me, isEBOD, patch, memberById, notify, audit, showToast, setOverlay } = useApp();
  const m = db.members.find((x) => x.id === id);
  const [edit, setEdit] = useState(false);
  const [f, setF] = useState(m ? clone(m) : null);
  const [statusOpen, setStatusOpen] = useState(false);
  if (!m) return null;
  const self = me.id === m.id;
  const canEdit = self || isEBOD;
  const att = attendanceStats(db, m.id, year.id);
  const acct = memberAccount(db, m.id, year.id);
  const hours = serviceHoursOf(db, m.id);
  const cur = db.duesConfig.currency;
  const comms = (year.committees || []).filter((c) => c.members.includes(m.id));
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const saveProfile = () => {
    patch((d) => { const mm = d.members.find((x) => x.id === id); Object.assign(mm, { ...f, id: mm.id, role: mm.role, status: mm.status, positions: mm.positions, statusHistory: mm.statusHistory }); audit(d, "Profile updated", mm.name); });
    setEdit(false); showToast("Profile saved");
  };
  const changeRole = (role) => patch((d) => {
    const mm = d.members.find((x) => x.id === id);
    audit(d, "Role assigned", `${mm.name}: ${mm.role} → ${role}`);
    notify(d, { type: "membership", title: "Club role assigned", body: `${mm.name} is now ${role}.` });
    mm.role = role;
  });
  const changeStatus = (status, note) => {
    patch((d) => {
      const mm = d.members.find((x) => x.id === id);
      mm.statusHistory.push({ id: uid(), from: mm.status, to: status, note, at: Date.now(), by: me.name });
      mm.status = status;
      if (status === "Active" && mm.role === "Prospect") mm.role = "Member";
      notify(d, { type: "membership", title: "Membership update", body: `${mm.name}: ${status}${note ? ` — ${note}` : ""}.` });
      audit(d, "Membership status changed", `${mm.name} → ${status}`);
    });
    setStatusOpen(false); showToast(`Status set to ${status}`);
  };
  const toggleCommittee = (cid) => patch((d) => {
    const y = d.years.find((x) => x.id === d.activeYearId);
    const c = y.committees.find((x) => x.id === cid);
    if (c.members.includes(id)) c.members = c.members.filter((x) => x !== id);
    else c.members.push(id);
    audit(d, "Committee membership changed", `${m.name} · ${c.name}`);
  });
  return (
    <FullScreen title={m.name} onClose={onClose}
      right={canEdit ? <button onClick={() => { setF(clone(m)); setEdit(true); }} className="rounded-full p-1.5" style={{ background: "rgba(255,255,255,.18)" }} aria-label="Edit"><Icon name="edit" size={17} color="#fff" /></button> : null}>
      <div className="flex items-center gap-4">
        <Avatar m={m} size={72} />
        <div>
          <div className="font-black" style={{ fontFamily: DISPLAY, fontSize: 20 }}>{m.name}</div>
          <div className="flex gap-1.5 mt-1 flex-wrap"><Badge color={ROLE_COLOR[m.role]} filled>{m.role}</Badge><Badge color={m.status === "Active" ? OK : GOLD}>{m.status}</Badge></div>
        </div>
      </div>
      <div className="grid grid-cols-3 gap-2 mt-4">
        {[["Attendance", att.pct === null ? "—" : att.pct + "%"], ["Service hrs", hours], ["Dues", acct.overdue > 0 ? "Overdue" : acct.balance > 0 ? "Owing" : "Good ✓"]].map(([k, v]) => (
          <Card key={k} className="p-3 text-center"><div className="font-black" style={{ fontFamily: DISPLAY, fontSize: 17, color: k === "Dues" ? (acct.overdue > 0 ? BAD : acct.balance > 0 ? "#8A5A00" : OK) : INK }}>{v}</div><div style={{ fontSize: 10.5, color: "#8A7580" }} className="font-bold uppercase">{k}</div></Card>
        ))}
      </div>
      {(self || isEBOD) && <div className="mt-2.5"><Btn kind="ghost" small onClick={() => setOverlay({ type: "myAccount", id: m.id })}>View dues account & receipts</Btn></div>}

      <SectionTitle>Profile</SectionTitle>
      <Card className="p-4">
        <KV k="Email" v={m.email} /><KV k="Telephone" v={m.phone} /><KV k="Date of birth" v={m.dob && fmtDate(m.dob)} />
        <KV k="Address" v={m.address} /><KV k="Occupation" v={m.occupation} /><KV k="Employer" v={m.employer} />
        <KV k="Emergency contact" v={m.emergency} /><KV k="Date joined" v={fmtDate(m.joined)} /><KV k="Rotary ID" v={m.rotaryId} />
        <KV k="Areas of interest" v={m.interests} /><KV k="Skills" v={m.skills} />
        {comms.length > 0 && <KV k="Committees" v={comms.map((c) => c.name + (c.chair === m.id ? " (chair)" : "")).join(", ")} />}
      </Card>

      <SectionTitle>Positions held</SectionTitle>
      <Card className="p-3">
        <div style={{ fontSize: 13.5 }}><b>{year.label}:</b> {m.role}</div>
        {m.positions.map((po, i) => <div key={i} style={{ fontSize: 13.5, color: "#6B5A64" }}><b>{db.years.find((y) => y.id === po.yearId)?.label || po.yearId}:</b> {po.role}</div>)}
      </Card>

      {m.statusHistory.length > 0 && (
        <><SectionTitle>Status history</SectionTitle>
        <Card className="p-3">{m.statusHistory.slice().reverse().map((s) => <div key={s.id} className="py-1" style={{ fontSize: 12.5, borderBottom: `1px dashed ${LINE}` }}>{s.from} → <b>{s.to}</b>{s.note ? ` — ${s.note}` : ""} <span style={{ color: "#A896A0" }}>· {s.by} · {timeAgo(s.at)}</span></div>)}</Card></>
      )}

      {isEBOD && !self && (
        <>
          <SectionTitle>EBOD actions</SectionTitle>
          <Card className="p-4">
            {["Applied", "Invited", "Prospect"].includes(m.status) && (
              <div className="mb-3"><Btn kind="ok" onClick={() => changeStatus("Active", "Application approved")}>Approve as active member</Btn></div>
            )}
            <Field label="Club role"><Select value={m.role} onChange={(e) => changeRole(e.target.value)}>{ROLES.map((r) => <option key={r}>{r}</option>)}</Select></Field>
            {(year.committees || []).length > 0 && (
              <Field label="Committees">
                <div className="flex gap-1.5 flex-wrap">{year.committees.map((c) => (
                  <button key={c.id} onClick={() => toggleCommittee(c.id)} className="rounded-full font-semibold" style={{ fontSize: 12, padding: "5px 11px", background: c.members.includes(m.id) ? AZURE : "#fff", color: c.members.includes(m.id) ? "#fff" : "#6B5A64", border: `1px solid ${LINE}` }}>{c.name}</button>
                ))}</div>
              </Field>
            )}
            <Btn kind="quiet" onClick={() => setStatusOpen(true)}>Change membership status…</Btn>
          </Card>
        </>
      )}

      <Sheet open={edit} onClose={() => setEdit(false)} title="Edit profile" tall>
        {f && (<>
          <FilePick label="Profile photo" accept="image/*" hint="A small square photo works best." onFile={(file) => setF({ ...f, photo: file.dataUrl })} />
          {f.photo && <div className="mb-3 flex items-center gap-2"><img src={f.photo} alt="" className="rounded-full object-cover" style={{ width: 56, height: 56 }} /><Btn small kind="quiet" onClick={() => setF({ ...f, photo: null })}>Remove</Btn></div>}
          <Field label="Full name"><Input value={f.name} onChange={set("name")} /></Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Email"><Input value={f.email} onChange={set("email")} /></Field>
            <Field label="Telephone"><Input value={f.phone} onChange={set("phone")} /></Field>
            <Field label="Date of birth"><Input type="date" value={f.dob} onChange={set("dob")} /></Field>
            <Field label="Date joined"><Input type="date" value={f.joined} onChange={set("joined")} disabled={!isEBOD} /></Field>
          </div>
          <Field label="Address"><Input value={f.address} onChange={set("address")} /></Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Occupation"><Input value={f.occupation} onChange={set("occupation")} /></Field>
            <Field label="Employer"><Input value={f.employer} onChange={set("employer")} /></Field>
          </div>
          <Field label="Emergency contact"><Input value={f.emergency} onChange={set("emergency")} placeholder="Name · relationship · phone" /></Field>
          <Field label="Rotary ID (if applicable)"><Input value={f.rotaryId} onChange={set("rotaryId")} /></Field>
          <Field label="Areas of interest"><Input value={f.interests} onChange={set("interests")} placeholder="e.g. Environment, youth mentoring" /></Field>
          <Field label="Skills"><Input value={f.skills} onChange={set("skills")} placeholder="e.g. Graphic design, first aid" /></Field>
          <Btn onClick={saveProfile}>Save profile</Btn>
        </>)}
      </Sheet>
      {statusOpen && <StatusSheet current={m.status} onSave={changeStatus} onClose={() => setStatusOpen(false)} />}
    </FullScreen>
  );
}

function StatusSheet({ current, onSave, onClose }) {
  const [status, setStatus] = useState(current);
  const [note, setNote] = useState("");
  const presets = { "On Leave": "Leave of absence", Transferred: "Transferred to another club", Resigned: "Resignation recorded" };
  return (
    <Sheet open onClose={onClose} title="Change membership status">
      <Field label="New status"><Select value={status} onChange={(e) => { setStatus(e.target.value); if (presets[e.target.value] && !note) setNote(presets[e.target.value]); }}>{MEMBER_STATUSES.map((s) => <option key={s}>{s}</option>)}</Select></Field>
      <Field label="Note" hint="e.g. leave dates, destination club, or reason — kept in the member's history."><Input value={note} onChange={(e) => setNote(e.target.value)} /></Field>
      <Btn onClick={() => onSave(status, note.trim())}>Save status change</Btn>
    </Sheet>
  );
}

/* ============ Library ============ */
function LibraryScreen({ onClose }) {
  const { db, patch, isEBOD, me, notify, audit, showToast } = useApp();
  const [adding, setAdding] = useState(false);
  const [cat, setCat] = useState("all");
  const [f, setF] = useState({ category: LIB_CATS[0], title: "", desc: "", url: "", content: "", files: [] });
  const list = db.library.filter((l) => cat === "all" || l.category === cat);
  const add = () => {
    if (!f.title.trim()) return;
    patch((d) => {
      d.library.unshift({ id: uid(), ...f, title: f.title.trim() });
      notify(d, { type: "library", title: "New library document", body: `"${f.title.trim()}" was added under ${f.category}.` });
      audit(d, "Library document added", f.title.trim());
    });
    setAdding(false); setF({ category: LIB_CATS[0], title: "", desc: "", url: "", content: "", files: [] }); showToast("Added to the library");
  };
  return (
    <FullScreen title="Resource library" onClose={onClose}
      right={isEBOD ? <button onClick={() => setAdding(true)} className="rounded-full p-1.5" style={{ background: "rgba(255,255,255,.18)" }} aria-label="Add"><Icon name="plus" size={18} color="#fff" /></button> : null}>
      <Chips options={[{ id: "all", label: "All" }, ...LIB_CATS.map((c) => ({ id: c, label: c }))]} value={cat} onChange={setCat} />
      <div className="flex flex-col gap-2 mt-3">
        {list.length ? list.map((l) => (
          <Card key={l.id} className="p-4">
            <div className="flex items-start justify-between gap-2">
              <div className="font-extrabold" style={{ fontFamily: DISPLAY, fontSize: 15 }}>{l.title}</div>
              {isEBOD && <button onClick={() => patch((d) => { d.library = d.library.filter((x) => x.id !== l.id); audit(d, "Library document removed", l.title); })} aria-label="Remove"><Icon name="trash" size={15} color="#C9B8C1" /></button>}
            </div>
            <div style={{ fontSize: 11.5, color: CRAN }} className="font-bold uppercase">{l.category}</div>
            <p className="mt-1" style={{ fontSize: 13, color: "#4A3A44" }}>{l.desc}</p>
            {l.content && <details className="mt-1.5"><summary className="font-bold cursor-pointer" style={{ fontSize: 12.5, color: AZURE }}>Read contents</summary><p className="mt-1" style={{ fontSize: 13, whiteSpace: "pre-wrap" }}>{l.content}</p></details>}
            {l.url && <a href={l.url} target="_blank" rel="noreferrer" className="inline-flex items-center gap-1 mt-1.5 font-semibold" style={{ fontSize: 13, color: AZURE }}><Icon name="link" size={13} /> Open link</a>}
            {(l.files || []).length > 0 && <div className="mt-2">{l.files.map((a) => <FileChip key={a.id} f={a} />)}</div>}
          </Card>
        )) : <Empty icon="book" title="Nothing in this category" />}
      </div>
      <Sheet open={adding} onClose={() => setAdding(false)} title="Add a resource" tall>
        <Field label="Category"><Select value={f.category} onChange={(e) => setF({ ...f, category: e.target.value })}>{LIB_CATS.map((c) => <option key={c}>{c}</option>)}</Select></Field>
        <Field label="Title"><Input value={f.title} onChange={(e) => setF({ ...f, title: e.target.value })} /></Field>
        <Field label="Description"><Input value={f.desc} onChange={(e) => setF({ ...f, desc: e.target.value })} /></Field>
        <Field label="Link (optional)"><Input value={f.url} onChange={(e) => setF({ ...f, url: e.target.value })} placeholder="https://…" /></Field>
        <Field label="Text content (optional)"><TextArea value={f.content} onChange={(e) => setF({ ...f, content: e.target.value })} /></Field>
        <FilePick label="Attach files" accept="*/*" onFile={(file) => setF({ ...f, files: [...f.files, file] })} />
        {f.files.length > 0 && <div className="mb-3">{f.files.map((a) => <FileChip key={a.id} f={a} onRemove={() => setF({ ...f, files: f.files.filter((x) => x.id !== a.id) })} />)}</div>}
        <Btn onClick={add} disabled={!f.title.trim()}>Add & notify members</Btn>
      </Sheet>
    </FullScreen>
  );
}

/* ============ More tab ============ */
function MoreTab() {
  const { db, me, year, isEBOD, isPres, setOverlay, signOut, patch, showToast, demo } = useApp();
  const Row = ({ icon, label, sub, onClick, color }) => (
    <Card onClick={onClick} className="flex items-center gap-3 p-3.5 mb-2">
      <div className="rounded-xl p-2" style={{ background: (color || CRAN) + "12", color: color || CRAN }}><Icon name={icon} size={18} /></div>
      <div className="flex-1"><div className="font-bold" style={{ fontSize: 14.5 }}>{label}</div>{sub && <div style={{ fontSize: 12, color: "#8A7580" }}>{sub}</div>}</div>
      <Icon name="chevR" size={18} color="#C9B8C1" />
    </Card>
  );
  const resetDemo = async () => {
    const d = seedDb(); ensureObligations(d);
    await saveShared("doc", d);
    window.location.reload();
  };
  return (
    <div>
      <h2 className="font-black mt-5 mb-3" style={{ fontFamily: DISPLAY, fontSize: 24 }}>More</h2>
      <Card onClick={() => setOverlay({ type: "memberProfile", id: me.id })} className="flex items-center gap-3 p-4 mb-4">
        <Avatar m={me} size={52} />
        <div className="flex-1"><div className="font-extrabold" style={{ fontFamily: DISPLAY, fontSize: 17 }}>{me.name}</div><div style={{ fontSize: 13, color: ROLE_COLOR[me.role] }} className="font-semibold">{me.role} · {year.label}</div></div>
        <Icon name="chevR" size={18} color="#C9B8C1" />
      </Card>
      <Row icon="users" label="Members" sub="Directory, profiles, and membership actions" onClick={() => setOverlay({ type: "members" })} />
      <Row icon="book" label="Resource library" sub="Constitution, guidelines, branding, assets" color={AZURE} onClick={() => setOverlay({ type: "library" })} />
      <Row icon="wallet" label="My dues account" sub="Balance, receipts, and payment history" color={OK} onClick={() => setOverlay({ type: "myAccount", id: me.id })} />
      <Row icon="bell" label="Notification preferences" sub="Choose what reaches you" color="#7A4BA6" onClick={() => setOverlay({ type: "notifPrefs" })} />
      {isEBOD && <Row icon="flag" label="Rotary year manager" sub={`Active: ${year.label} — rollover, archives, year-end report`} color="#B07400" onClick={() => setOverlay({ type: "yearManager" })} />}
      {isPres && <Row icon="edit" label="Club settings" sub="Profile, logo, colours, committees" color={CRAN} onClick={() => setOverlay({ type: "clubSettings" })} />}
      <div className="mt-5 flex flex-col gap-2">
        <Btn kind="quiet" onClick={signOut}><span className="flex items-center justify-center gap-2"><Icon name="logout" size={16} /> {demo ? "Switch member" : "Sign out"}</span></Btn>
        {demo && <Btn kind="danger" onClick={resetDemo}>Reset demo data</Btn>}
      </div>
      <p className="text-center mt-5" style={{ fontSize: 11.5, color: "#B8A8B0" }}>Fellowship through service · {db.club.name}</p>
    </div>
  );
}

/* ============ Club settings ============ */
function ClubSettings({ onClose }) {
  const { db, patch, audit, showToast, year, joinCode } = useApp();
  const [f, setF] = useState({ name: db.club.name, tagline: db.club.tagline, primary: db.club.colors.primary, secondary: db.club.colors.secondary, logo: db.club.logo });
  const [cm, setCm] = useState("");
  const save = () => {
    patch((d) => {
      d.club.name = f.name.trim() || d.club.name;
      d.club.tagline = f.tagline;
      d.club.colors = { primary: f.primary, secondary: f.secondary };
      d.club.logo = f.logo;
      audit(d, "Club settings updated", "");
    });
    showToast("Club settings saved");
  };
  const addCommittee = () => {
    if (!cm.trim()) return;
    patch((d) => { const y = d.years.find((x) => x.id === d.activeYearId); y.committees.push({ id: uid(), name: cm.trim(), chair: "", members: [] }); audit(d, "Committee created", cm.trim()); });
    setCm("");
  };
  return (
    <FullScreen title="Club settings" onClose={onClose} accent={INK}>
      <SectionTitle>Club profile</SectionTitle>
      <Card className="p-4">
        <Field label="Club name"><Input value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} /></Field>
        <Field label="Tagline"><Input value={f.tagline} onChange={(e) => setF({ ...f, tagline: e.target.value })} /></Field>
        <FilePick label="Club logo" accept="image/*" onFile={(file) => setF({ ...f, logo: file.dataUrl })} />
        {f.logo && <div className="mb-3 flex items-center gap-2"><img src={f.logo} alt="" className="rounded-full object-cover" style={{ width: 52, height: 52 }} /><Btn small kind="quiet" onClick={() => setF({ ...f, logo: null })}>Remove</Btn></div>}
        <div className="grid grid-cols-2 gap-3">
          <Field label="Primary colour"><input type="color" value={f.primary} onChange={(e) => setF({ ...f, primary: e.target.value })} style={{ width: "100%", height: 44, border: `1.5px solid ${LINE}`, borderRadius: 12, background: "#fff" }} /></Field>
          <Field label="Secondary colour"><input type="color" value={f.secondary} onChange={(e) => setF({ ...f, secondary: e.target.value })} style={{ width: "100%", height: 44, border: `1.5px solid ${LINE}`, borderRadius: 12, background: "#fff" }} /></Field>
        </div>
        <Btn onClick={save}>Save settings</Btn>
        <p className="mt-2" style={{ fontSize: 12, color: "#9A8B93" }}>Cranberry #D41367 is the official Rotaract colour — but the app follows whatever you set.</p>
      </Card>
      {joinCode && (
        <Card className="p-4 mt-2.5" style={{ borderColor: AZURE + "55" }}>
          <div className="font-bold" style={{ fontSize: 13, color: AZURE }}>Member invite code</div>
          <div className="font-black" style={{ fontFamily: DISPLAY, fontSize: 26, letterSpacing: ".14em" }}>{joinCode}</div>
          <p style={{ fontSize: 12.5, color: "#8A7580" }}>Share this code. New members create an account, enter the code, and appear under Members as "Applied" for the board to approve.</p>
        </Card>
      )}
      <SectionTitle>Committees · {year.label}</SectionTitle>
      <Card className="p-4">
        {(year.committees || []).map((c) => (
          <div key={c.id} className="flex items-center gap-2 mb-2" style={{ fontSize: 13.5 }}>
            <Icon name="users" size={15} color={AZURE} /><span className="flex-1"><b>{c.name}</b> · {c.members.length} member{c.members.length === 1 ? "" : "s"}</span>
            <button onClick={() => patch((d) => { const y = d.years.find((x) => x.id === d.activeYearId); y.committees = y.committees.filter((x) => x.id !== c.id); })}><Icon name="trash" size={14} color="#C9B8C1" /></button>
          </div>
        ))}
        <div className="flex gap-2"><Input value={cm} onChange={(e) => setCm(e.target.value)} placeholder="New committee name" /><Btn small onClick={addCommittee}>Add</Btn></div>
        <p className="mt-2" style={{ fontSize: 12, color: "#9A8B93" }}>Assign members to committees from their profile page.</p>
      </Card>
    </FullScreen>
  );
}

/* ============ Rotary Year manager ============ */
function yearEndReportHtml(db, y) {
  const cur = db.duesConfig.currency;
  const totals = clubTotals(db, y.id);
  const officers = EBOD.map((r) => `<tr><th>${r}</th><td>${db.members.find((m) => m.role === r && ACTIVE_LIKE.includes(m.status))?.name || "—"}</td></tr>`).join("");
  const memberRows = db.members.filter((m) => ACTIVE_LIKE.includes(m.status)).map((m) => {
    const att = attendanceStats(db, m.id, y.id); const a = memberAccount(db, m.id, y.id);
    return `<tr><td>${m.name}</td><td>${m.role}</td><td>${att.pct === null ? "—" : att.pct + "%"}</td><td>${serviceHoursOf(db, m.id)}</td><td>${a.balance <= 0 ? "In good standing" : money(a.balance, cur) + " owing"}</td></tr>`;
  }).join("");
  const projRows = db.projects.filter((p) => p.yearId === y.id && p.status !== "Draft").map((p) => `<tr><td>${p.title}</td><td>${p.area}</td><td>${p.status}</td><td>${p.serviceHours.reduce((s, h) => s + Number(h.hours || 0), 0)} h</td></tr>`).join("");
  const meetings = db.meetings.filter((m) => m.yearId === y.id && m.status === "published").length;
  return docShell(`Year-end report ${y.label}`, `
    <h1>${db.club.name}</h1><div class="muted">Year-end report · Rotary year ${y.label} (${y.start} – ${y.end})</div>
    <h2>Officers</h2><table>${officers}</table>
    <h2>Club at a glance</h2>
    <table><tr><th>Active members</th><td>${db.members.filter((m) => ACTIVE_LIKE.includes(m.status)).length}</td></tr>
    <tr><th>Meetings held</th><td>${meetings}</td></tr>
    <tr><th>Projects (non-draft)</th><td>${db.projects.filter((p) => p.yearId === y.id && p.status !== "Draft").length}</td></tr>
    <tr><th>Total service hours</th><td>${db.projects.reduce((s, p) => s + p.serviceHours.reduce((a, h) => a + Number(h.hours || 0), 0), 0)} h</td></tr></table>
    <h2>Finances</h2>
    <table><tr><th>Income</th><td>${money(totals.income, cur)}</td></tr><tr><th>Expenses</th><td>${money(totals.expense, cur)}</td></tr><tr><th>Closing balance</th><td class="big">${money(totals.balance, cur)}</td></tr></table>
    <h2>Members</h2><table><tr><th>Name</th><th>Role</th><th>Attendance</th><th>Service hrs</th><th>Dues standing</th></tr>${memberRows}</table>
    <h2>Projects</h2><table><tr><th>Project</th><th>Area of focus</th><th>Status</th><th>Hours</th></tr>${projRows || "<tr><td colspan=4>None</td></tr>"}</table>`, CRAN);
}

function YearManager({ onClose }) {
  const { db, patch, me, year, isPres, notify, audit, showToast, memberById } = useApp();
  const [wizard, setWizard] = useState(false);
  const [step, setStep] = useState(0);
  const activeIds = db.members.filter((m) => ACTIVE_LIKE.includes(m.status)).map((m) => m.id);
  const [carry, setCarry] = useState(activeIds);
  const [board, setBoard] = useState(() => { const b = {}; EBOD.forEach((r) => { b[r] = db.members.find((m) => m.role === r && ACTIVE_LIKE.includes(m.status))?.id || ""; }); return b; });
  const [opts, setOpts] = useState({ balances: true, projects: true, library: true });
  const nextStartYear = year.startYear + 1;

  const rollover = () => {
    patch((d) => {
      const oldY = d.years.find((x) => x.id === d.activeYearId);
      const totals = clubTotals(d, oldY.id);
      const newY = { id: ryId(nextStartYear), startYear: nextStartYear, label: ryLabel(nextStartYear), ...ryBounds(nextStartYear), active: true, archivedAt: null, committees: clone(oldY.committees || []) };
      if (d.years.some((x) => x.id === newY.id)) { showToast("That Rotary year already exists."); return false; }
      /* archive old year */
      oldY.active = false; oldY.archivedAt = Date.now();
      /* members: record positions, set carried members, resolve statuses */
      d.members.forEach((m) => {
        if (ACTIVE_LIKE.includes(m.status)) {
          m.positions.unshift({ yearId: oldY.id, role: m.role });
          if (!carry.includes(m.id)) {
            m.statusHistory.push({ id: uid(), from: m.status, to: "Alumni", note: `Not continued into ${newY.label}`, at: Date.now(), by: me.name });
            m.status = "Alumni";
          }
        }
      });
      /* new board */
      const boardIds = Object.values(board).filter(Boolean);
      d.members.forEach((m) => {
        const role = Object.keys(board).find((r) => board[r] === m.id);
        if (role) m.role = role;
        else if (carry.includes(m.id) && EBOD.includes(m.role)) m.role = "Member";
      });
      /* carry forward member balances */
      if (opts.balances) {
        carry.forEach((mid) => {
          const a = memberAccount(d, mid, oldY.id);
          if (a.balance > 0.004) {
            d.charges.push({ id: uid(), yearId: newY.id, memberId: mid, kind: "carryforward", label: `Balance carried forward from ${oldY.label}`, period: ryMonths(nextStartYear)[0], amount: Math.round(a.balance * 100) / 100, dueDate: newY.start, at: Date.now(), by: me.id, reversed: false });
          }
        });
        /* club opening balance */
        if (Math.abs(totals.balance) > 0.004) {
          d.transactions.push({ id: uid(), yearId: newY.id, type: totals.balance >= 0 ? "income" : "expense", category: "Other", desc: `Opening balance carried from ${oldY.label}`, payee: "", amount: Math.abs(Math.round(totals.balance * 100) / 100), date: newY.start, by: me.id, approval: "approved", projectId: "", receipt: null, reversed: false, reversalOf: null });
        }
      }
      /* copy pending projects into new year */
      if (opts.projects) {
        d.projects.forEach((p) => {
          if (p.yearId === oldY.id && !["Completed", "Denied"].includes(p.status)) p.yearId = newY.id;
        });
      }
      /* library persists at club level (option kept for clarity) */
      d.years.unshift(newY);
      d.activeYearId = newY.id;
      ensureObligations(d);
      notify(d, { type: "announcements", title: `Welcome to Rotary year ${newY.label}! 🎉`, body: `${oldY.label} is archived. New board: ${EBOD.map((r) => `${r} — ${d.members.find((m) => m.id === board[r])?.name || "TBA"}`).join("; ")}.` });
      audit(d, "Rotary year rollover", `${oldY.label} → ${newY.label}; ${carry.length} members carried${opts.balances ? "; balances carried" : ""}${opts.projects ? "; pending projects carried" : ""}`);
    });
    setWizard(false); showToast(`Rotary year ${ryLabel(nextStartYear)} is now active`);
  };

  return (
    <FullScreen title="Rotary year manager" onClose={onClose} accent="#B07400">
      <Card className="p-4" style={{ borderColor: GOLD }}>
        <div className="font-bold uppercase" style={{ fontSize: 10.5, letterSpacing: ".14em", color: "#8A5A00", fontFamily: DISPLAY }}>Active year</div>
        <div className="font-black" style={{ fontFamily: DISPLAY, fontSize: 24 }}>{year.label}</div>
        <div style={{ fontSize: 12.5, color: "#8A7580" }}>{fmtDate(year.start)} → {fmtDate(year.end)} · only one year is active at a time</div>
        <div className="flex gap-1.5 mt-3 flex-wrap">
          <Btn small kind="quiet" onClick={() => printHtml(`Year-end-${year.id}`, yearEndReportHtml(db, year))}>Year-end report</Btn>
          {isPres && <Btn small kind="dark" onClick={() => { setStep(0); setWizard(true); }}>Start {ryLabel(nextStartYear)} →</Btn>}
        </div>
      </Card>
      <SectionTitle>Archived years</SectionTitle>
      {db.years.filter((y) => !y.active).map((y) => {
        const t = clubTotals(db, y.id);
        return (
          <Card key={y.id} className="p-4 mb-2">
            <div className="flex items-center justify-between">
              <div className="font-extrabold" style={{ fontFamily: DISPLAY, fontSize: 16 }}>{y.label}</div>
              <Badge>Archived</Badge>
            </div>
            <div style={{ fontSize: 12.5, color: "#8A7580" }} className="mt-0.5">
              {db.meetings.filter((m) => m.yearId === y.id).length} meetings · {db.projects.filter((p) => p.yearId === y.id).length} projects · net {money(t.balance, db.duesConfig.currency)}
            </div>
            <div className="mt-2"><Btn small kind="quiet" onClick={() => printHtml(`Year-end-${y.id}`, yearEndReportHtml(db, y))}>Download year-end report</Btn></div>
          </Card>
        );
      })}
      <Sheet open={wizard} onClose={() => setWizard(false)} title={`Begin Rotary year ${ryLabel(nextStartYear)}`} tall>
        <div className="flex gap-1.5 mb-4">{["Members", "New board", "Carry over"].map((s, i) => (
          <div key={s} className="flex-1 rounded-full text-center font-bold py-1.5" style={{ fontFamily: DISPLAY, fontSize: 12, background: step === i ? "#B07400" : step > i ? "#B0740022" : "#fff", color: step === i ? "#fff" : "#8A5A00", border: `1px solid ${step >= i ? "transparent" : LINE}` }}>{i + 1}. {s}</div>
        ))}</div>
        {step === 0 && (<>
          <p style={{ fontSize: 13, color: "#6B5A64" }} className="mb-3">Choose who continues as an active member. Unchecked members are recorded as Alumni.</p>
          {db.members.filter((m) => ACTIVE_LIKE.includes(m.status)).map((m) => (
            <button key={m.id} onClick={() => setCarry(carry.includes(m.id) ? carry.filter((x) => x !== m.id) : [...carry, m.id])} className="w-full flex items-center gap-3 mb-1.5 p-2.5 rounded-xl text-left" style={{ background: "#fff", border: `1.5px solid ${carry.includes(m.id) ? OK : LINE}` }}>
              <div className="rounded-full flex items-center justify-center" style={{ width: 22, height: 22, border: `2px solid ${carry.includes(m.id) ? OK : "#C9B8C1"}`, background: carry.includes(m.id) ? OK : "transparent" }}>{carry.includes(m.id) && <Icon name="check" size={12} color="#fff" />}</div>
              <Avatar m={m} size={30} /><span className="font-semibold" style={{ fontSize: 13.5 }}>{m.name}</span>
            </button>
          ))}
          <Btn onClick={() => setStep(1)}>Next — assign the board</Btn>
        </>)}
        {step === 1 && (<>
          <p style={{ fontSize: 13, color: "#6B5A64" }} className="mb-3">Assign the incoming executive board. Outgoing officers keep their old role in their positions history.</p>
          {EBOD.map((r) => (
            <Field key={r} label={r}>
              <Select value={board[r]} onChange={(e) => setBoard({ ...board, [r]: e.target.value })}>
                <option value="">— unassigned —</option>
                {db.members.filter((m) => carry.includes(m.id)).map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}
              </Select>
            </Field>
          ))}
          <div className="flex gap-2"><Btn kind="quiet" onClick={() => setStep(0)} style={{ flex: 1 }}>Back</Btn><Btn onClick={() => setStep(2)} style={{ flex: 2 }}>Next</Btn></div>
        </>)}
        {step === 2 && (<>
          {[["balances", "Carry forward balances", "Outstanding member dues become opening charges; the club's closing balance opens the new ledger."],
            ["projects", "Copy pending projects", "Approved and in-review projects continue into the new year."],
            ["library", "Keep library documents", "The resource library carries over unchanged."]].map(([k, label, sub]) => (
            <button key={k} onClick={() => setOpts({ ...opts, [k]: !opts[k] })} className="w-full flex items-start gap-3 mb-2 p-3 rounded-xl text-left" style={{ background: "#fff", border: `1.5px solid ${opts[k] ? OK : LINE}` }}>
              <div className="rounded-full flex items-center justify-center flex-shrink-0 mt-0.5" style={{ width: 22, height: 22, border: `2px solid ${opts[k] ? OK : "#C9B8C1"}`, background: opts[k] ? OK : "transparent" }}>{opts[k] && <Icon name="check" size={12} color="#fff" />}</div>
              <div><div className="font-bold" style={{ fontSize: 14 }}>{label}</div><div style={{ fontSize: 12, color: "#8A7580" }}>{sub}</div></div>
            </button>
          ))}
          <Card className="p-3 mb-3" style={{ background: GOLD + "10", borderColor: GOLD }}>
            <div style={{ fontSize: 12.5, color: "#7A5A00" }}>Before finishing, download the <b>year-end report</b> for {year.label} above. Rolling over archives {year.label} and activates {ryLabel(nextStartYear)}.</div>
          </Card>
          <div className="flex gap-2"><Btn kind="quiet" onClick={() => setStep(1)} style={{ flex: 1 }}>Back</Btn><Btn kind="dark" onClick={rollover} style={{ flex: 2 }}>Archive {year.label} & begin {ryLabel(nextStartYear)}</Btn></div>
        </>)}
      </Sheet>
    </FullScreen>
  );
}

