import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── COLORS ───────────────────────────────────────────────────────────────────
const C = {
  bg: "#0A0F1E", card: "#111827", card2: "#161f30", border: "#1F2937",
  accent: "#F97316", blue: "#3B82F6", green: "#22C55E",
  yellow: "#EAB308", red: "#EF4444", purple: "#A855F7",
  muted: "#6B7280", text: "#F9FAFB", sub: "#9CA3AF", teal: "#14B8A6",
};
const APP_VERSION = "18.9";

// ─── SUPABASE STORAGE HELPERS ────────────────────────────────────────────────
// Calls server-side API routes which talk to Supabase.
// Data is shared across all devices and users.
function tableFor(key) {
  if (key === "concrete-data") return "concrete_data";
  if (key === "tm-data") return "tm_data";
  return "certs_data";
}
async function storageGet(key) {
  try {
    const res = await fetch(`/api/data-get?table=${tableFor(key)}`);
    if (!res.ok) return null;
    return await res.json();
  } catch { return null; }
}
async function storageSet(key, value) {
  // A brief API/database hiccup should not lose a user's change. Retry the
  // same payload before reporting a failure.
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const res = await fetch("/api/data-set", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ table: tableFor(key), data: value }),
      });
      if (res.ok) return true;
      console.error(`storageSet failed (${res.status})`, await res.text());
    } catch (error) {
      console.error("storageSet network error", error);
    }
    if (attempt < 2) await new Promise(resolve=>setTimeout(resolve, 500*(attempt+1)));
  }
  return false;
}
async function storageDel(key) {
  const empty = key === "concrete-data"
    ? { tickets: [], invoices: [] }
    : key === "tm-data"
      ? { items: [], labourEntries: [], submittedWeeks: [] }
      : { certs: [] };
  return storageSet(key, empty);
}


// ─── CERT TYPES ───────────────────────────────────────────────────────────────
const CERT_TYPES = [
  "WHMIS","Fall Protection","First Aid / CPR","Confined Space Entry",
  "Aerial Lift (Scissor Lift / Boom Lift)","Telehandler / Forklift",
  "Rigging & Hoisting","Crane Signalling","Respirator Fit Testing",
  "Silica Awareness","Fire Extinguisher Training","Lockout / Tagout Awareness",
  "Traffic Control / Flagging","Power Tool Safety","Ladder Safety",
  "Scaffold Safety","Occupational Health & Safety Orientation",
  "Site-Specific Safety Orientation","Working at Heights","Propane Handling",
  "Asbestos Awareness","TDG (Transportation of Dangerous Goods)",
  "Elevated Work Platform Certification","Skid Steer / Bobcat Training",
  "Excavation / Trenching Safety","Hot Work / Fire Watch Training",
];

// ─── TRADE / SAFETY PROGRAM DOCUMENT TYPES ────────────────────────────────────
const TRADE_DOC_TYPES = [
  "WCB Clearance Letter",
  "Safety Data Sheet (SDS)",
  "COR Certification (or Letter Confirming Process Started)",
  "CGL Insurance Certificate",
];

const CGL_MIN_LIMIT = 5000000;
const CGL_CERT_HOLDER_LABEL = "Summer Wind Holdings Limited";
const CGL_ADDITIONAL_INSURED_LABELS = ["Summer Wind Holdings Limited", "Southwest Construction Management"];

function normEntity(s) {
  return (s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}
function isSummerWindHolding(s) {
  return normEntity(s).includes("summerwindholding");
}
function isSouthwestConstruction(s) {
  return normEntity(s).includes("southwestconstruction");
}
function parseLimit(v) {
  if (v == null) return null;
  const n = parseFloat(String(v).replace(/[^0-9.]/g, ""));
  return isNaN(n) ? null : n;
}
// Returns an array of human-readable compliance issues for a CGL doc. Empty array = compliant.
function cglComplianceFlags(doc) {
  if (doc.doc_type !== "CGL Insurance Certificate") return [];
  const flags = [];
  const limit = parseLimit(doc.per_occurrence_limit);
  if (limit == null || limit < CGL_MIN_LIMIT) {
    flags.push(`Per-occurrence limit below $5,000,000 (found: ${doc.per_occurrence_limit || "not found"})`);
  }
  if (!isSummerWindHolding(doc.certificate_holder)) {
    flags.push(`Certificate holder is not "${CGL_CERT_HOLDER_LABEL}" (found: ${doc.certificate_holder || "not found"})`);
  }
  const insuredList = Array.isArray(doc.additional_insured) ? doc.additional_insured : (doc.additional_insured ? [doc.additional_insured] : []);
  const insuredText = insuredList.join(" | ");
  if (!isSummerWindHolding(insuredText)) {
    flags.push(`"${CGL_CERT_HOLDER_LABEL}" not listed as additional insured`);
  }
  if (!isSouthwestConstruction(insuredText)) {
    flags.push(`"Southwest Construction Management" not listed as additional insured`);
  }
  return flags;
}

// ─── CONCRETE SCOPE ───────────────────────────────────────────────────────────
const SCOPE = [
  { area:"Mud Slabs",   item:"Slabs",                m3:185.0,  mpa:"20 MPa"       },
  { area:"Crane Base",  item:"Interior foundations", m3:137.7,  mpa:"35 MPa"       },
  { area:"SOG",         item:"Slabs",                m3:305.0,  mpa:"25 MPa/N-CF" },
  { area:"Foundations", item:"Wall",                 m3:287.3,  mpa:"25 MPa/F-2"  },
  { area:"Foundations", item:"Raft",                 m3:327.6,  mpa:"25 MPa/F-2"  },
  { area:"Foundations", item:"Strip foundations",    m3:42.1,   mpa:"25 MPa/F-2"  },
  { area:"Foundations", item:"Columns",              m3:36.1,   mpa:"25 MPa/F-2"  },
  { area:"Foundations", item:"Interior foundations", m3:128.0,  mpa:"25 MPa/F-2"  },
  { area:"Foundations", item:"Slabs",                m3:276.9,  mpa:"25 MPa/N-CF" },
  { area:"P2",          item:"Wall",                 m3:239.6,  mpa:"25 MPa/N-CF" },
  { area:"P2",          item:"Columns",              m3:29.9,   mpa:"35 MPa/N-CF" },
  { area:"P2",          item:"Slabs",                m3:495.5,  mpa:"25 MPa/N-CF" },
  { area:"P1",          item:"Wall",                 m3:290.2,  mpa:"35 MPa/N-CF" },
  { area:"P1",          item:"Columns",              m3:38.9,   mpa:"35 MPa/N-CF" },
  { area:"P1",          item:"Slabs",                m3:534.6,  mpa:"35 MPa/N-CF" },
  { area:"P1",          item:"Curbs",                m3:12.1,   mpa:"35 MPa/N-CF"  },
  { area:"Level 1",     item:"Wall",                 m3:51.5,   mpa:"35 MPa/N-CF" },
  { area:"Level 1",     item:"Columns",              m3:47.5,   mpa:"35 MPa/N-CF" },
  { area:"Level 1",     item:"Slabs",                m3:590.7,  mpa:"35 MPa/N-CF" },
  { area:"Level 1",     item:"Curbs",                m3:2.1,    mpa:"35 MPa/N-CF"  },
  { area:"2nd",         item:"Wall",                 m3:43.2,   mpa:"35 MPa/N-CF" },
  { area:"2nd",         item:"Columns",              m3:36.9,   mpa:"35 MPa/N-CF" },
  { area:"2nd",         item:"Slabs",                m3:390.7,  mpa:"35 MPa/N-CF" },
  { area:"3rd",         item:"Wall",                 m3:48.7,   mpa:"35 MPa/N-CF" },
  { area:"3rd",         item:"Columns",              m3:41.9,   mpa:"35 MPa/N-CF" },
  { area:"3rd",         item:"Slabs",                m3:387.6,  mpa:"35 MPa/N-CF" },
  { area:"4th",         item:"Wall",                 m3:47.7,   mpa:"35 MPa/N-CF" },
  { area:"4th",         item:"Columns",              m3:20.1,   mpa:"35 MPa/N-CF" },
  { area:"4th",         item:"Slabs",                m3:458.9,  mpa:"35 MPa/N-CF" },
  { area:"4th",         item:"Curbs",                m3:28.2,   mpa:"35 MPa/N-CF"  },
  { area:"5th",         item:"Wall",                 m3:43.1,   mpa:"35 MPa/N-CF" },
  { area:"5th",         item:"Columns",              m3:17.6,   mpa:"35 MPa/N-CF" },
  { area:"5th",         item:"Slabs",                m3:236.5,  mpa:"35 MPa/N-CF" },
  { area:"5th",         item:"Curbs",                m3:1.0,    mpa:"35 MPa/N-CF"  },
  { area:"6th",         item:"Wall",                 m3:43.1,   mpa:"35 MPa/N-CF" },
  { area:"6th",         item:"Columns",              m3:17.6,   mpa:"35 MPa/N-CF" },
  { area:"6th",         item:"Slabs",                m3:236.5,  mpa:"35 MPa/N-CF" },
  { area:"6th",         item:"Curbs",                m3:1.0,    mpa:"35 MPa/N-CF"  },
  { area:"7th",         item:"Wall",                 m3:43.1,   mpa:"35 MPa/N-CF" },
  { area:"7th",         item:"Columns",              m3:17.6,   mpa:"35 MPa/N-CF" },
  { area:"7th",         item:"Slabs",                m3:236.5,  mpa:"35 MPa/N-CF" },
  { area:"7th",         item:"Curbs",                m3:1.0,    mpa:"35 MPa/N-CF"  },
  { area:"8th",         item:"Wall",                 m3:43.1,   mpa:"35 MPa/N-CF" },
  { area:"8th",         item:"Columns",              m3:17.6,   mpa:"35 MPa/N-CF" },
  { area:"8th",         item:"Slabs",                m3:236.5,  mpa:"35 MPa/N-CF" },
  { area:"8th",         item:"Curbs",                m3:1.0,    mpa:"35 MPa/N-CF"  },
  { area:"9th",         item:"Wall",                 m3:43.1,   mpa:"35 MPa/N-CF" },
  { area:"9th",         item:"Columns",              m3:17.6,   mpa:"35 MPa/N-CF" },
  { area:"9th",         item:"Slabs",                m3:236.5,  mpa:"35 MPa/N-CF" },
  { area:"9th",         item:"Curbs",                m3:1.0,    mpa:"35 MPa/N-CF"  },
  { area:"10th",        item:"Wall",                 m3:43.1,   mpa:"35 MPa/N-CF" },
  { area:"10th",        item:"Columns",              m3:17.6,   mpa:"35 MPa/N-CF" },
  { area:"10th",        item:"Slabs",                m3:236.5,  mpa:"35 MPa/N-CF" },
  { area:"10th",        item:"Curbs",                m3:1.0,    mpa:"35 MPa/N-CF"  },
  { area:"11th",        item:"Wall",                 m3:43.1,   mpa:"35 MPa/N-CF" },
  { area:"11th",        item:"Columns",              m3:17.6,   mpa:"35 MPa/N-CF" },
  { area:"11th",        item:"Slabs",                m3:236.5,  mpa:"35 MPa/N-CF" },
  { area:"11th",        item:"Curbs",                m3:1.0,    mpa:"35 MPa/N-CF"  },
  { area:"12th",        item:"Wall",                 m3:43.1,   mpa:"35 MPa/N-CF" },
  { area:"12th",        item:"Columns",              m3:17.6,   mpa:"35 MPa/N-CF" },
  { area:"12th",        item:"Slabs",                m3:236.5,  mpa:"35 MPa/N-CF" },
  { area:"12th",        item:"Curbs",                m3:1.0,    mpa:"35 MPa/N-CF"  },
  { area:"14th",        item:"Wall",                 m3:43.1,   mpa:"35 MPa/N-CF" },
  { area:"14th",        item:"Columns",              m3:17.6,   mpa:"35 MPa/N-CF" },
  { area:"14th",        item:"Slabs",                m3:236.5,  mpa:"35 MPa/N-CF" },
  { area:"14th",        item:"Curbs",                m3:1.0,    mpa:"35 MPa/N-CF"  },
  { area:"15th",        item:"Wall",                 m3:43.1,   mpa:"35 MPa/N-CF" },
  { area:"15th",        item:"Columns",              m3:17.6,   mpa:"35 MPa/N-CF" },
  { area:"15th",        item:"Slabs",                m3:236.5,  mpa:"35 MPa/N-CF" },
  { area:"15th",        item:"Curbs",                m3:1.0,    mpa:"35 MPa/N-CF"  },
  { area:"16th",        item:"Wall",                 m3:43.1,   mpa:"35 MPa/N-CF" },
  { area:"16th",        item:"Columns",              m3:17.6,   mpa:"35 MPa/N-CF" },
  { area:"16th",        item:"Slabs",                m3:236.5,  mpa:"35 MPa/N-CF" },
  { area:"16th",        item:"Curbs",                m3:1.0,    mpa:"35 MPa/N-CF"  },
  { area:"17th",        item:"Wall",                 m3:43.1,   mpa:"35 MPa/N-CF" },
  { area:"17th",        item:"Columns",              m3:17.6,   mpa:"35 MPa/N-CF" },
  { area:"17th",        item:"Slabs",                m3:236.5,  mpa:"35 MPa/N-CF" },
  { area:"17th",        item:"Curbs",                m3:1.0,    mpa:"35 MPa/N-CF"  },
  { area:"Penthouse",   item:"Wall",                 m3:56.9,   mpa:"35 MPa/N-CF" },
  { area:"Penthouse",   item:"Columns",              m3:17.0,   mpa:"35 MPa/N-CF" },
  { area:"Penthouse",   item:"Slabs",                m3:228.9,  mpa:"35 MPa/N-CF" },
  { area:"Mech. Roof",  item:"Slabs",                m3:6.4,    mpa:"35 MPa/N-CF" },
];

const TOTAL_SCOPE_M3 = SCOPE.reduce((s,r) => s + r.m3, 0);

// ─── PUMP BUDGET ──────────────────────────────────────────────────────────────
const PUMP_BUDGET = [
  { category: "Mud Slabs",          volume_m3: 185.00,   hours: 6.16   },
  { category: "Foundations",        volume_m3: 785.00,   hours: 39.25  },
  { category: "Slab on Grade",      volume_m3: 305.00,   hours: 19.04  },
  { category: "Sus Slabs up to L6", volume_m3: 166.55,   hours: 33.31  },
  { category: "Verticals up to L6", volume_m3: 1102.00,  hours: 92.0   },
];
const PUMP_CATEGORIES = PUMP_BUDGET.map(r => r.category);
const TOTAL_PUMP_BUDGET_M3 = PUMP_BUDGET.reduce((s,r) => s + r.volume_m3, 0);
const TOTAL_PUMP_BUDGET_HOURS = PUMP_BUDGET.reduce((s,r) => s + r.hours, 0);
const M3_TO_YD3 = 1.30795;
// Files are base64 encoded before being sent through the serverless API, which
// adds roughly 33% to their size. Keep raw files below 3 MB to stay safely
// under the request-body limit.
const MAX_API_FILE_BYTES = 3 * 1024 * 1024;
const AREAS = [...new Set(SCOPE.map(r => r.area))];
const ITEMS = [...new Set(SCOPE.map(r => r.item).filter(Boolean))];
const MPA_SPEC = {};
SCOPE.forEach(r => { if (r.area && r.item) MPA_SPEC[`${r.area}|||${r.item}`] = r.mpa; });

function parseMpaNum(str) {
  if (!str) return null;
  const s = String(str);
  // Match "35 MPa", "35MPa", "35MPA" etc
  const m = s.match(/(\d+)\s*[Mm][Pp][Aa]/);
  if (m) return parseInt(m[1]);
  // Match supplier product codes such as Q35NA1A, Q25NB1A, and Ocean's
  // S20P20. The first number after Q/S is the concrete strength; the later
  // number may be aggregate size and must not be used as the MPa value.
  const productCode = s.match(/[QqSs](\d+)[A-Za-z]/);
  if (productCode) return parseInt(productCode[1]);
  // Match plain number
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function checkMpaMismatch(ticket) {
  if (!ticket.area || !ticket.item || !ticket.mix_design) return null;
  const key = `${ticket.area}|||${ticket.item}`;
  const specStr = MPA_SPEC[key];
  if (!specStr) return null;
  const specNum = parseMpaNum(specStr);
  const ticketNum = parseMpaNum(ticket.mix_design);
  if (!specNum || !ticketNum) return null;
  if (ticketNum !== specNum) return { specMpa: specStr, ticketMpa: ticket.mix_design };
  return null;
}

function validItemsForArea(area) {
  return [...new Set(SCOPE.filter(r=>r.area===area&&r.item).map(r=>r.item))];
}
function ticketCodingIssue(ticket) {
  if (!(parseFloat(ticket.volume_m3)>0)) return null;
  if (!ticket.area) return "area";
  if (!ticket.item) return "element";
  if (!MPA_SPEC[`${ticket.area}|||${ticket.item}`]) return "valid area/element combination";
  if (!parseMpaNum(ticket.mix_design)) return "mix design / MPa";
  return null;
}

// v17 data correction: Ocean's July 27 invoice combined the crane-base pour
// with four 20 MPa mud-slab loads. Earlier versions assigned the whole batch
// to Crane Base. Keep this migration deliberately narrow so user-confirmed
// locations on unrelated tickets are never overwritten.
const V17_MUD_SLAB_TICKETS = new Set(["251787", "251790", "251793", "251795"]);
function migrateConcreteTicketsV17(allTickets) {
  let changed = false;
  const tickets = (allTickets || []).map(ticket => {
    const ticketNo = String(ticket.ticket_number || "").replace(/\D/g, "");
    const isJuly27MudSlab = ticket.date === "2026-07-27" && V17_MUD_SLAB_TICKETS.has(ticketNo);
    if (isJuly27MudSlab && (ticket.area !== "Mud Slabs" || ticket.item !== "Slabs")) {
      changed = true;
      return { ...ticket, area:"Mud Slabs", item:"Slabs", _v17_location_corrected:true };
    }
    // The crane-base pour is tracked as interior foundations. Versions 17.1
    // and 17.2 incorrectly changed these saved tickets to Slabs; reverse that
    // correction without changing ticket quantities or mix data.
    if (ticket.area === "Crane Base" && ticket.item !== "Interior foundations") {
      changed = true;
      return { ...ticket, item:"Interior foundations", _v17_location_corrected:true };
    }
    if (ticket.area === "Mud Slabs" && ticket.item !== "Slabs") {
      changed = true;
      return { ...ticket, item:"Slabs", _v17_location_corrected:true };
    }
    return ticket;
  });
  return { tickets, changed };
}

// Keep OCR date mistakes (for example 2026/07/20 becoming 2020-07-26)
// out of the permanent ticket log. The filename is a useful second source
// because site scans are commonly named for the pour date.
function dateFromFilename(name) {
  const s = String(name || "").replace(/\.[^.]+$/, "");
  const months = {jan:1,january:1,feb:2,february:2,mar:3,march:3,apr:4,april:4,may:5,jun:6,june:6,jul:7,july:7,aug:8,august:8,sep:9,september:9,oct:10,october:10,nov:11,november:11,dec:12,december:12};
  const named = s.match(/\b(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec)\s+(\d{1,2})[ ,_-]+(20\d{2})\b/i);
  if (named) return `${named[3]}-${String(months[named[1].toLowerCase()]).padStart(2,"0")}-${String(+named[2]).padStart(2,"0")}`;
  const numeric = s.match(/\b(20\d{2})[-_. ](\d{1,2})[-_. ](\d{1,2})\b/);
  if (numeric) return `${numeric[1]}-${numeric[2].padStart(2,"0")}-${numeric[3].padStart(2,"0")}`;
  return null;
}
function normalizeTicketDate(value, fileName) {
  const fallback = dateFromFilename(fileName);
  const m = String(value || "").match(/^(\d{4})[-/]([01]?\d)[-/]([0-3]?\d)$/);
  if (!m) return fallback || value || "";
  const normalized = `${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  const year = +m[1], currentYear = new Date().getFullYear();
  return (fallback && Math.abs(year-currentYear) > 1) ? fallback : normalized;
}
function ticketNumberKey(value) {
  return String(value || "").trim().replace(/\s+/g, "").toLowerCase();
}

// ─── EXPIRY HELPERS ────────────────────────────────────────────────────────────
function daysUntilExpiry(dateStr) {
  if (!dateStr) return null;
  const exp = new Date(dateStr);
  const now = new Date();
  exp.setHours(0,0,0,0); now.setHours(0,0,0,0);
  return Math.round((exp - now) / 86400000);
}
function expiryStatus(dateStr) {
  const d = daysUntilExpiry(dateStr);
  if (d === null) return { label:"No Expiry", color: C.muted, level: "none" };
  if (d < 0)      return { label:`Expired ${Math.abs(d)}d ago`, color: C.red,    level: "expired"  };
  if (d <= 30)    return { label:`Expires in ${d}d`,            color: C.red,    level: "critical" };
  if (d <= 60)    return { label:`Expires in ${d}d`,            color: C.yellow, level: "warning"  };
  return           { label:`Expires in ${d}d`,                  color: C.green,  level: "ok"       };
}

// ─── SHARED UI ────────────────────────────────────────────────────────────────
function Badge({ color, children }) {
  return <span style={{ background:color+"20", color, border:`1px solid ${color}44`, borderRadius:6, padding:"2px 9px", fontSize:11, fontWeight:700 }}>{children}</span>;
}
function Bar({ pct, color=C.accent }) {
  return <div style={{ background:"#1F2937", borderRadius:99, height:7, overflow:"hidden" }}>
    <div style={{ height:"100%", width:`${Math.min(100,Math.max(0,pct||0))}%`, background:color, borderRadius:99, transition:"width .5s ease" }} />
  </div>;
}
function Stat({ label, value, sub, color=C.accent }) {
  return <div style={{ background:C.card, border:`1px solid ${C.border}`, borderRadius:14, padding:"18px 22px", flex:1, minWidth:130 }}>
    <div style={{ color:C.muted, fontSize:10, fontWeight:700, letterSpacing:1, textTransform:"uppercase", marginBottom:6 }}>{label}</div>
    <div style={{ color, fontSize:26, fontWeight:800, fontFamily:"monospace" }}>{value}</div>
    {sub && <div style={{ color:C.muted, fontSize:12, marginTop:3 }}>{sub}</div>}
  </div>;
}
function MpaBadge({ mpa }) {
  if (!mpa) return null;
  const num = parseMpaNum(mpa);
  const color = num >= 35 ? C.purple : num >= 32 ? C.blue : C.accent;
  return <Badge color={color}>{mpa}</Badge>;
}
function fmt(n, d=2) {
  if (n==null||isNaN(n)) return "—";
  return Number(n).toFixed(d);
}

// ═══════════════════════════════════════════════════════════════════════════════
// LANDING SCREEN
// ═══════════════════════════════════════════════════════════════════════════════
function LandingScreen({ onSelect }) {
  const modules = [
    {
      id:"concrete", emoji:"🏗️", title:"Concrete Tickets",
      desc:"Scan delivery dockets, track pour volumes by area, validate MPa mix designs, match invoices.",
      color: C.accent,
      stats: ["Volume tracking","MPa validation","Invoice matching","Export to .xlsx"],
    },
    {
      id:"certs", emoji:"📋", title:"Worker Certificates",
      desc:"Scan worker training certifications, track expiry dates, get alerts before certifications lapse.",
      color: C.teal,
      stats: ["25 cert types","Expiry tracking","Worker roster","Email intake ready"],
    },
    {
      id:"tradedocs", emoji:"🦺", title:"Trade Documents",
      desc:"Track each trade's compliance documents — WCB letters, SDS sheets, COR status, and CGL insurance — with automatic compliance checks.",
      color: C.purple,
      stats: ["WCB · SDS · COR · CGL","$5M CGL auto-check","Company roster","Expiry tracking"],
    },
    {
      id:"tm", emoji:"⏱️", title:"Time & Materials",
      desc:"Track extra work, connect field sheets to invoices and change orders, reconcile costs, and submit Southwest labour hours.",
      color: C.yellow,
      stats: ["T&M register","Document matching","Invoice checks","Labour hours"],
    },
  ];

  return (
    <div style={{ minHeight:"100vh", background:C.bg, color:C.text, fontFamily:"'DM Sans','Segoe UI',sans-serif", display:"flex", flexDirection:"column" }}>
      <div style={{ background:C.card, borderBottom:`1px solid ${C.border}`, padding:"20px 32px", display:"flex", alignItems:"center", gap:14 }}>
        <div style={{ width:44, height:44, borderRadius:12, background:`linear-gradient(135deg,${C.accent},${C.teal})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:22 }}>⚡</div>
        <div>
          <div style={{ fontWeight:800, fontSize:19 }}>Southwest Project Tools</div>
          <div style={{ color:C.muted, fontSize:13 }}>Fortuna Project</div>
        </div>
      </div>

      <div style={{ flex:1, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", padding:"40px 24px" }}>
        <div style={{ textAlign:"center", marginBottom:40 }}>
          <div style={{ fontSize:13, fontWeight:700, letterSpacing:2, textTransform:"uppercase", color:C.muted, marginBottom:12 }}>Select a module</div>
          <div style={{ fontSize:28, fontWeight:800, marginBottom:8 }}>What are you working on?</div>
          <div style={{ color:C.muted, fontSize:15 }}>Scan any site document — the app handles the rest.</div>
        </div>

        <div style={{ display:"flex", gap:20, flexWrap:"wrap", justifyContent:"center", width:"100%", maxWidth:780 }}>
          {modules.map(mod => (
            <div key={mod.id} onClick={() => onSelect(mod.id)}
              style={{ background:C.card, border:`2px solid ${mod.color}44`, borderRadius:20, padding:"32px 28px", flex:1, minWidth:280, maxWidth:360, cursor:"pointer", transition:"all .2s", position:"relative", overflow:"hidden" }}
              onMouseEnter={e => e.currentTarget.style.border=`2px solid ${mod.color}`}
              onMouseLeave={e => e.currentTarget.style.border=`2px solid ${mod.color}44`}>
              <div style={{ position:"absolute", top:-20, right:-20, fontSize:80, opacity:.06 }}>{mod.emoji}</div>
              <div style={{ width:54, height:54, borderRadius:15, background:mod.color+"22", border:`1px solid ${mod.color}44`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:26, marginBottom:18 }}>{mod.emoji}</div>
              <div style={{ fontWeight:800, fontSize:20, marginBottom:8, color:mod.color }}>{mod.title}</div>
              <div style={{ color:C.sub, fontSize:13, lineHeight:1.6, marginBottom:20 }}>{mod.desc}</div>
              <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
                {mod.stats.map(s => <Badge key={s} color={mod.color}>{s}</Badge>)}
              </div>
              <div style={{ marginTop:24, background:mod.color, color:"#fff", borderRadius:10, padding:"11px 0", textAlign:"center", fontWeight:800, fontSize:14 }}>
                Open Module →
              </div>
            </div>
          ))}
        </div>

      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// TRAINING CERTIFICATES MODULE
// ═══════════════════════════════════════════════════════════════════════════════
function extractJSON(text) {
  // Strip ALL markdown fences and preamble first
  let cleaned = text.replace(/```json|```/gi,"").trim();
  // Remove any preamble before the first [ or {
  const arrStart = cleaned.indexOf("[");
  const objStart = cleaned.indexOf("{");
  let jsonStart = -1;
  if (arrStart !== -1 && objStart !== -1) jsonStart = Math.min(arrStart, objStart);
  else if (arrStart !== -1) jsonStart = arrStart;
  else if (objStart !== -1) jsonStart = objStart;
  if (jsonStart > 0) cleaned = cleaned.slice(jsonStart);
  // Find matching end bracket
  const isArr = cleaned.startsWith("[");
  const endChar = isArr ? "]" : "}";
  const endIdx = cleaned.lastIndexOf(endChar);
  if (endIdx !== -1) cleaned = cleaned.slice(0, endIdx + 1);
  // Try to parse
  try { return JSON.parse(cleaned); } catch(e) {}
  // Last resort — try original stripped
  const stripped = text.replace(/```json|```/gi,"").trim();
  try { return JSON.parse(stripped); } catch(e) {}
  throw new Error("Could not parse response: " + text.slice(0, 120));
}

function CertsModule({ onBack }) {
  const [certs, setCerts] = useState([]);
  const [tab, setTab] = useState("dashboard");
  const [loading, setLoading] = useState(false);
  const [loadMsg, setLoadMsg] = useState("");
  const [toast, setToast] = useState(null);
  const [drag, setDrag] = useState(false);
  const [selectedWorker, setSelectedWorker] = useState(null);
  const [manualOpen, setManualOpen] = useState(false);
  const [manual, setManual] = useState({ worker_name:"", employer:"", cert_type:"", issued_date:"", expiry_date:"", cert_number:"", notes:"" });
  const [filterExpiry, setFilterExpiry] = useState("all");
  const [storageReady, setStorageReady] = useState(false);
  const fileRef = useRef();

  useEffect(() => {
    storageGet("certs-data").then(saved => {
      if (saved?.certs) setCerts(saved.certs);
      setStorageReady(true);
    });
  }, []);

  useEffect(() => {
    if (!storageReady) return;
    storageSet("certs-data", { certs });
  }, [certs, storageReady]);

  // This module only manages worker certs. Trade documents (added via the
  // Trade Documents module) share the same storage row but are tagged
  // category:"trade" and filtered out here so the two never mix.
  const workerCerts = certs.filter(c => c.category !== "trade");

  function showToast(msg, type="ok") {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3500);
  }

  async function toB64(file) {
    return new Promise((res,rej) => {
      const r = new FileReader(); r.onload = () => res(r.result.split(",")[1]); r.onerror = rej; r.readAsDataURL(file);
    });
  }
  async function toDataURL(file) {
    return new Promise((res,rej) => {
      const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(file);
    });
  }

  async function extractCert(file) {
    const b64 = await toB64(file);
    const isPDF = file.type === "application/pdf";
    const block = isPDF
      ? { type:"document", source:{ type:"base64", media_type:"application/pdf", data:b64 } }
      : { type:"image",    source:{ type:"base64", media_type:file.type, data:b64 } };

    const prompt = `You are a construction safety records assistant. Extract all information from this training certificate or safety document.

Known certificate types: ${CERT_TYPES.join(", ")}

Return ONLY valid JSON (no markdown):
{
  "worker_name": "full name of the worker/certificate holder",
  "employer": "company or employer name if shown, else null",
  "cert_type": "best matching certificate type from the known list, or the type as written if not in list",
  "issued_date": "YYYY-MM-DD or as written, else null",
  "expiry_date": "YYYY-MM-DD or as written, else null — look carefully for expiry/renewal/valid until dates",
  "cert_number": "certificate or card number if shown, else null",
  "issuing_body": "organization that issued the certificate, else null",
  "notes": "any other relevant info such as restrictions or endorsements, else null"
}`;

    const res = await fetch("/api/claude", {
      method:"POST",
      headers:{ "Content-Type":"application/json" },
      body: JSON.stringify({ model:"claude-sonnet-4-6", max_tokens:1000,
        messages:[{ role:"user", content:[block, { type:"text", text:prompt }] }] })
    });
    const data = await res.json();
    if (data.error) throw new Error("API error: " + (data.error.message || JSON.stringify(data.error)));
    const text = data.content?.map(b => b.text||"").join("") || "";
    if (!text) throw new Error("Empty response from API (HTTP " + res.status + ")");
    return extractJSON(text);
  }

  async function handleFiles(files) {
    if (!files?.length) return;
    setLoading(true);
    let added = 0;
    for (const file of Array.from(files)) {
      setLoadMsg(`Reading: "${file.name}"…`);
      try {
        const [extracted, dataURL] = await Promise.all([extractCert(file), toDataURL(file)]);
        setCerts(prev => [...prev, {
          id: Date.now() + Math.random(),
          filename: file.name,
          originalFile: dataURL,
          fileType: file.type,
          added_at: new Date().toISOString(),
          category: "worker",
          ...extracted
        }]);
        added++;
      } catch(e) {
        showToast(`Could not read "${file.name}": ${e.message}`, "err");
      }
    }
    setLoading(false); setLoadMsg("");
    if (added) showToast(`${added} certificate${added>1?"s":""} added ✓`);
  }

  function addManual() {
    if (!manual.worker_name || !manual.cert_type) { showToast("Worker name and cert type are required.", "err"); return; }
    setCerts(prev => [...prev, { id:Date.now(), filename:"Manual entry", added_at:new Date().toISOString(), category:"worker", ...manual }]);
    setManual({ worker_name:"", employer:"", cert_type:"", issued_date:"", expiry_date:"", cert_number:"", notes:"" });
    setManualOpen(false);
    showToast("Certificate added ✓");
  }

  function exportXLSX() {
    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.json_to_sheet(workerCerts.map((c,i) => {
      const s = expiryStatus(c.expiry_date);
      return {
        "#":i+1, "Worker":c.worker_name||"", "Employer":c.employer||"",
        "Certificate Type":c.cert_type||"", "Cert #":c.cert_number||"",
        "Issued":c.issued_date||"", "Expiry":c.expiry_date||"",
        "Status":s.label, "Days Remaining":daysUntilExpiry(c.expiry_date)??""  ,
        "Issuing Body":c.issuing_body||"", "Notes":c.notes||"",
      };
    }));
    ws1["!cols"] = [4,22,22,28,14,14,14,18,14,22,22].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws1, "All Certificates");

    // Expiring soon sheet
    const expiring = workerCerts.filter(c => { const d = daysUntilExpiry(c.expiry_date); return d !== null && d <= 60; })
      .sort((a,b) => daysUntilExpiry(a.expiry_date) - daysUntilExpiry(b.expiry_date));
    if (expiring.length > 0) {
      const ws2 = XLSX.utils.json_to_sheet(expiring.map(c => ({
        "Worker":c.worker_name||"", "Employer":c.employer||"",
        "Certificate Type":c.cert_type||"", "Expiry":c.expiry_date||"",
        "Days Remaining":daysUntilExpiry(c.expiry_date),
        "Status":daysUntilExpiry(c.expiry_date) < 0 ? "EXPIRED" : daysUntilExpiry(c.expiry_date)<=30 ? "CRITICAL" : "WARNING",
      })));
      ws2["!cols"] = [22,22,28,14,14,12].map(w=>({wch:w}));
      XLSX.utils.book_append_sheet(wb, ws2, "⚠ Expiring Soon");
    }

    // Worker roster sheet
    const workers = [...new Set(workerCerts.map(c=>c.worker_name).filter(Boolean))];
    const rosterRows = workers.map(w => {
      const wCerts = workerCerts.filter(c => c.worker_name === w);
      const critical = wCerts.filter(c => { const d = daysUntilExpiry(c.expiry_date); return d!==null && d<=30; });
      return {
        "Worker":w,
        "Employer":wCerts[0]?.employer||"",
        "Total Certs":wCerts.length,
        "Expiring ≤30d":critical.length,
        "Cert Types":wCerts.map(c=>c.cert_type).join(", "),
      };
    });
    const ws3 = XLSX.utils.json_to_sheet(rosterRows);
    ws3["!cols"] = [22,22,12,14,60].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws3, "Worker Roster");

    XLSX.writeFile(wb, `training-certs-${new Date().toISOString().slice(0,10)}.xlsx`);
    showToast("Spreadsheet downloaded ✓");
  }

  // Stats
  const expired  = workerCerts.filter(c => { const d = daysUntilExpiry(c.expiry_date); return d!==null && d<0; });
  const critical = workerCerts.filter(c => { const d = daysUntilExpiry(c.expiry_date); return d!==null && d>=0 && d<=30; });
  const warning  = workerCerts.filter(c => { const d = daysUntilExpiry(c.expiry_date); return d!==null && d>30 && d<=60; });
  const workers  = [...new Set(workerCerts.map(c=>c.worker_name).filter(Boolean))];

  // Filtered cert list
  const filteredCerts = workerCerts.filter(c => {
    if (filterExpiry === "expired")  return daysUntilExpiry(c.expiry_date) !== null && daysUntilExpiry(c.expiry_date) < 0;
    if (filterExpiry === "critical") return daysUntilExpiry(c.expiry_date) !== null && daysUntilExpiry(c.expiry_date) >= 0 && daysUntilExpiry(c.expiry_date) <= 30;
    if (filterExpiry === "warning")  return daysUntilExpiry(c.expiry_date) !== null && daysUntilExpiry(c.expiry_date) > 30 && daysUntilExpiry(c.expiry_date) <= 60;
    return true;
  });

  // Worker detail modal
  function WorkerModal({ name, onClose }) {
    const wCerts = workerCerts.filter(c => c.worker_name === name);
    const employer = wCerts[0]?.employer || "";
    return (
      <div style={{ position:"fixed", inset:0, background:"#000c", zIndex:200, display:"flex", alignItems:"center", justifyContent:"center" }}
        onClick={e => e.target===e.currentTarget && onClose()}>
        <div style={{ background:C.card, border:`1px solid ${C.border}`, borderRadius:16, padding:28, width:"94%", maxWidth:540, maxHeight:"90vh", overflowY:"auto" }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:20 }}>
            <div>
              <div style={{ fontWeight:800, fontSize:18 }}>👷 {name}</div>
              {employer && <div style={{ color:C.muted, fontSize:13 }}>{employer}</div>}
            </div>
            <Badge color={C.teal}>{wCerts.length} cert{wCerts.length!==1?"s":""}</Badge>
          </div>
          {wCerts.sort((a,b) => (daysUntilExpiry(a.expiry_date)??9999) - (daysUntilExpiry(b.expiry_date)??9999)).map(c => {
            const s = expiryStatus(c.expiry_date);
            return (
              <div key={c.id} style={{ background:C.bg, borderRadius:10, padding:"13px 16px", marginBottom:10, border:`1px solid ${s.level==="expired"||s.level==="critical" ? C.red+"44" : s.level==="warning" ? C.yellow+"44" : C.border}` }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8 }}>
                  <div style={{ fontWeight:700, fontSize:14 }}>{c.cert_type}</div>
                  <Badge color={s.color}>{s.label}</Badge>
                </div>
                <div style={{ display:"flex", gap:16, marginTop:7, fontSize:12, color:C.sub, flexWrap:"wrap" }}>
                  {c.issued_date  && <span>Issued: {c.issued_date}</span>}
                  {c.expiry_date  && <span>Expires: {c.expiry_date}</span>}
                  {c.cert_number  && <span>Cert #: {c.cert_number}</span>}
                  {c.issuing_body && <span>By: {c.issuing_body}</span>}
                </div>
                {c.notes && <div style={{ marginTop:6, fontSize:12, color:C.muted, fontStyle:"italic" }}>{c.notes}</div>}
              </div>
            );
          })}
          <button onClick={onClose} style={{ background:C.bg, color:C.muted, border:`1px solid ${C.border}`, borderRadius:9, padding:"10px 20px", fontWeight:700, cursor:"pointer", width:"100%", marginTop:8 }}>Close</button>
        </div>
      </div>
    );
  }

  const TAB = (t, label) => (
    <button onClick={() => setTab(t)} style={{
      padding:"8px 16px", borderRadius:8, cursor:"pointer", fontWeight:700, fontSize:12, border:"none",
      background: tab===t ? C.teal : "transparent", color: tab===t ? "#fff" : C.muted, transition:"all .15s", whiteSpace:"nowrap"
    }}>{label}</button>
  );

  const INPUT = (key, label, type="text", opts=null) => (
    <div style={{ marginBottom:12 }}>
      <label style={{ display:"block", color:C.muted, fontSize:10, fontWeight:700, letterSpacing:1, textTransform:"uppercase", marginBottom:4 }}>{label}</label>
      {opts
        ? <select value={manual[key]} onChange={e => setManual(m=>({...m,[key]:e.target.value}))}
            style={{ width:"100%", background:C.bg, border:`1px solid ${C.border}`, borderRadius:8, padding:"9px 12px", color:C.text, fontSize:14, boxSizing:"border-box" }}>
            <option value="">— select —</option>
            {opts.map(o => <option key={o} value={o}>{o}</option>)}
          </select>
        : <input type={type} value={manual[key]} onChange={e => setManual(m=>({...m,[key]:e.target.value}))}
            style={{ width:"100%", background:C.bg, border:`1px solid ${C.border}`, borderRadius:8, padding:"9px 12px", color:C.text, fontSize:14, boxSizing:"border-box" }} />
      }
    </div>
  );

  return (
    <div style={{ minHeight:"100vh", background:C.bg, color:C.text, fontFamily:"'DM Sans','Segoe UI',sans-serif", paddingBottom:60 }}>
      {toast && (
        <div style={{ position:"fixed", top:18, right:18, zIndex:999,
          background:toast.type==="err"?"#450a0a":"#052e16",
          color:toast.type==="err"?"#fca5a5":"#86efac",
          border:`1px solid ${toast.type==="err"?C.red:C.green}`,
          borderRadius:10, padding:"12px 22px", fontWeight:600, fontSize:14, boxShadow:"0 8px 32px #0009" }}>{toast.msg}</div>
      )}
      {selectedWorker && <WorkerModal name={selectedWorker} onClose={() => setSelectedWorker(null)} />}

      {/* Header */}
      <div style={{ background:C.card, borderBottom:`1px solid ${C.border}`, padding:"16px 28px", display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:12 }}>
        <div style={{ display:"flex", alignItems:"center", gap:13 }}>
          <button onClick={onBack} style={{ background:"transparent", border:`1px solid ${C.border}`, color:C.muted, borderRadius:8, padding:"6px 13px", fontWeight:700, fontSize:12, cursor:"pointer" }}>← Back</button>
          <div style={{ width:40, height:40, borderRadius:11, background:C.teal, display:"flex", alignItems:"center", justifyContent:"center", fontSize:20 }}>📋</div>
          <div>
            <div style={{ fontWeight:800, fontSize:17 }}>Worker Certificates</div>
            <div style={{ color:C.muted, fontSize:12 }}>{workerCerts.length} certificates · {workers.length} workers{expired.length>0?` · ⚠ ${expired.length} expired`:""}{ critical.length>0?` · ⚠ ${critical.length} expiring soon`:""}</div>
          </div>
        </div>
        <div style={{ display:"flex", gap:8, alignItems:"center" }}>
          <span style={{ fontSize:11, color:storageReady?C.green:C.muted, fontWeight:700 }}>{storageReady ? "💾 Auto-saved" : "⏳ Loading..."}</span>
          <button onClick={()=>{ if(!window.confirm("Clear ALL worker certificates? Trade documents are not affected. This cannot be undone.")) return; setCerts(prev=>prev.filter(c=>c.category==="trade")); showToast("Worker certificates cleared."); }} style={{ background:"transparent", color:C.red, border:`1px solid ${C.red}44`, borderRadius:7, padding:"5px 11px", fontSize:11, fontWeight:700, cursor:"pointer" }}>🗑 Clear Data</button>
          <button onClick={exportXLSX} style={{ background:C.green, color:"#052e16", border:"none", borderRadius:9, padding:"10px 22px", fontWeight:800, fontSize:14, cursor:"pointer" }}>⬇ Export .xlsx</button>
        </div>
      </div>

      {/* Alert bar */}
      {(expired.length > 0 || critical.length > 0) && (
        <div style={{ background:"#450a0a", borderBottom:`1px solid ${C.red}`, padding:"10px 28px", display:"flex", alignItems:"center", gap:12, flexWrap:"wrap" }}>
          {expired.length>0  && <span style={{ color:C.red, fontWeight:800, fontSize:13 }}>🚨 {expired.length} expired cert{expired.length>1?"s":""}</span>}
          {critical.length>0 && <span style={{ color:C.yellow, fontWeight:800, fontSize:13 }}>⚠ {critical.length} expiring within 30 days</span>}
          <button onClick={() => { setFilterExpiry("expired"); setTab("certs"); }} style={{ background:C.red+"22", color:C.red, border:`1px solid ${C.red}44`, borderRadius:6, padding:"4px 12px", fontSize:12, fontWeight:700, cursor:"pointer" }}>View →</button>
        </div>
      )}

      {/* Tabs */}
      <div style={{ padding:"14px 28px 0", display:"flex", gap:4, borderBottom:`1px solid ${C.border}`, overflowX:"auto" }}>
        {TAB("dashboard", "📊 Dashboard")}
        {TAB("certs",     `📋 Certificates (${workerCerts.length})`)}
        {TAB("workers",   `👷 Workers (${workers.length})`)}
        {TAB("expiring",  `⚠ Expiring (${expired.length+critical.length+warning.length})`)}
      </div>

      <div style={{ padding:"26px 28px" }}>

        {/* DASHBOARD */}
        {tab==="dashboard" && (
          <div>
            <div style={{ display:"flex", gap:14, flexWrap:"wrap", marginBottom:26 }}>
              <Stat label="Total Certs"    value={certs.length}     sub="on file"                   color={C.teal}   />
              <Stat label="Workers"        value={workers.length}   sub="on roster"                 color={C.blue}   />
              <Stat label="Expired"        value={expired.length}   sub={expired.length>0?"action required":"✓ none"} color={expired.length>0?C.red:C.green} />
              <Stat label="Expiring ≤30d"  value={critical.length}  sub={critical.length>0?"renew soon":"✓ clear"}   color={critical.length>0?C.yellow:C.green} />
            </div>

            {/* Upload zone */}
            <div
              onDragOver={e=>{e.preventDefault();setDrag(true);}}
              onDragLeave={()=>setDrag(false)}
              onDrop={e=>{e.preventDefault();setDrag(false);handleFiles(e.dataTransfer.files);}}
              onClick={()=>fileRef.current.click()}
              style={{ border:`2px dashed ${drag?C.teal:C.border}`, borderRadius:16, padding:"36px 24px", textAlign:"center", cursor:"pointer", background:drag?C.teal+"11":C.card, transition:"all .2s", marginBottom:16 }}>
              <div style={{ fontSize:36, marginBottom:10 }}>📎</div>
              <div style={{ fontWeight:700, fontSize:16, marginBottom:6 }}>Scan a Certificate</div>
              <div style={{ color:C.muted, fontSize:13, marginBottom:6 }}>Photo or PDF · Claude reads worker name, cert type, and expiry automatically</div>
              <div style={{ color:C.muted, fontSize:12 }}>Drag & drop or click to browse</div>
              <input ref={fileRef} type="file" multiple accept="image/*,application/pdf" style={{ display:"none" }} onChange={e=>handleFiles(e.target.files)} />
            </div>

            {loading && (
              <div style={{ background:"#1e3a5f", border:`1px solid ${C.blue}`, borderRadius:11, padding:"13px 20px", color:"#93c5fd", fontWeight:600, marginBottom:16 }}>
                ⏳ {loadMsg || "Processing…"}
              </div>
            )}

            <div style={{ display:"flex", gap:12, flexWrap:"wrap", marginBottom:28 }}>
              <button onClick={()=>setManualOpen(true)} style={{ background:C.card, border:`1px solid ${C.border}`, color:C.text, borderRadius:9, padding:"10px 20px", fontWeight:700, cursor:"pointer", fontSize:13 }}>✏️ Add Manually</button>
            </div>

            {/* Email intake note */}
            <div style={{ background:"#1a1040", border:`1px solid ${C.purple}44`, borderRadius:14, padding:"20px 24px" }}>
              <div style={{ fontWeight:700, fontSize:15, marginBottom:8 }}>📧 Email Intake — Coming Next</div>
              <div style={{ color:C.sub, fontSize:13, lineHeight:1.7 }}>
                Certificates often arrive by email. The next step is setting up a dedicated intake address (shared Gmail) so certificates emailed by workers or employers are automatically routed into this module — no manual upload needed. This uses Gmail + Google Apps Script and is completely free.
              </div>
            </div>
          </div>
        )}

        {/* CERTIFICATES LIST */}
        {tab==="certs" && (
          <div>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16, flexWrap:"wrap", gap:12 }}>
              <div style={{ fontWeight:700, fontSize:18 }}>Certificate Log</div>
              <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
                {["all","expired","critical","warning"].map(f => (
                  <button key={f} onClick={()=>setFilterExpiry(f)} style={{ background:filterExpiry===f?C.teal:"transparent", color:filterExpiry===f?"#fff":C.muted, border:`1px solid ${filterExpiry===f?C.teal:C.border}`, borderRadius:7, padding:"5px 14px", fontSize:12, fontWeight:700, cursor:"pointer" }}>
                    {f==="all"?"All":f==="expired"?"Expired":f==="critical"?"≤30 days":"31-60 days"}
                  </button>
                ))}
              </div>
            </div>
            {filteredCerts.length===0
              ? <div style={{ color:C.muted, textAlign:"center", padding:"60px 0" }}>No certificates match this filter.</div>
              : [...filteredCerts].sort((a,b)=>(daysUntilExpiry(a.expiry_date)??9999)-(daysUntilExpiry(b.expiry_date)??9999)).map(c => {
                const s = expiryStatus(c.expiry_date);
                return (
                  <div key={c.id} onClick={()=>setSelectedWorker(c.worker_name)} style={{ background:C.card, border:`1px solid ${s.level==="expired"||s.level==="critical"?C.red+"55":s.level==="warning"?C.yellow+"44":C.border}`, borderRadius:12, padding:"15px 20px", marginBottom:10, cursor:"pointer" }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8, marginBottom:7 }}>
                      <div>
                        <span style={{ fontWeight:800, fontSize:15 }}>👷 {c.worker_name||"Unknown Worker"}</span>
                        {c.employer && <span style={{ color:C.muted, fontSize:12, marginLeft:10 }}>{c.employer}</span>}
                      </div>
                      <div style={{ display:"flex", gap:7, alignItems:"center" }}>
                        <Badge color={s.color}>{s.label}</Badge>
                        <button onClick={e=>{ e.stopPropagation(); if(window.confirm(`Delete cert for ${c.worker_name}?`)) setCerts(prev=>prev.filter(x=>x.id!==c.id)); }} style={{ background:"transparent", border:`1px solid ${C.red}44`, color:C.red, borderRadius:6, padding:"3px 9px", fontSize:12, fontWeight:700, cursor:"pointer" }}>✕</button>
                      </div>
                    </div>
                    <div style={{ fontWeight:600, color:C.teal, fontSize:14, marginBottom:6 }}>{c.cert_type||"Unknown Cert"}</div>
                    <div style={{ display:"flex", gap:16, fontSize:12, color:C.sub, flexWrap:"wrap" }}>
                      {c.issued_date  && <span>Issued: {c.issued_date}</span>}
                      {c.expiry_date  && <span>Expires: {c.expiry_date}</span>}
                      {c.cert_number  && <span>Cert #: {c.cert_number}</span>}
                      {c.issuing_body && <span>By: {c.issuing_body}</span>}
                    </div>
                  </div>
                );
              })
            }
          </div>
        )}

        {/* WORKERS */}
        {tab==="workers" && (
          <div>
            <div style={{ fontWeight:700, fontSize:18, marginBottom:20 }}>Worker Roster</div>
            {workers.length===0
              ? <div style={{ color:C.muted, textAlign:"center", padding:"60px 0" }}>No workers yet. Scan a certificate to add to the roster.</div>
              : workers.map(w => {
                const wCerts = certs.filter(c=>c.worker_name===w);
                const wExpired  = wCerts.filter(c=>{ const d=daysUntilExpiry(c.expiry_date); return d!==null&&d<0; });
                const wCritical = wCerts.filter(c=>{ const d=daysUntilExpiry(c.expiry_date); return d!==null&&d>=0&&d<=30; });
                const hasIssues = wExpired.length>0||wCritical.length>0;
                return (
                  <div key={w} onClick={()=>setSelectedWorker(w)} style={{ background:C.card, border:`1px solid ${hasIssues?C.red+"55":C.border}`, borderRadius:12, padding:"16px 20px", marginBottom:10, cursor:"pointer" }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8, marginBottom:8 }}>
                      <div>
                        <span style={{ fontWeight:800, fontSize:15 }}>👷 {w}</span>
                        {wCerts[0]?.employer && <span style={{ color:C.muted, fontSize:12, marginLeft:10 }}>{wCerts[0].employer}</span>}
                      </div>
                      <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                        <Badge color={C.teal}>{wCerts.length} cert{wCerts.length!==1?"s":""}</Badge>
                        {wExpired.length>0  && <Badge color={C.red}>{wExpired.length} expired</Badge>}
                        {wCritical.length>0 && <Badge color={C.yellow}>{wCritical.length} expiring soon</Badge>}
                      </div>
                    </div>
                    <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                      {wCerts.map(c => {
                        const s = expiryStatus(c.expiry_date);
                        return <Badge key={c.id} color={s.color}>{c.cert_type?.split(" ")[0]||"Cert"}</Badge>;
                      })}
                    </div>
                    <div style={{ marginTop:7, color:C.muted, fontSize:11 }}>Tap to view full certificate list →</div>
                  </div>
                );
              })
            }
          </div>
        )}

        {/* EXPIRING */}
        {tab==="expiring" && (
          <div>
            <div style={{ fontWeight:700, fontSize:18, marginBottom:6 }}>Expiry Alerts</div>
            <div style={{ color:C.muted, fontSize:13, marginBottom:22 }}>Certificates expired or expiring within 60 days — review with Site Super and Safety Rep</div>

            {expired.length===0 && critical.length===0 && warning.length===0
              ? <div style={{ background:"#052e16", border:`1px solid ${C.green}44`, borderRadius:14, padding:"32px 24px", textAlign:"center" }}>
                  <div style={{ fontSize:32, marginBottom:10 }}>✅</div>
                  <div style={{ fontWeight:700, fontSize:16, color:C.green }}>All certifications are current</div>
                  <div style={{ color:C.muted, fontSize:13, marginTop:6 }}>No expirations within 60 days</div>
                </div>
              : <>
                {[
                  { group: expired,  label:"🚨 Expired", color: C.red    },
                  { group: critical, label:"⚠ Expiring Within 30 Days", color: C.red    },
                  { group: warning,  label:"⏰ Expiring in 31–60 Days",  color: C.yellow },
                ].map(({ group, label, color }) => group.length > 0 && (
                  <div key={label} style={{ marginBottom:28 }}>
                    <div style={{ fontWeight:700, fontSize:14, color, marginBottom:12, display:"flex", alignItems:"center", gap:8 }}>
                      {label} <Badge color={color}>{group.length}</Badge>
                    </div>
                    {group.map(c => {
                      const s = expiryStatus(c.expiry_date);
                      return (
                        <div key={c.id} onClick={()=>setSelectedWorker(c.worker_name)} style={{ background:C.card, border:`1px solid ${color}44`, borderRadius:11, padding:"14px 18px", marginBottom:9, cursor:"pointer" }}>
                          <div style={{ display:"flex", justifyContent:"space-between", flexWrap:"wrap", gap:8, marginBottom:6 }}>
                            <span style={{ fontWeight:800 }}>👷 {c.worker_name||"Unknown"}</span>
                            <Badge color={color}>{s.label}</Badge>
                          </div>
                          <div style={{ color:C.teal, fontWeight:600, fontSize:13, marginBottom:4 }}>{c.cert_type}</div>
                          <div style={{ fontSize:12, color:C.sub }}>
                            {c.employer && <span style={{ marginRight:12 }}>{c.employer}</span>}
                            {c.expiry_date && <span>Expires: {c.expiry_date}</span>}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                ))}
              </>
            }
          </div>
        )}
      </div>

      {/* Manual entry modal */}
      {manualOpen && (
        <div style={{ position:"fixed", inset:0, background:"#000b", zIndex:100, display:"flex", alignItems:"center", justifyContent:"center" }}
          onClick={e=>e.target===e.currentTarget&&setManualOpen(false)}>
          <div style={{ background:C.card, border:`1px solid ${C.border}`, borderRadius:16, padding:30, width:"92%", maxWidth:500, maxHeight:"90vh", overflowY:"auto" }}>
            <div style={{ fontWeight:800, fontSize:18, marginBottom:20 }}>✏️ Add Certificate Manually</div>
            {INPUT("worker_name","Worker Full Name *")}
            {INPUT("employer","Employer / Company")}
            {INPUT("cert_type","Certificate Type *","text",CERT_TYPES)}
            {INPUT("cert_number","Certificate / Card Number")}
            {INPUT("issued_date","Issue Date","date")}
            {INPUT("expiry_date","Expiry Date","date")}
            {INPUT("notes","Notes")}
            {manual.expiry_date && (() => {
              const s = expiryStatus(manual.expiry_date);
              if (s.level==="none") return null;
              return <div style={{ background:s.color+"15", border:`1px solid ${s.color}44`, borderRadius:8, padding:"9px 14px", marginBottom:12, fontSize:13, color:s.color, fontWeight:600 }}>{s.label}</div>;
            })()}
            <div style={{ display:"flex", gap:10, marginTop:18 }}>
              <button onClick={addManual} style={{ background:C.teal, color:"#fff", border:"none", borderRadius:9, padding:"11px 0", fontWeight:800, cursor:"pointer", flex:1 }}>Add Certificate</button>
              <button onClick={()=>setManualOpen(false)} style={{ background:C.bg, color:C.muted, border:`1px solid ${C.border}`, borderRadius:9, padding:"11px 18px", fontWeight:700, cursor:"pointer" }}>Cancel</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// TRADE SAFETY DOCUMENTS MODULE
// ═══════════════════════════════════════════════════════════════════════════════
function TradeDocsModule({ onBack }) {
  const [certs, setCerts] = useState([]);
  const [tab, setTab] = useState("dashboard");
  const [loading, setLoading] = useState(false);
  const [loadMsg, setLoadMsg] = useState("");
  const [toast, setToast] = useState(null);
  const [drag, setDrag] = useState(false);
  const [selectedCompany, setSelectedCompany] = useState(null);
  const [manualOpen, setManualOpen] = useState(false);
  const [manual, setManual] = useState({ company_name:"", doc_type:"", issued_date:"", expiry_date:"", doc_number:"", issuing_body:"", per_occurrence_limit:"", certificate_holder:"", additional_insured:"", notes:"" });
  const [storageReady, setStorageReady] = useState(false);
  const fileRef = useRef();

  useEffect(() => {
    storageGet("certs-data").then(saved => {
      if (saved?.certs) setCerts(saved.certs);
      setStorageReady(true);
    });
  }, []);

  useEffect(() => {
    if (!storageReady) return;
    storageSet("certs-data", { certs });
  }, [certs, storageReady]);

  // This module only manages trade/company documents (category:"trade").
  // Worker certs (managed by the Worker Certificates module) share the same
  // storage row but are filtered out here so the two never mix.
  const tradeDocs = certs.filter(c => c.category === "trade");

  function showToast(msg, type="ok") {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3500);
  }

  async function toB64(file) {
    return new Promise((res,rej) => {
      const r = new FileReader(); r.onload = () => res(r.result.split(",")[1]); r.onerror = rej; r.readAsDataURL(file);
    });
  }
  async function toDataURL(file) {
    return new Promise((res,rej) => {
      const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(file);
    });
  }

  async function extractTradeDoc(file) {
    const b64 = await toB64(file);
    const isPDF = file.type === "application/pdf";
    const block = isPDF
      ? { type:"document", source:{ type:"base64", media_type:"application/pdf", data:b64 } }
      : { type:"image",    source:{ type:"base64", media_type:file.type, data:b64 } };

    const prompt = `You are a construction compliance assistant. Extract all information from this trade/subcontractor Safety Program Document.

Known document types: ${TRADE_DOC_TYPES.join(", ")}

Return ONLY valid JSON (no markdown):
{
  "company_name": "the trade/subcontractor company this document belongs to",
  "doc_type": "best matching document type from the known list, or as written if not listed",
  "issued_date": "YYYY-MM-DD or as written, else null",
  "expiry_date": "YYYY-MM-DD or as written, else null — look for expiry/renewal/valid until dates",
  "doc_number": "policy, certificate, clearance, or letter number if shown, else null",
  "issuing_body": "organization/insurer/board/broker that issued this document, else null",
  "per_occurrence_limit": "ONLY for CGL insurance certificates — the per-occurrence coverage limit as a plain number with no currency symbols or commas, else null",
  "certificate_holder": "ONLY for CGL insurance certificates — the certificate holder name exactly as shown, else null",
  "additional_insured": "ONLY for CGL insurance certificates — array of all additional insured names exactly as listed, else null",
  "notes": "any other relevant info such as restrictions, exclusions, or conditions, else null"
}`;

    const res = await fetch("/api/claude", {
      method:"POST",
      headers:{ "Content-Type":"application/json" },
      body: JSON.stringify({ model:"claude-sonnet-4-6", max_tokens:1000,
        messages:[{ role:"user", content:[block, { type:"text", text:prompt }] }] })
    });
    const data = await res.json();
    if (data.error) throw new Error("API error: " + (data.error.message || JSON.stringify(data.error)));
    const text = data.content?.map(b => b.text||"").join("") || "";
    if (!text) throw new Error("Empty response from API (HTTP " + res.status + ")");
    return extractJSON(text);
  }

  async function handleFiles(files) {
    if (!files?.length) return;
    setLoading(true);
    let added = 0;
    for (const file of Array.from(files)) {
      setLoadMsg(`Reading: "${file.name}"…`);
      try {
        const [extracted, dataURL] = await Promise.all([extractTradeDoc(file), toDataURL(file)]);
        setCerts(prev => [...prev, {
          id: Date.now() + Math.random(),
          filename: file.name,
          originalFile: dataURL,
          fileType: file.type,
          added_at: new Date().toISOString(),
          category: "trade",
          ...extracted
        }]);
        added++;
      } catch(e) {
        showToast(`Could not read "${file.name}": ${e.message}`, "err");
      }
    }
    setLoading(false); setLoadMsg("");
    if (added) showToast(`${added} document${added>1?"s":""} added ✓`);
  }

  function addManual() {
    if (!manual.company_name || !manual.doc_type) { showToast("Company name and document type are required.", "err"); return; }
    const additional_insured = manual.additional_insured ? manual.additional_insured.split(",").map(s=>s.trim()).filter(Boolean) : [];
    setCerts(prev => [...prev, { id:Date.now(), filename:"Manual entry", added_at:new Date().toISOString(), category:"trade", ...manual, additional_insured }]);
    setManual({ company_name:"", doc_type:"", issued_date:"", expiry_date:"", doc_number:"", issuing_body:"", per_occurrence_limit:"", certificate_holder:"", additional_insured:"", notes:"" });
    setManualOpen(false);
    showToast("Document added ✓");
  }

  function exportXLSX() {
    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.json_to_sheet(tradeDocs.map((c,i) => {
      const s = expiryStatus(c.expiry_date);
      const flags = cglComplianceFlags(c);
      return {
        "#":i+1, "Company":c.company_name||"", "Document Type":c.doc_type||"",
        "Doc #":c.doc_number||"", "Issued":c.issued_date||"", "Expiry":c.expiry_date||"",
        "Status":s.label, "Days Remaining":daysUntilExpiry(c.expiry_date)??"",
        "Issuing Body":c.issuing_body||"",
        "CGL Limit":c.per_occurrence_limit||"", "CGL Certificate Holder":c.certificate_holder||"",
        "CGL Additional Insured":Array.isArray(c.additional_insured)?c.additional_insured.join("; "):(c.additional_insured||""),
        "Compliance Issues":flags.join("; "),
        "Notes":c.notes||"",
      };
    }));
    ws1["!cols"] = [4,24,26,14,14,14,18,14,22,14,26,30,40,22].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws1, "All Trade Documents");

    const issues = tradeDocs.filter(c => cglComplianceFlags(c).length > 0);
    if (issues.length > 0) {
      const ws2 = XLSX.utils.json_to_sheet(issues.map(c => ({
        "Company":c.company_name||"", "Document Type":c.doc_type||"",
        "Compliance Issues":cglComplianceFlags(c).join("; "),
      })));
      ws2["!cols"] = [24,26,60].map(w=>({wch:w}));
      XLSX.utils.book_append_sheet(wb, ws2, "⚠ Compliance Issues");
    }

    const expiring = tradeDocs.filter(c => { const d = daysUntilExpiry(c.expiry_date); return d !== null && d <= 60; })
      .sort((a,b) => daysUntilExpiry(a.expiry_date) - daysUntilExpiry(b.expiry_date));
    if (expiring.length > 0) {
      const ws3 = XLSX.utils.json_to_sheet(expiring.map(c => ({
        "Company":c.company_name||"", "Document Type":c.doc_type||"", "Expiry":c.expiry_date||"",
        "Days Remaining":daysUntilExpiry(c.expiry_date),
        "Status":daysUntilExpiry(c.expiry_date) < 0 ? "EXPIRED" : daysUntilExpiry(c.expiry_date)<=30 ? "CRITICAL" : "WARNING",
      })));
      ws3["!cols"] = [24,26,14,14,12].map(w=>({wch:w}));
      XLSX.utils.book_append_sheet(wb, ws3, "⚠ Expiring Soon");
    }

    XLSX.writeFile(wb, `trade-safety-documents-${new Date().toISOString().slice(0,10)}.xlsx`);
    showToast("Spreadsheet downloaded ✓");
  }

  // Stats
  const expired  = tradeDocs.filter(c => { const d = daysUntilExpiry(c.expiry_date); return d!==null && d<0; });
  const critical = tradeDocs.filter(c => { const d = daysUntilExpiry(c.expiry_date); return d!==null && d>=0 && d<=30; });
  const warning  = tradeDocs.filter(c => { const d = daysUntilExpiry(c.expiry_date); return d!==null && d>30 && d<=60; });
  const companies = [...new Set(tradeDocs.map(c=>c.company_name).filter(Boolean))];
  const complianceIssues = tradeDocs.filter(c => cglComplianceFlags(c).length > 0);

  // Company detail modal
  function CompanyModal({ name, onClose }) {
    const cDocs = tradeDocs.filter(c => c.company_name === name);
    return (
      <div style={{ position:"fixed", inset:0, background:"#000c", zIndex:200, display:"flex", alignItems:"center", justifyContent:"center" }}
        onClick={e => e.target===e.currentTarget && onClose()}>
        <div style={{ background:C.card, border:`1px solid ${C.border}`, borderRadius:16, padding:28, width:"94%", maxWidth:560, maxHeight:"90vh", overflowY:"auto" }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:20 }}>
            <div style={{ fontWeight:800, fontSize:18 }}>🏢 {name}</div>
            <Badge color={C.purple}>{cDocs.length} doc{cDocs.length!==1?"s":""}</Badge>
          </div>
          {cDocs.sort((a,b) => (daysUntilExpiry(a.expiry_date)??9999) - (daysUntilExpiry(b.expiry_date)??9999)).map(c => {
            const s = expiryStatus(c.expiry_date);
            const flags = cglComplianceFlags(c);
            return (
              <div key={c.id} style={{ background:C.bg, borderRadius:10, padding:"13px 16px", marginBottom:10, border:`1px solid ${flags.length>0 ? C.red+"66" : s.level==="expired"||s.level==="critical" ? C.red+"44" : s.level==="warning" ? C.yellow+"44" : C.border}` }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8 }}>
                  <div style={{ fontWeight:700, fontSize:14 }}>{c.doc_type}</div>
                  <div style={{ display:"flex", gap:7, alignItems:"center" }}>
                    {(c.file_url||c.originalFile)&&<button onClick={e=>{ e.stopPropagation(); const src=c.file_url||c.originalFile; const isImg=/^data:image|\.(jpg|jpeg|png|gif|webp|heic)/i.test(src); const w=window.open(); w.document.write(isImg?`<html><body style="margin:0;background:#000;display:flex;align-items:center;justify-content:center;min-height:100vh"><img src="${src}" style="max-width:100%;max-height:100vh;object-fit:contain"></body></html>`:`<iframe src="${src}" width="100%" height="100%" style="border:none;position:fixed;top:0;left:0"></iframe>`); }} style={{background:"transparent",border:`1px solid ${C.blue}44`,color:C.blue,borderRadius:6,padding:"3px 9px",fontSize:12,fontWeight:700,cursor:"pointer"}}>📄 View</button>}
                    <Badge color={s.color}>{s.label}</Badge>
                  </div>
                </div>
                <div style={{ display:"flex", gap:16, marginTop:7, fontSize:12, color:C.sub, flexWrap:"wrap" }}>
                  {c.issued_date  && <span>Issued: {c.issued_date}</span>}
                  {c.expiry_date  && <span>Expires: {c.expiry_date}</span>}
                  {c.doc_number   && <span>Doc #: {c.doc_number}</span>}
                  {c.issuing_body && <span>By: {c.issuing_body}</span>}
                </div>
                {c.doc_type === "CGL Insurance Certificate" && (
                  <div style={{ marginTop:8, fontSize:12, color:C.sub }}>
                    {c.per_occurrence_limit && <div>Per-occurrence limit: {c.per_occurrence_limit}</div>}
                    {c.certificate_holder && <div>Certificate holder: {c.certificate_holder}</div>}
                    {c.additional_insured && <div>Additional insured: {Array.isArray(c.additional_insured)?c.additional_insured.join(", "):c.additional_insured}</div>}
                  </div>
                )}
                {flags.length > 0 && (
                  <div style={{ marginTop:8, background:"#450a0a", border:`1px solid ${C.red}44`, borderRadius:7, padding:"8px 12px" }}>
                    {flags.map((f,i) => <div key={i} style={{ color:"#fca5a5", fontSize:12, fontWeight:600 }}>⚠ {f}</div>)}
                  </div>
                )}
                {c.notes && <div style={{ marginTop:6, fontSize:12, color:C.muted, fontStyle:"italic" }}>{c.notes}</div>}
              </div>
            );
          })}
          <button onClick={onClose} style={{ background:C.bg, color:C.muted, border:`1px solid ${C.border}`, borderRadius:9, padding:"10px 20px", fontWeight:700, cursor:"pointer", width:"100%", marginTop:8 }}>Close</button>
        </div>
      </div>
    );
  }

  const TAB = (t, label) => (
    <button onClick={() => setTab(t)} style={{
      padding:"8px 16px", borderRadius:8, cursor:"pointer", fontWeight:700, fontSize:12, border:"none",
      background: tab===t ? C.purple : "transparent", color: tab===t ? "#fff" : C.muted, transition:"all .15s", whiteSpace:"nowrap"
    }}>{label}</button>
  );

  const INPUT = (key, label, type="text", opts=null) => (
    <div style={{ marginBottom:12 }}>
      <label style={{ display:"block", color:C.muted, fontSize:10, fontWeight:700, letterSpacing:1, textTransform:"uppercase", marginBottom:4 }}>{label}</label>
      {opts
        ? <select value={manual[key]} onChange={e => setManual(m=>({...m,[key]:e.target.value}))}
            style={{ width:"100%", background:C.bg, border:`1px solid ${C.border}`, borderRadius:8, padding:"9px 12px", color:C.text, fontSize:14, boxSizing:"border-box" }}>
            <option value="">— select —</option>
            {opts.map(o => <option key={o} value={o}>{o}</option>)}
          </select>
        : <input type={type} value={manual[key]} onChange={e => setManual(m=>({...m,[key]:e.target.value}))}
            style={{ width:"100%", background:C.bg, border:`1px solid ${C.border}`, borderRadius:8, padding:"9px 12px", color:C.text, fontSize:14, boxSizing:"border-box" }} />
      }
    </div>
  );

  return (
    <div style={{ minHeight:"100vh", background:C.bg, color:C.text, fontFamily:"'DM Sans','Segoe UI',sans-serif", paddingBottom:60 }}>
      {toast && (
        <div style={{ position:"fixed", top:18, right:18, zIndex:999,
          background:toast.type==="err"?"#450a0a":"#052e16",
          color:toast.type==="err"?"#fca5a5":"#86efac",
          border:`1px solid ${toast.type==="err"?C.red:C.green}`,
          borderRadius:10, padding:"12px 22px", fontWeight:600, fontSize:14, boxShadow:"0 8px 32px #0009" }}>{toast.msg}</div>
      )}
      {selectedCompany && <CompanyModal name={selectedCompany} onClose={() => setSelectedCompany(null)} />}

      {/* Header */}
      <div style={{ background:C.card, borderBottom:`1px solid ${C.border}`, padding:"16px 28px", display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:12 }}>
        <div style={{ display:"flex", alignItems:"center", gap:13 }}>
          <button onClick={onBack} style={{ background:"transparent", border:`1px solid ${C.border}`, color:C.muted, borderRadius:8, padding:"6px 13px", fontWeight:700, fontSize:12, cursor:"pointer" }}>← Back</button>
          <div style={{ width:40, height:40, borderRadius:11, background:C.purple, display:"flex", alignItems:"center", justifyContent:"center", fontSize:20 }}>🦺</div>
          <div>
            <div style={{ fontWeight:800, fontSize:17 }}>Trade Documents</div>
            <div style={{ color:C.muted, fontSize:12 }}>{tradeDocs.length} documents · {companies.length} companies{expired.length>0?` · ⚠ ${expired.length} expired`:""}{ complianceIssues.length>0?` · 🚨 ${complianceIssues.length} compliance issue${complianceIssues.length>1?"s":""}`:""}</div>
          </div>
        </div>
        <div style={{ display:"flex", gap:8, alignItems:"center" }}>
          <span style={{ fontSize:11, color:storageReady?C.green:C.muted, fontWeight:700 }}>{storageReady ? "💾 Auto-saved" : "⏳ Loading..."}</span>
          <button onClick={()=>{ if(!window.confirm("Clear ALL trade documents? Worker certificates are not affected. This cannot be undone.")) return; setCerts(prev=>prev.filter(c=>c.category!=="trade")); showToast("Trade documents cleared."); }} style={{ background:"transparent", color:C.red, border:`1px solid ${C.red}44`, borderRadius:7, padding:"5px 11px", fontSize:11, fontWeight:700, cursor:"pointer" }}>🗑 Clear Data</button>
          <button onClick={exportXLSX} style={{ background:C.green, color:"#052e16", border:"none", borderRadius:9, padding:"10px 22px", fontWeight:800, fontSize:14, cursor:"pointer" }}>⬇ Export .xlsx</button>
        </div>
      </div>

      {/* Alert bar */}
      {(complianceIssues.length > 0 || expired.length > 0) && (
        <div style={{ background:"#450a0a", borderBottom:`1px solid ${C.red}`, padding:"10px 28px", display:"flex", alignItems:"center", gap:12, flexWrap:"wrap" }}>
          {complianceIssues.length>0 && <span style={{ color:C.red, fontWeight:800, fontSize:13 }}>🚨 {complianceIssues.length} CGL compliance issue{complianceIssues.length>1?"s":""}</span>}
          {expired.length>0  && <span style={{ color:C.yellow, fontWeight:800, fontSize:13 }}>⚠ {expired.length} expired document{expired.length>1?"s":""}</span>}
          <button onClick={() => setTab("compliance")} style={{ background:C.red+"22", color:C.red, border:`1px solid ${C.red}44`, borderRadius:6, padding:"4px 12px", fontSize:12, fontWeight:700, cursor:"pointer" }}>View →</button>
        </div>
      )}

      {/* Tabs */}
      <div style={{ padding:"14px 28px 0", display:"flex", gap:4, borderBottom:`1px solid ${C.border}`, overflowX:"auto" }}>
        {TAB("dashboard",   "📊 Dashboard")}
        {TAB("docs",        `🦺 Documents (${tradeDocs.length})`)}
        {TAB("companies",   `🏢 Companies (${companies.length})`)}
        {TAB("compliance",  `🚨 Compliance (${complianceIssues.length})`)}
        {TAB("expiring",    `⚠ Expiring (${expired.length+critical.length+warning.length})`)}
      </div>

      <div style={{ padding:"26px 28px" }}>

        {/* DASHBOARD */}
        {tab==="dashboard" && (
          <div>
            <div style={{ display:"flex", gap:14, flexWrap:"wrap", marginBottom:26 }}>
              <Stat label="Total Documents" value={tradeDocs.length}      sub="on file"                   color={C.purple} />
              <Stat label="Companies"       value={companies.length}      sub="tracked"                   color={C.blue}   />
              <Stat label="Compliance Issues" value={complianceIssues.length} sub={complianceIssues.length>0?"CGL review needed":"✓ none"} color={complianceIssues.length>0?C.red:C.green} />
              <Stat label="Expired"         value={expired.length}        sub={expired.length>0?"action required":"✓ none"} color={expired.length>0?C.red:C.green} />
            </div>

            {/* Upload zone */}
            <div
              onDragOver={e=>{e.preventDefault();setDrag(true);}}
              onDragLeave={()=>setDrag(false)}
              onDrop={e=>{e.preventDefault();setDrag(false);handleFiles(e.dataTransfer.files);}}
              onClick={()=>fileRef.current.click()}
              style={{ border:`2px dashed ${drag?C.purple:C.border}`, borderRadius:16, padding:"36px 24px", textAlign:"center", cursor:"pointer", background:drag?C.purple+"11":C.card, transition:"all .2s", marginBottom:16 }}>
              <div style={{ fontSize:36, marginBottom:10 }}>📎</div>
              <div style={{ fontWeight:700, fontSize:16, marginBottom:6 }}>Scan a Trade Document</div>
              <div style={{ color:C.muted, fontSize:13, marginBottom:6 }}>WCB letter, SDS, COR cert, or CGL insurance certificate · Photo or PDF</div>
              <div style={{ color:C.muted, fontSize:12 }}>Drag & drop or click to browse</div>
              <input ref={fileRef} type="file" multiple accept="image/*,application/pdf" style={{ display:"none" }} onChange={e=>handleFiles(e.target.files)} />
            </div>

            {loading && (
              <div style={{ background:"#1e3a5f", border:`1px solid ${C.blue}`, borderRadius:11, padding:"13px 20px", color:"#93c5fd", fontWeight:600, marginBottom:16 }}>
                ⏳ {loadMsg || "Processing…"}
              </div>
            )}

            <div style={{ display:"flex", gap:12, flexWrap:"wrap", marginBottom:28 }}>
              <button onClick={()=>setManualOpen(true)} style={{ background:C.card, border:`1px solid ${C.border}`, color:C.text, borderRadius:9, padding:"10px 20px", fontWeight:700, cursor:"pointer", fontSize:13 }}>✏️ Add Manually</button>
            </div>

            <div style={{ background:"#1a1040", border:`1px solid ${C.purple}44`, borderRadius:14, padding:"20px 24px" }}>
              <div style={{ fontWeight:700, fontSize:15, marginBottom:8 }}>🚨 CGL Compliance Check</div>
              <div style={{ color:C.sub, fontSize:13, lineHeight:1.7 }}>
                Every CGL Insurance Certificate is automatically checked for a per-occurrence limit of at least $5,000,000, with certificate holder "{CGL_CERT_HOLDER_LABEL}" and additional insured listing both "{CGL_ADDITIONAL_INSURED_LABELS[0]}" and "{CGL_ADDITIONAL_INSURED_LABELS[1]}". Anything that doesn't match shows up in the Compliance tab.
              </div>
            </div>
          </div>
        )}

        {/* DOCUMENTS LIST */}
        {tab==="docs" && (
          <div>
            <div style={{ fontWeight:700, fontSize:18, marginBottom:16 }}>Document Log</div>
            {tradeDocs.length===0
              ? <div style={{ color:C.muted, textAlign:"center", padding:"60px 0" }}>No trade documents yet. Scan one to get started.</div>
              : [...tradeDocs].sort((a,b)=>(daysUntilExpiry(a.expiry_date)??9999)-(daysUntilExpiry(b.expiry_date)??9999)).map(c => {
                const s = expiryStatus(c.expiry_date);
                const flags = cglComplianceFlags(c);
                return (
                  <div key={c.id} onClick={()=>setSelectedCompany(c.company_name)} style={{ background:C.card, border:`1px solid ${flags.length>0?C.red+"66":s.level==="expired"||s.level==="critical"?C.red+"55":s.level==="warning"?C.yellow+"44":C.border}`, borderRadius:12, padding:"15px 20px", marginBottom:10, cursor:"pointer" }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8, marginBottom:7 }}>
                      <div>
                        <span style={{ fontWeight:800, fontSize:15 }}>🏢 {c.company_name||"Unknown Company"}</span>
                      </div>
                      <div style={{ display:"flex", gap:7, alignItems:"center" }}>
                        {flags.length>0 && <Badge color={C.red}>⚠ compliance</Badge>}
                        <Badge color={s.color}>{s.label}</Badge>
                        {(c.file_url||c.originalFile)&&<button onClick={e=>{ e.stopPropagation(); const src=c.file_url||c.originalFile; const isImg=/^data:image|\.(jpg|jpeg|png|gif|webp|heic)/i.test(src); const w=window.open(); w.document.write(isImg?`<html><body style="margin:0;background:#000;display:flex;align-items:center;justify-content:center;min-height:100vh"><img src="${src}" style="max-width:100%;max-height:100vh;object-fit:contain"></body></html>`:`<iframe src="${src}" width="100%" height="100%" style="border:none;position:fixed;top:0;left:0"></iframe>`); }} style={{background:"transparent",border:`1px solid ${C.blue}44`,color:C.blue,borderRadius:6,padding:"3px 9px",fontSize:12,fontWeight:700,cursor:"pointer"}}>📄 View</button>}
                        <button onClick={e=>{ e.stopPropagation(); if(window.confirm(`Delete document for ${c.company_name}?`)) setCerts(prev=>prev.filter(x=>x.id!==c.id)); }} style={{ background:"transparent", border:`1px solid ${C.red}44`, color:C.red, borderRadius:6, padding:"3px 9px", fontSize:12, fontWeight:700, cursor:"pointer" }}>✕</button>
                      </div>
                    </div>
                    <div style={{ fontWeight:600, color:C.purple, fontSize:14, marginBottom:6 }}>{c.doc_type||"Unknown Document"}</div>
                    <div style={{ display:"flex", gap:16, fontSize:12, color:C.sub, flexWrap:"wrap" }}>
                      {c.issued_date  && <span>Issued: {c.issued_date}</span>}
                      {c.expiry_date  && <span>Expires: {c.expiry_date}</span>}
                      {c.doc_number   && <span>Doc #: {c.doc_number}</span>}
                      {c.issuing_body && <span>By: {c.issuing_body}</span>}
                    </div>
                  </div>
                );
              })
            }
          </div>
        )}

        {/* COMPANIES */}
        {tab==="companies" && (
          <div>
            <div style={{ fontWeight:700, fontSize:18, marginBottom:20 }}>Company Roster</div>
            {companies.length===0
              ? <div style={{ color:C.muted, textAlign:"center", padding:"60px 0" }}>No companies yet. Scan a document to add one.</div>
              : companies.map(name => {
                const cDocs = tradeDocs.filter(c=>c.company_name===name);
                const cExpired = cDocs.filter(c=>{ const d=daysUntilExpiry(c.expiry_date); return d!==null&&d<0; });
                const cIssues = cDocs.filter(c=>cglComplianceFlags(c).length>0);
                const hasIssues = cExpired.length>0||cIssues.length>0;
                const presentTypes = new Set(cDocs.map(c=>c.doc_type));
                const missingTypes = TRADE_DOC_TYPES.filter(t=>!presentTypes.has(t));
                return (
                  <div key={name} onClick={()=>setSelectedCompany(name)} style={{ background:C.card, border:`1px solid ${hasIssues?C.red+"55":C.border}`, borderRadius:12, padding:"16px 20px", marginBottom:10, cursor:"pointer" }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8, marginBottom:8 }}>
                      <span style={{ fontWeight:800, fontSize:15 }}>🏢 {name}</span>
                      <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                        <Badge color={C.purple}>{cDocs.length} doc{cDocs.length!==1?"s":""}</Badge>
                        {cExpired.length>0 && <Badge color={C.red}>{cExpired.length} expired</Badge>}
                        {cIssues.length>0 && <Badge color={C.red}>{cIssues.length} compliance issue{cIssues.length>1?"s":""}</Badge>}
                      </div>
                    </div>
                    {missingTypes.length>0 && (
                      <div style={{ fontSize:11, color:C.yellow, marginBottom:6 }}>Missing: {missingTypes.join(", ")}</div>
                    )}
                    <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                      {cDocs.map(c => {
                        const s = expiryStatus(c.expiry_date);
                        return <Badge key={c.id} color={s.color}>{c.doc_type?.split(" ")[0]||"Doc"}</Badge>;
                      })}
                    </div>
                    <div style={{ marginTop:7, color:C.muted, fontSize:11 }}>Tap to view full document list →</div>
                  </div>
                );
              })
            }
          </div>
        )}

        {/* COMPLIANCE */}
        {tab==="compliance" && (
          <div>
            <div style={{ fontWeight:700, fontSize:18, marginBottom:6 }}>CGL Compliance Issues</div>
            <div style={{ color:C.muted, fontSize:13, marginBottom:22 }}>Certificates that don't meet the $5,000,000 limit, certificate holder, or additional insured requirements</div>
            {complianceIssues.length===0
              ? <div style={{ background:"#052e16", border:`1px solid ${C.green}44`, borderRadius:14, padding:"32px 24px", textAlign:"center" }}>
                  <div style={{ fontSize:32, marginBottom:10 }}>✅</div>
                  <div style={{ fontWeight:700, fontSize:16, color:C.green }}>All CGL certificates on file are compliant</div>
                </div>
              : complianceIssues.map(c => {
                const flags = cglComplianceFlags(c);
                return (
                  <div key={c.id} onClick={()=>setSelectedCompany(c.company_name)} style={{ background:C.card, border:`1px solid ${C.red}55`, borderRadius:11, padding:"14px 18px", marginBottom:9, cursor:"pointer" }}>
                    <div style={{ display:"flex", justifyContent:"space-between", flexWrap:"wrap", gap:8, marginBottom:8 }}>
                      <span style={{ fontWeight:800 }}>🏢 {c.company_name||"Unknown"}</span>
                      <Badge color={C.red}>{flags.length} issue{flags.length>1?"s":""}</Badge>
                    </div>
                    {flags.map((f,i) => <div key={i} style={{ color:"#fca5a5", fontSize:12, fontWeight:600, marginBottom:3 }}>⚠ {f}</div>)}
                  </div>
                );
              })
            }
          </div>
        )}

        {/* EXPIRING */}
        {tab==="expiring" && (
          <div>
            <div style={{ fontWeight:700, fontSize:18, marginBottom:6 }}>Expiry Alerts</div>
            <div style={{ color:C.muted, fontSize:13, marginBottom:22 }}>Documents expired or expiring within 60 days</div>
            {expired.length===0 && critical.length===0 && warning.length===0
              ? <div style={{ background:"#052e16", border:`1px solid ${C.green}44`, borderRadius:14, padding:"32px 24px", textAlign:"center" }}>
                  <div style={{ fontSize:32, marginBottom:10 }}>✅</div>
                  <div style={{ fontWeight:700, fontSize:16, color:C.green }}>All documents are current</div>
                  <div style={{ color:C.muted, fontSize:13, marginTop:6 }}>No expirations within 60 days</div>
                </div>
              : <>
                {[
                  { group: expired,  label:"🚨 Expired", color: C.red    },
                  { group: critical, label:"⚠ Expiring Within 30 Days", color: C.red    },
                  { group: warning,  label:"⏰ Expiring in 31–60 Days",  color: C.yellow },
                ].map(({ group, label, color }) => group.length > 0 && (
                  <div key={label} style={{ marginBottom:28 }}>
                    <div style={{ fontWeight:700, fontSize:14, color, marginBottom:12, display:"flex", alignItems:"center", gap:8 }}>
                      {label} <Badge color={color}>{group.length}</Badge>
                    </div>
                    {group.map(c => {
                      const s = expiryStatus(c.expiry_date);
                      return (
                        <div key={c.id} onClick={()=>setSelectedCompany(c.company_name)} style={{ background:C.card, border:`1px solid ${color}44`, borderRadius:11, padding:"14px 18px", marginBottom:9, cursor:"pointer" }}>
                          <div style={{ display:"flex", justifyContent:"space-between", flexWrap:"wrap", gap:8, marginBottom:6 }}>
                            <span style={{ fontWeight:800 }}>🏢 {c.company_name||"Unknown"}</span>
                            <Badge color={color}>{s.label}</Badge>
                          </div>
                          <div style={{ color:C.purple, fontWeight:600, fontSize:13, marginBottom:4 }}>{c.doc_type}</div>
                          <div style={{ fontSize:12, color:C.sub }}>
                            {c.expiry_date && <span>Expires: {c.expiry_date}</span>}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                ))}
              </>
            }
          </div>
        )}
      </div>

      {/* Manual entry modal */}
      {manualOpen && (
        <div style={{ position:"fixed", inset:0, background:"#000b", zIndex:100, display:"flex", alignItems:"center", justifyContent:"center" }}
          onClick={e=>e.target===e.currentTarget&&setManualOpen(false)}>
          <div style={{ background:C.card, border:`1px solid ${C.border}`, borderRadius:16, padding:30, width:"92%", maxWidth:500, maxHeight:"90vh", overflowY:"auto" }}>
            <div style={{ fontWeight:800, fontSize:18, marginBottom:20 }}>✏️ Add Trade Document Manually</div>
            {INPUT("company_name","Company / Trade Name *")}
            {INPUT("doc_type","Document Type *","text",TRADE_DOC_TYPES)}
            {INPUT("doc_number","Document / Policy Number")}
            {INPUT("issuing_body","Issuing Body / Insurer")}
            {INPUT("issued_date","Issue Date","date")}
            {INPUT("expiry_date","Expiry Date","date")}
            {manual.doc_type === "CGL Insurance Certificate" && <>
              {INPUT("per_occurrence_limit","Per-Occurrence Limit ($)","number")}
              {INPUT("certificate_holder","Certificate Holder")}
              {INPUT("additional_insured","Additional Insured (comma-separated)")}
            </>}
            {INPUT("notes","Notes")}
            {manual.expiry_date && (() => {
              const s = expiryStatus(manual.expiry_date);
              if (s.level==="none") return null;
              return <div style={{ background:s.color+"15", border:`1px solid ${s.color}44`, borderRadius:8, padding:"9px 14px", marginBottom:12, fontSize:13, color:s.color, fontWeight:600 }}>{s.label}</div>;
            })()}
            <div style={{ display:"flex", gap:10, marginTop:18 }}>
              <button onClick={addManual} style={{ background:C.purple, color:"#fff", border:"none", borderRadius:9, padding:"11px 0", fontWeight:800, cursor:"pointer", flex:1 }}>Add Document</button>
              <button onClick={()=>setManualOpen(false)} style={{ background:C.bg, color:C.muted, border:`1px solid ${C.border}`, borderRadius:9, padding:"11px 18px", fontWeight:700, cursor:"pointer" }}>Cancel</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// CONCRETE MODULE (full v2 app, wrapped)
// ═══════════════════════════════════════════════════════════════════════════════
function ConcreteModule({ onBack }) {
  const [tickets, setTickets]   = useState([]);
  const [invoices, setInvoices] = useState([]);
  const [tab, setTab]           = useState("dashboard");
  const [loading, setLoading]   = useState(false);
  const [loadMsg, setLoadMsg]   = useState("");
  const [toast, setToast]       = useState(null);
  const [manualOpen, setManualOpen] = useState(false);
  const [drag, setDrag]         = useState(false);
  const [invDrag, setInvDrag]   = useState(false);
  const [selectedInvoice, setSelectedInvoice] = useState(null);
  const [selectedTicket,  setSelectedTicket]  = useState(null);
  const [ratePerM3, setRatePerM3] = useState("");
  const [manual, setManual] = useState({ date:"",ticket_number:"",supplier:"",mix_design:"",volume_m3:"",volume_yd3:"",area:"",item:"",invoice_number:"",notes:"" });
  const [reviewQueue, setReviewQueue] = useState([]); // tickets pending area/element confirmation
  const [tests, setTests] = useState([]);
  const [storageReady, setStorageReady] = useState(false);
  const [saveStatus, setSaveStatus] = useState("loading");
  const [ticketSearch, setTicketSearch] = useState("");
  const [invoiceSearch, setInvoiceSearch] = useState("");
  const [deletedRecords, setDeletedRecords] = useState([]);
  const [deletedOpen, setDeletedOpen] = useState(false);
  const [feedbackOpen, setFeedbackOpen] = useState(false);
  const skipInitialSaveRef = useRef(true);
  const fileRef    = useRef();
  const invFileRef = useRef();

  useEffect(() => {
    let cancelled=false;
    async function loadConcreteData(){
      for(let attempt=0;attempt<3;attempt++){
        const saved=await storageGet("concrete-data");
        if(saved!==null){
          if(cancelled)return;
          const migrated = migrateConcreteTicketsV17(saved?.tickets || []);
          setTickets(migrated.tickets);
          if(saved?.invoices) setInvoices(saved.invoices);
          if(saved?.tests)    setTests(saved.tests);
          if(saved?.deletedRecords) setDeletedRecords(saved.deletedRecords);
          if(migrated.changed){
            const migrationSaved = await storageSet("concrete-data", { ...saved, tickets:migrated.tickets });
            if(!migrationSaved) console.error("v17 ticket-location migration could not be saved");
          }
          skipInitialSaveRef.current=true;
          setStorageReady(true);
          setSaveStatus("saved");
          return;
        }
        await new Promise(resolve=>setTimeout(resolve,500*(attempt+1)));
      }
      if(!cancelled){
        setSaveStatus("error");
        showToast("Could not load saved data. Nothing was overwritten—please refresh and try again.","err");
      }
    }
    loadConcreteData();
    return()=>{cancelled=true;};
  }, []);

  useEffect(() => {
    if (!storageReady) return;
    if(skipInitialSaveRef.current){
      skipInitialSaveRef.current=false;
      return;
    }
    setSaveStatus("saving");
    const timer=setTimeout(async()=>{
      const ok=await storageSet("concrete-data", { tickets, invoices, tests, deletedRecords });
      setSaveStatus(ok?"saved":"error");
      if (!ok) showToast("Could not save after 3 attempts. Keep this page open and try the action again.", "err");
    },500);
    return()=>clearTimeout(timer);
  }, [tickets, invoices, tests, deletedRecords, storageReady]);

  function showToast(msg, type="ok") { setToast({msg,type}); setTimeout(()=>setToast(null),3500); }
  function deleteTicket(ticket){
    setDeletedRecords(prev=>[{kind:"ticket",record:ticket,deleted_at:new Date().toISOString()},...prev].slice(0,25));
    setTickets(prev=>prev.filter(x=>x.id!==ticket.id));
    showToast(`Ticket #${ticket.ticket_number||"—"} deleted — it can be restored from Recently Deleted.`);
  }
  function deleteInvoice(invoice){
    setDeletedRecords(prev=>[{kind:"invoice",record:invoice,deleted_at:new Date().toISOString()},...prev].slice(0,25));
    setInvoices(prev=>prev.filter(x=>x.id!==invoice.id));
    showToast(`Invoice ${invoice.invoice_number||"—"} deleted — it can be restored from Recently Deleted.`);
  }
  function undoDelete(entry){
    if(entry.kind==="ticket"){
      const key=ticketNumberKey(entry.record.ticket_number);
      if(key&&tickets.some(t=>ticketNumberKey(t.ticket_number)===key)){
        showToast(`Ticket #${entry.record.ticket_number} is already in the log and cannot be restored twice.`,"err");
        return;
      }
      setTickets(prev=>[...prev,entry.record]);
    }else{
      setInvoices(prev=>[...prev,entry.record]);
    }
    setDeletedRecords(prev=>prev.filter(x=>x!==entry));
    showToast(`${entry.kind==="ticket"?"Ticket":"Invoice"} restored ✓`);
  }
  async function copyFeedbackTemplate(){
    const template=`Fortuna Tracker Feedback — v${APP_VERSION}

Date/time:
Your name:
Page/tab:
What were you trying to do?
What did you expect to happen?
What actually happened?
Ticket or invoice number (if applicable):
Browser/device:
Screenshot attached: Yes / No`;
    try{
      await navigator.clipboard.writeText(template);
      showToast("Feedback template copied ✓");
    }catch{
      showToast("Could not copy automatically—please select the template manually.","err");
    }
  }
  async function toB64(file) { return new Promise((res,rej)=>{ const r=new FileReader(); r.onload=()=>res(r.result.split(",")[1]); r.onerror=rej; r.readAsDataURL(file); }); }
  async function toDataURL(file) { return new Promise((res,rej)=>{ const r=new FileReader(); r.onload=()=>res(r.result); r.onerror=rej; r.readAsDataURL(file); }); }

  async function extractTicket(file) {
    const b64 = await toB64(file);
    const isPDF = file.type==="application/pdf";
    const block = isPDF ? {type:"document",source:{type:"base64",media_type:"application/pdf",data:b64}} : {type:"image",source:{type:"base64",media_type:file.type,data:b64}};
    const mpaRef = SCOPE.filter(r=>r.item).map(r=>`${r.area} ${r.item}: ${r.mpa}`).join(", ");
    const prompt = `You are a construction data extraction assistant. This file may contain ONE or MULTIPLE concrete delivery dockets AND separate pumping/equipment slips. Extract EVERY concrete ticket and EVERY pumping slip as its own record. Do not ignore the last page just because it is a different form.
Project areas: ${AREAS.join(", ")}. Element types: ${ITEMS.join(", ")}.

CRITICAL FIELD EXTRACTION RULES — read carefully:

1. ticket_number: For a concrete delivery ticket, read ONLY from the box explicitly labelled "TICKET NO", "NO. BILLET", or "TICKET NO / NO. BILLET". For a separate "EXTRA WORK ORDERS / HOURLY EQUIPMENT RENTALS" pumping form, use the number labelled "Slip No." as ticket_number. Do NOT use ORDER NO, PO number, customer number, or unit number.

2. mix_design (MOST IMPORTANT): You must find the concrete STRENGTH in MPa. Look for ANY of these:
   - A field labelled "PRODUCT CODE" or "CODE DU PRODUIT" — will contain codes like "Q35NA1A", "Q25NB1A", "HRWR10"
   - A field labelled "MIX DESIGN" or "DESIGNATION"
   - Text anywhere on the ticket showing a number followed by "MPA" or "MPa" such as "35MPA N 20MM", "25 MPA", "32MPa"
   - On Quality Concrete tickets, look in the middle section of the ticket for the product/mix code row
   - The strength will be a number like 25, 32, or 35 followed by MPa
   - DO NOT use "ULTRA SLUMP", slump values in mm, or admixture descriptions as the mix_design
   - If you find a product code like "Q35NA1A" extract it AND note the 35 MPa strength
   - Return the full mix description you find, e.g. "35 MPa N 20mm" or "Q35NA1A — 35 MPa"
   - If the product code is partially visible but you can see a number (e.g. "Q35" or "35N"), return what you can see like "Q35NA1A" or "35 MPa"
   - If truly unreadable, return null. NEVER guess the strength; the user must review it.
   - NEVER return "ULTRA SLUMP" or slump/admixture descriptions as the mix_design

3. volume_m3: Read from "QUANTITY" or "QUANTITE" field — a number like 8.00, 7.50 in m³. Do NOT use yd³ values here.

4. date: The delivery/load date on the ticket in YYYY-MM-DD format. Read the printed dispatch date exactly. This project is active in 2026. For example, printed 2026/07/20 or handwritten 20/07/26 means 2026-07-20 — never 2020-07-26. Do not swap the day with digits from the year.

4A. location: Treat the printed WORK TYPE as authoritative when it clearly
names a project location. In particular, "CRANE BASE" means area "Crane Base"
and item "Interior foundations"; "MUD SLAB" means area "Mud Slabs" and item "Slabs".

5. pumping: Look for a line item labelled "Pumping", "Pump", or "Pompage" on a delivery ticket AND look for a separate yellow or white "EXTRA WORK ORDERS / HOURLY EQUIPMENT RENTALS" form where TYPE OF EQUIPMENT says Pump. A pumping form is a valid record even though it has no concrete mix design. Extract:
   - pump_volume_m3: the volume pumped in m³ (e.g. 8.00)
   - pump_cost: the dollar amount charged for pumping (e.g. 450.00, as a number without $ sign)
   - pump_hours_worked: the number under TIME WORKED
   - pump_travel_hours: the number under TRAVEL TIME
   - pump_hours_charged: the number under TOTAL HOURS CHARGED
   - For a pumping form, put the pump unit number in truck_number and the operator/laborer in driver.
   - For a pumping form, set volume_m3 and volume_yd3 to null. The number under MATERIALS / QUANTITY USED is pump_volume_m3 and must not be counted again as delivered concrete.
   - pump_category: choose exactly one of: ${PUMP_CATEGORIES.join(", ")}. Use the form's description of work. "Mud slab" means "Mud Slabs"; do not also classify it as "Slab on Grade".
   - Use the description of work to select area/item. "Mud slab" should be area "Mud Slabs" and item "Slabs".
   If no pumping information exists on this record, return null for all pumping fields.

Return ONLY a valid JSON array (even if only one ticket). No markdown, no explanation:
[{"date":"YYYY-MM-DD","ticket_number":"ticket number or pumping Slip No.","supplier":"supplier name","mix_design":"MPa strength and mix code, or null for pumping slip","volume_m3":number or null,"volume_yd3":number or null,"pump_volume_m3":number or null,"pump_cost":number or null,"pump_hours_worked":number or null,"pump_travel_hours":number or null,"pump_hours_charged":number or null,"pump_category":"one exact pumping budget category or null","area":"best match from area list or null","item":"best match from element list or null","invoice_number":"string or null","driver":"driver or pump operator","truck_number":"truck or pump unit number","notes":"string or null"}]`;
    // Multi-page ticket PDFs can contain many records. A 4,000-token response
    // limit can cut the JSON array off mid-record, which makes it impossible to
    // parse even though Claude read the PDF successfully.
    const res = await fetch("/api/claude",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:16000,messages:[{role:"user",content:[block,{type:"text",text:prompt}]}]})});
    const data = await res.json();
    if(data.error) throw new Error("API error: "+(data.error.message||JSON.stringify(data.error)));
    const text = data.content?.map(b=>b.text||"").join("")||"";
    if(!text) throw new Error("Empty API response (HTTP "+res.status+")");
    if(data.stop_reason === "max_tokens") {
      throw new Error("The PDF contains too many tickets to finish reading in one response. Split it into two smaller PDFs and upload each one.");
    }
    const parsed = extractJSON(text);
    return Array.isArray(parsed) ? parsed : [parsed];
  }

  async function extractInvoice(file) {
    const b64 = await toB64(file);
    const isPDF = file.type==="application/pdf";
    const block = isPDF ? {type:"document",source:{type:"base64",media_type:"application/pdf",data:b64}} : {type:"image",source:{type:"base64",media_type:file.type,data:b64}};
    const prompt = `You are a construction accounts assistant. Extract ALL information from this concrete supplier invoice.
Return ONLY valid JSON (no markdown):
{"invoice_number":"string","invoice_date":"YYYY-MM-DD","supplier":"name","total_amount":number or null,"currency":"CAD/USD/AUD","ticket_numbers":["array"],"total_volume_m3":number or null,"total_volume_yd3":number or null,"line_items":[{"description":"string","quantity":number or null,"unit":"string","unit_price":number or null,"amount":number or null}],"notes":"string or null"}`;
    const res = await fetch("/api/claude",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:1500,messages:[{role:"user",content:[block,{type:"text",text:prompt}]}]})});
    const data = await res.json();
    if(data.error) throw new Error("API error: "+(data.error.message||JSON.stringify(data.error)));
    const text = data.content?.map(b=>b.text||"").join("")||"";
    if(!text) throw new Error("Empty API response (HTTP "+res.status+")");
    return extractJSON(text);
  }

  function suggestLocation(ticketMpa, currentTickets) {
    const ticketMpaNum = parseMpaNum(ticketMpa);
    // Build poured map from ALL tickets including pending batch
    const poured = {};
    currentTickets.forEach(t => {
      if (!t.area || !t.item) return;
      const key = `${t.area}|||${t.item}`;
      poured[key] = (poured[key] || 0) + (parseFloat(t.volume_m3) || 0);
    });
    // Walk SCOPE in order — first line where MPa matches AND still has capacity
    for (const row of SCOPE) {
      if (!row.area || !row.item || !row.mpa) continue;
      const specNum = parseMpaNum(row.mpa);
      if (!ticketMpaNum || !specNum) continue;
      if (specNum !== ticketMpaNum) continue;
      const key = `${row.area}|||${row.item}`;
      const alreadyPoured = poured[key] || 0;
      const remaining = row.m3 - alreadyPoured;
      if (remaining > 0.01) {
        return { area: row.area, item: row.item, specMpa: row.mpa };
      }
    }
    // Fallback — first MPa match regardless of remaining
    for (const row of SCOPE) {
      if (!row.area || !row.item || !row.mpa) continue;
      const specNum = parseMpaNum(row.mpa);
      if (ticketMpaNum && specNum && specNum === ticketMpaNum) {
        return { area: row.area, item: row.item, specMpa: row.mpa };
      }
    }
    return { area: "", item: "", specMpa: null };
  }

  function matchInvoiceToTickets(invoice, allTickets) {
    const invoiceTicketNums = (invoice.ticket_numbers||[]).map(n=>String(n).trim().toLowerCase());
    const explicitlyMatched = allTickets.filter(t=>{
      const tNum=String(t.ticket_number||"").trim().toLowerCase();
      const tInv=String(t.invoice_number||"").trim().toLowerCase();
      const invNum=String(invoice.invoice_number||"").trim().toLowerCase();
      return (invoiceTicketNums.includes(tNum)&&tNum!=="")||(tInv!==""&&invNum!==""&&tInv===invNum);
    });

    const supplierText=String(invoice.supplier||"").toLowerCase();
    const isOcean=supplierText.includes("ocean");
    const invoiceDate=String(invoice.invoice_date||"");
    const invoiceVolume=parseFloat(invoice.total_volume_m3)||0;
    // Ocean invoices show the first/reference ticket for each pour rather than
    // every delivery docket. When the complete same-day Ocean batch reconciles
    // to the invoice total, use that batch. Exact ticket matching remains the
    // default for Quality Concrete and all other suppliers.
    const oceanDateBatch=isOcean&&invoiceDate
      ? allTickets.filter(t=>String(t.date||"")===invoiceDate&&String(t.supplier||"").toLowerCase().includes("ocean")&&(parseFloat(t.volume_m3)||0)>0)
      : [];
    const oceanBatchVolume=oceanDateBatch.reduce((s,t)=>s+(parseFloat(t.volume_m3)||0),0);
    const usesConsolidatedBatch=isOcean&&invoiceVolume>0&&oceanDateBatch.length>0&&Math.abs(oceanBatchVolume-invoiceVolume)<0.5;
    const ticketsOnInvoice=usesConsolidatedBatch?oceanDateBatch:explicitlyMatched;
    const matched=ticketsOnInvoice.map(t=>({invoiceRef:usesConsolidatedBatch?"Ocean date batch":t.ticket_number,ticket:t}));
    const unmatched=usesConsolidatedBatch?[]:invoiceTicketNums.filter(n=>!allTickets.some(t=>String(t.ticket_number||"").trim().toLowerCase()===n));
    const ticketVolume=ticketsOnInvoice.reduce((s,t)=>s+(parseFloat(t.volume_m3)||0),0);
    const volumeMatch=invoiceVolume>0?Math.abs(ticketVolume-invoiceVolume)<0.5:null;
    return {matched,unmatched,ticketsOnInvoice,ticketVolume,invoiceVolume,volumeMatch,usesConsolidatedBatch};
  }

  // ── Upload file to Supabase Storage and return public URL ──
  async function uploadFile(file, folder) {
    try {
      console.log("uploadFile: starting for", file.name, folder);
      const base64 = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result.split(",")[1]);
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
      console.log("uploadFile: base64 ready, length:", base64.length);
      const res = await fetch("/api/file-upload", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ base64, fileName: file.name, mimeType: file.type, folder }),
      });
      console.log("uploadFile: response status:", res.status);
      const data = await res.json();
      console.log("uploadFile: response data:", data);
      return data.url || null;
    } catch(e) {
      console.error("uploadFile error:", e);
      return null;
    }
  }

  async function extractTest(file) {
    const b64 = await toB64(file);
    const isPDF = file.type === "application/pdf";
    const block = isPDF
      ? { type:"document", source:{ type:"base64", media_type:"application/pdf", data:b64 } }
      : { type:"image",    source:{ type:"base64", media_type:file.type,          data:b64 } };
    const prompt = `You are a construction quality control assistant. Extract data from this concrete cylinder break / compressive strength test report.
Project areas: ${AREAS.join(", ")}.

Extract ALL of the following:
1. report_number: lab report or sample ID number
2. date_sampled: date concrete was sampled (YYYY-MM-DD)
3. date_cast: date cylinders were cast (YYYY-MM-DD)
4. pour_area: which project area this test relates to — match to one of: ${AREAS.join(", ")} (or null if unclear)
5. pour_element: element type e.g. "Slab", "Wall", "Column" (or null)
6. mix_design: concrete mix or MPa strength e.g. "35 MPa" or product code
7. supplier: concrete supplier name
8. ticket_number: delivery ticket number if shown
9. slump_mm: slump value in mm (number only)
10. air_content_pct: air content percentage (number only)
11. results: array of break results, each with:
    - age_days: number (7, 14, 28, 56 etc)
    - strength_mpa: compressive strength in MPa (number)
    - break_date: date of break test (YYYY-MM-DD or null)
    - result: "pass" if meets spec, "fail" if below spec, "pending" if not yet tested
12. specified_mpa: the specified design strength in MPa (number)
13. lab_name: testing laboratory name
14. technician: technician name if shown
15. notes: any other relevant notes

Return ONLY valid JSON, no markdown:
{"report_number":"string","date_sampled":"YYYY-MM-DD","date_cast":"YYYY-MM-DD","pour_area":"area or null","pour_element":"element or null","mix_design":"string","supplier":"string or null","ticket_number":"string or null","slump_mm":number or null,"air_content_pct":number or null,"specified_mpa":number or null,"lab_name":"string or null","technician":"string or null","results":[{"age_days":number,"strength_mpa":number or null,"break_date":"YYYY-MM-DD or null","result":"pass|fail|pending"}],"notes":"string or null"}`;
    const res = await fetch("/api/claude", { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ model:"claude-sonnet-4-6", max_tokens:2000, messages:[{ role:"user", content:[block, { type:"text", text:prompt }] }] }) });
    const data = await res.json();
    if (data.error) throw new Error("API error: " + (data.error.message || JSON.stringify(data.error)));
    const text = data.content?.map(b => b.text||"").join("") || "";
    const clean = text.replace(/```json|```/g,"").trim();
    return { ...JSON.parse(clean), id: Date.now() + Math.random(), file_name: file.name };
  }

  async function handleTestFiles(files) {
    if (!files?.length) return;
    setLoading(true);
    for (const file of Array.from(files)) {
      setLoadMsg(`Reading test report "${file.name}"…`);
      try {
        const extracted = await extractTest(file);
        setLoadMsg(`Saving "${file.name}" to storage…`);
        const fileUrl = await uploadFile(file, "tests");
        setTests(prev => [{ ...extracted, file_url: fileUrl }, ...prev]);
        showToast(`Test report ${extracted.report_number || file.name} added ✓`);
      } catch(e) {
        showToast(`Failed to read ${file.name}: ${e.message}`, "err");
      }
    }
    setLoading(false);
    setLoadMsg("");
  }

  async function handleTicketFiles(files) {
    if(!files?.length) return; setLoading(true);
    const pending = [];
    const selectedFiles=Array.from(files);
    const oversized=selectedFiles.filter(file=>file.size>MAX_API_FILE_BYTES);
    const uploadable=selectedFiles.filter(file=>file.size<=MAX_API_FILE_BYTES);
    if(oversized.length){
      showToast(`${oversized.map(file=>file.name).join(", ")} is too large. Use a PDF under 3 MB or split it into smaller files.`,"err");
    }
    if(!uploadable.length){setLoading(false);return;}
    // Refresh first so duplicates uploaded by another user/device are caught.
    const latestSaved = await storageGet("concrete-data");
    const existingTickets = latestSaved?.tickets || tickets;
    const duplicateNumbers = [];
    for(const file of uploadable){
      setLoadMsg(`Reading "${file.name}"…`);
      try{
        const [extractedArr,fileUrl]=await Promise.all([extractTicket(file),uploadFile(file,"tickets")]);
        for(const extracted of extractedArr){
          extracted.date=normalizeTicketDate(extracted.date,file.name);
          const numberKey = ticketNumberKey(extracted.ticket_number);
          const isDuplicate = numberKey && [...existingTickets, ...pending].some(t=>ticketNumberKey(t.ticket_number)===numberKey);
          if(isDuplicate){ duplicateNumbers.push(extracted.ticket_number); continue; }
          if(extracted.volume_m3&&!extracted.volume_yd3) extracted.volume_yd3=+(extracted.volume_m3*M3_TO_YD3).toFixed(3);
          if(extracted.volume_yd3&&!extracted.volume_m3) extracted.volume_m3=+(extracted.volume_yd3/M3_TO_YD3).toFixed(3);
          // Pass growing pending array so each ticket in batch accounts for previous ones
          const allSoFar = [...tickets, ...pending];
          const suggestion = suggestLocation(extracted.mix_design, allSoFar);
          const printedArea = AREAS.find(area=>area.toLowerCase()===String(extracted.area||"").trim().toLowerCase()) || "";
          const printedItem = ITEMS.find(item=>item.toLowerCase()===String(extracted.item||"").trim().toLowerCase()) || "";
          extracted.area = printedArea || suggestion.area || "";
          extracted.item = printedItem || suggestion.item || "";
          pending.push({id:Date.now()+Math.random(),filename:file.name,fileType:file.type,file_url:fileUrl,added_at:new Date().toISOString(),...extracted,_suggested:!printedArea&&!!(suggestion.area)});
        }
      }catch(e){ showToast(`Could not read "${file.name}": ${e.message}`,"err"); }
    }
    setLoading(false); setLoadMsg("");
    if(duplicateNumbers.length > 0){
      showToast(`Duplicate ticket${duplicateNumbers.length>1?"s":""} blocked: ${duplicateNumbers.join(", ")}`,"err");
    }
    if(pending.length > 0){
      setReviewQueue(pending);
      setTab("dashboard");
      showToast(`${pending.length} ticket${pending.length>1?"s":""} scanned — please confirm location below`);
    }
  }

  async function handleInvoiceFiles(files) {
    if(!files?.length) return; setLoading(true); let added=0;
    for(const file of Array.from(files)){
      setLoadMsg(`Reading invoice: "${file.name}"…`);
      try{
        const [extracted,fileUrl]=await Promise.all([extractInvoice(file),uploadFile(file,"invoices")]);
        setInvoices(prev=>[...prev,{id:Date.now()+Math.random(),filename:file.name,fileType:file.type,file_url:fileUrl,added_at:new Date().toISOString(),...extracted}]);
        added++;
      }catch(e){ showToast(`Could not read invoice "${file.name}": ${e.message}`,"err"); }
    }
    setLoading(false); setLoadMsg(""); if(added){showToast(`${added} invoice${added>1?"s":""} scanned ✓`);setTab("invoices");}
  }

  const totalPoured=tickets.reduce((s,t)=>s+(parseFloat(t.volume_m3)||0),0);
  const totalYd3=tickets.reduce((s,t)=>s+(parseFloat(t.volume_yd3)||0),0);
  const totalPumpM3   = tickets.reduce((s,t) => s + (parseFloat(t.pump_volume_m3)||0), 0);
  const totalPumpHours = tickets.reduce((s,t) => s + (parseFloat(t.pump_hours_charged)||0), 0);
  // Pump slips usually record volume/hours but not pricing. Pull pumping charges
  // from uploaded invoice line items, while retaining ticket-entered costs only
  // when that ticket is not already represented by a priced invoice.
  const invoicePumpData = invoices.map(inv => {
    const pumpCost = (inv.line_items||[])
      .filter(line=>/pump/i.test(String(line.description||"")))
      .reduce((s,line)=>s+(parseFloat(line.amount)||0),0);
    const ticketKeys = new Set((inv.ticket_numbers||[]).map(ticketNumberKey).filter(Boolean));
    return { pumpCost, ticketKeys };
  });
  const invoicePumpCost = invoicePumpData.reduce((s,inv)=>s+inv.pumpCost,0);
  const invoicedPumpTicketKeys = new Set(invoicePumpData.filter(inv=>inv.pumpCost>0).flatMap(inv=>[...inv.ticketKeys]));
  const ticketPumpCost = tickets.reduce((s,t)=>{
    const key=ticketNumberKey(t.ticket_number);
    return s+(key&&invoicedPumpTicketKeys.has(key)?0:(parseFloat(t.pump_cost)||0));
  },0);
  const totalPumpCost = invoicePumpCost + ticketPumpCost;
  const pumpRemaining = Math.max(0, TOTAL_PUMP_BUDGET_M3 - totalPumpM3);
  const pumpHoursRemaining = Math.max(0, TOTAL_PUMP_BUDGET_HOURS - totalPumpHours);
  const pumpPct       = TOTAL_PUMP_BUDGET_M3 > 0 ? Math.min(100,(totalPumpM3/TOTAL_PUMP_BUDGET_M3)*100) : 0;
  const mpaMismatches=tickets.filter(t=>checkMpaMismatch(t));
  const pouredMap={};
  tickets.forEach(t=>{ const key=`${t.area||"Unknown"}|||${t.item||""}`; pouredMap[key]=(pouredMap[key]||0)+(parseFloat(t.volume_m3)||0); });
  const scopeProgress=SCOPE.map(r=>{ const poured=pouredMap[`${r.area}|||${r.item}`]||0; return{...r,poured,remaining:Math.max(0,r.m3-poured),overage:Math.max(0,poured-r.m3)}; });
  // Remaining work and overages are kept separate. An overage in one element
  // must never cancel unfinished work in another element.
  const remaining=scopeProgress.reduce((s,r)=>s+r.remaining,0);
  const totalOverage=scopeProgress.reduce((s,r)=>s+r.overage,0);
  const overageElements=scopeProgress.filter(r=>r.overage>0.01);
  const codedPoured=scopeProgress.reduce((s,r)=>s+r.poured,0);
  const pct=(codedPoured/TOTAL_SCOPE_M3)*100;
  const areaTotals={};
  SCOPE.forEach(r=>{ if(!areaTotals[r.area]) areaTotals[r.area]={scope:0,poured:0}; areaTotals[r.area].scope+=r.m3; });
  Object.entries(pouredMap).forEach(([key,val])=>{ const area=key.split("|||")[0]; if(!areaTotals[area]) areaTotals[area]={scope:0,poured:0}; areaTotals[area].poured+=val; });
  const invoicesWithIssues=invoices.filter(inv=>{ const m=matchInvoiceToTickets(inv,tickets); return m.unmatched.length>0||m.volumeMatch===false; }).length;
  const totalInvoiced=invoices.reduce((s,inv)=>s+(parseFloat(inv.total_amount)||0),0);
  const searchableDate=value=>{
    const raw=String(value||"").trim();
    if(!raw) return "";
    const variants=[raw,raw.replace(/-/g,"/")];
    const match=raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if(match){
      const d=new Date(+match[1],+match[2]-1,+match[3],12,0,0);
      variants.push(
        `${match[2]}/${match[3]}/${match[1]}`,
        `${match[3]}/${match[2]}/${match[1]}`,
        d.toLocaleDateString("en-CA",{year:"numeric",month:"long",day:"numeric"}),
        d.toLocaleDateString("en-CA",{month:"short",day:"numeric",year:"numeric"})
      );
    }
    return variants.join(" ").toLowerCase();
  };
  const ticketNeedle=ticketSearch.trim().toLowerCase();
  const invoiceNeedle=invoiceSearch.trim().toLowerCase();
  const filteredTickets=ticketNeedle?tickets.filter(t=>[
    searchableDate(t.date),t.ticket_number,t.supplier,t.mix_design,t.area,t.item,t.invoice_number,t.driver,t.truck_number
  ].some(value=>String(value||"").toLowerCase().includes(ticketNeedle))):tickets;
  const filteredInvoices=invoiceNeedle?invoices.filter(inv=>[
    searchableDate(inv.invoice_date),inv.invoice_number,inv.supplier,inv.total_amount,
    ...(inv.ticket_numbers||[])
  ].some(value=>String(value||"").toLowerCase().includes(invoiceNeedle))):invoices;

  function exportXLSX() {
    const wb=XLSX.utils.book_new();
    const toExcelDate=value=>{
      const match=String(value||"").match(/^(\d{4})-(\d{2})-(\d{2})$/);
      return match?new Date(+match[1],+match[2]-1,+match[3],12,0,0):value||"";
    };
    const ticketRows=tickets.map((t,i)=>{
      const mismatch=checkMpaMismatch(t);
      const ticketKey=String(t.ticket_number||"").trim().toLowerCase();
      const matchedInvoice=invoices.find(inv=>{
        const invoiceTickets=(inv.ticket_numbers||[]).map(n=>String(n).trim().toLowerCase());
        const directMatch=ticketKey!==""&&invoiceTickets.includes(ticketKey);
        const savedInvoiceMatch=String(t.invoice_number||"").trim()!==""&&String(t.invoice_number).trim().toLowerCase()===String(inv.invoice_number||"").trim().toLowerCase();
        return directMatch||savedInvoiceMatch;
      });
      const exportInvoiceNumber=matchedInvoice?.invoice_number||t.invoice_number||"";
      return {"#":i+1,"Date":toExcelDate(t.date),"Ticket #":t.ticket_number||"","Supplier":t.supplier||"","Mix Design (Ticket)":t.mix_design||"","Spec MPa":t.area&&t.item?(MPA_SPEC[`${t.area}|||${t.item}`]||""):"","MPa Status":mismatch?`⚠ MISMATCH (spec: ${mismatch.specMpa})`:t.mix_design?"✓ OK":"—","Area":t.area||"","Element":t.item||"","Volume (m³)":parseFloat(t.volume_m3)||"","Volume (yd³)":parseFloat(t.volume_yd3)||"","Pumped (m³)":parseFloat(t.pump_volume_m3)||"","Pump Hours Charged":parseFloat(t.pump_hours_charged)||"","Invoice #":exportInvoiceNumber,"Driver / Operator":t.driver||"","Truck / Unit #":t.truck_number||"","Notes":t.notes||""};
    });
    const ws1=XLSX.utils.json_to_sheet(ticketRows,{cellDates:true,dateNF:"yyyy-mm-dd"});
    ws1["!cols"]=[4,12,16,22,18,16,20,14,14,14,14,14,14,12,22].map(w=>({wch:w}));
    if(ws1["!ref"]) ws1["!autofilter"]={ref:ws1["!ref"]};
    XLSX.utils.book_append_sheet(wb,ws1,"Ticket Log");
    const ws2=XLSX.utils.json_to_sheet(scopeProgress.map(r=>({"Area":r.area,"Element":r.item,"Spec MPa":r.mpa||"","Scope (m³)":r.m3,"Poured (m³)":r.poured||"","Remaining (m³)":r.remaining||"","Overage (m³)":r.overage||"","Variance (m³)":+(r.poured-r.m3).toFixed(2),"% Complete":r.m3>0?((r.poured/r.m3)*100).toFixed(1)+"%":"0%"})));
    ws2["!cols"]=[14,22,14,14,14,16,14,14,12].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb,ws2,"Progress by Element");
    if(mpaMismatches.length>0){ const wsm=XLSX.utils.json_to_sheet(mpaMismatches.map(t=>({ "Ticket #":t.ticket_number||"","Date":t.date||"","Area":t.area||"","Element":t.item||"","Ticket Mix":t.mix_design||"","Spec MPa":MPA_SPEC[`${t.area}|||${t.item}`]||"","Supplier":t.supplier||"","Volume (m³)":parseFloat(t.volume_m3)||"" }))); wsm["!cols"]=[14,12,14,16,20,16,20,14].map(w=>({wch:w})); XLSX.utils.book_append_sheet(wb,wsm,"⚠ MPa Mismatches"); }
    XLSX.writeFile(wb,`concrete-tracker-${new Date().toISOString().slice(0,10)}.xlsx`);
    showToast("Spreadsheet downloaded ✓");
  }

  async function addManual() {
    if(!manual.ticket_number&&!manual.date){showToast("Enter at least a date or ticket number.","err");return;}
    const codingIssue=ticketCodingIssue(manual);
    if(codingIssue){showToast(`Select a ${codingIssue} before saving this concrete ticket.`,"err");return;}
    const manualMismatch=checkMpaMismatch(manual);
    if(manualMismatch&&!window.confirm(`MPa mismatch: ticket shows ${manualMismatch.ticketMpa}, but ${manual.area} — ${manual.item} requires ${manualMismatch.specMpa}. Save the actual ticket anyway?`)) return;
    const latestSaved = await storageGet("concrete-data");
    const existingTickets = latestSaved?.tickets || tickets;
    const numberKey = ticketNumberKey(manual.ticket_number);
    if(numberKey && existingTickets.some(t=>ticketNumberKey(t.ticket_number)===numberKey)){
      showToast(`Duplicate ticket blocked: ${manual.ticket_number} is already saved.`,"err");
      return;
    }
    let m={...manual};
    if(m.volume_m3&&!m.volume_yd3) m.volume_yd3=+(parseFloat(m.volume_m3)*M3_TO_YD3).toFixed(3);
    if(m.volume_yd3&&!m.volume_m3) m.volume_m3=+(parseFloat(m.volume_yd3)/M3_TO_YD3).toFixed(3);
    const ticket={id:Date.now(),filename:"Manual entry",added_at:new Date().toISOString(),...m};
    setTickets(prev=>[...prev,ticket]);
    const mismatch=checkMpaMismatch(ticket);
    setManual({date:"",ticket_number:"",supplier:"",mix_design:"",volume_m3:"",volume_yd3:"",area:"",item:"",invoice_number:"",notes:""});
    setManualOpen(false);
    if(mismatch) showToast(`Ticket added — ⚠ MPa mismatch! Ticket: ${mismatch.ticketMpa}, Spec: ${mismatch.specMpa}`,"err");
    else showToast("Ticket added ✓");
  }

  async function saveTicketEdits(original, draft) {
    const codingIssue=ticketCodingIssue(draft);
    if(codingIssue){showToast(`Select a ${codingIssue} before saving this concrete ticket.`,"err");return false;}
    const mismatch=checkMpaMismatch(draft);
    if(mismatch&&!window.confirm(`MPa mismatch: ticket shows ${mismatch.ticketMpa}, but ${draft.area} — ${draft.item} requires ${mismatch.specMpa}. Save the actual ticket anyway?`)) return false;
    const latestSaved=await storageGet("concrete-data");
    const latestTickets=latestSaved?.tickets||tickets;
    const newNumberKey=ticketNumberKey(draft.ticket_number);
    const duplicate=newNumberKey&&latestTickets.some(t=>t.id!==original.id&&ticketNumberKey(t.ticket_number)===newNumberKey);
    if(duplicate){showToast(`Ticket number ${draft.ticket_number} is already in the log.`,"err");return false;}
    const volumeM3=parseFloat(draft.volume_m3)||0;
    const updated={...original,...draft,volume_m3:volumeM3||null,volume_yd3:volumeM3?+(volumeM3*M3_TO_YD3).toFixed(3):null,modified_at:new Date().toISOString()};
    const exists=latestTickets.some(t=>t.id===original.id);
    if(!exists){showToast("This ticket changed on another device. Refresh and try again.","err");return false;}
    setTickets(latestTickets.map(t=>t.id===original.id?updated:t));
    setSelectedTicket(updated);
    showToast(`Ticket #${updated.ticket_number||"—"} updated ✓`);
    return true;
  }

  const TAB=(t,label)=>(<button onClick={()=>setTab(t)} style={{padding:"8px 16px",borderRadius:8,cursor:"pointer",fontWeight:700,fontSize:12,border:"none",background:tab===t?C.accent:"transparent",color:tab===t?"#fff":C.muted,transition:"all .15s",whiteSpace:"nowrap"}}>{label}</button>);
  const INPUT=(key,label,type="text",opts=null)=>(<div style={{marginBottom:12}}><label style={{display:"block",color:C.muted,fontSize:10,fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:4}}>{label}</label>{opts?<select value={manual[key]} onChange={e=>setManual(m=>({...m,[key]:e.target.value}))} style={{width:"100%",background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,padding:"9px 12px",color:C.text,fontSize:14,boxSizing:"border-box"}}><option value="">— select —</option>{opts.map(o=><option key={o} value={o}>{o}</option>)}</select>:<input type={type} value={manual[key]} onChange={e=>setManual(m=>({...m,[key]:e.target.value}))} style={{width:"100%",background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,padding:"9px 12px",color:C.text,fontSize:14,boxSizing:"border-box"}}/>}</div>);

  function TicketModal({ticket,onClose,onSave}){
    const [editing,setEditing]=useState(false);
    const [draft,setDraft]=useState({...ticket});
    const shown=editing?draft:ticket;
    const mismatch=checkMpaMismatch(shown);
    const specMpa=shown.area&&shown.item?MPA_SPEC[`${shown.area}|||${shown.item}`]:null;
    const editFieldStyle={width:"100%",background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,padding:"9px 12px",color:C.text,fontSize:14,boxSizing:"border-box"};
    const editField=(label,field,type="text")=><div style={{marginBottom:11}}><label style={{display:"block",color:C.muted,fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>{label}</label><input type={type} value={draft[field]??""} onChange={e=>setDraft(d=>({...d,[field]:e.target.value}))} style={editFieldStyle}/></div>;
    return(<div style={{position:"fixed",inset:0,background:"#000c",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={{background:C.card,border:`1px solid ${mismatch?C.red:C.border}`,borderRadius:16,padding:28,width:"94%",maxWidth:520,maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <div><div style={{fontWeight:800,fontSize:18}}>🧾 Ticket #{shown.ticket_number||"—"}</div><div style={{color:C.muted,fontSize:13}}>{shown.supplier} · {shown.date}</div></div>
          <Badge color={mismatch?C.red:C.green}>{mismatch?"⚠ MPa Mismatch":"✓ OK"}</Badge>
        </div>
        {mismatch&&<div style={{background:"#450a0a",border:`1px solid ${C.red}`,borderRadius:10,padding:"14px 16px",marginBottom:16}}><div style={{color:C.red,fontWeight:800,fontSize:14,marginBottom:6}}>⚠ Mix Design Mismatch — Do Not Pour</div><div style={{color:"#fca5a5",fontSize:13}}>Ticket shows <b>{mismatch.ticketMpa}</b> but this element requires <b>{mismatch.specMpa}</b>.<br/>Verify with supplier before proceeding.</div></div>}
        {editing?<div style={{marginBottom:16}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>{editField("Ticket #","ticket_number")}{editField("Date","date","date")}</div>
          {editField("Supplier","supplier")}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>{editField("Volume (m³)","volume_m3","number")}{editField("Ticket Mix / MPa","mix_design")}</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <div style={{marginBottom:11}}><label style={{display:"block",color:C.muted,fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>Area</label><select value={draft.area||""} onChange={e=>setDraft(d=>({...d,area:e.target.value,item:""}))} style={editFieldStyle}><option value="">— select area —</option>{AREAS.map(a=><option key={a} value={a}>{a}</option>)}</select></div>
            <div style={{marginBottom:11}}><label style={{display:"block",color:C.muted,fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>Element</label><select value={draft.item||""} onChange={e=>setDraft(d=>({...d,item:e.target.value}))} style={editFieldStyle}><option value="">— select element —</option>{ITEMS.map(it=><option key={it} value={it}>{it}</option>)}</select></div>
          </div>
          {editField("Invoice #","invoice_number")}
          <div style={{marginBottom:11}}><label style={{display:"block",color:C.muted,fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>Notes</label><textarea value={draft.notes||""} onChange={e=>setDraft(d=>({...d,notes:e.target.value}))} rows={3} style={{...editFieldStyle,resize:"vertical"}}/></div>
          {specMpa&&<div style={{background:mismatch?"#450a0a":"#052e16",border:`1px solid ${mismatch?C.red:C.green}55`,borderRadius:8,padding:"9px 12px",fontSize:13,color:mismatch?"#fca5a5":"#86efac"}}>Specified mix for {draft.area} — {draft.item}: <b>{specMpa}</b></div>}
        </div>:<>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:16}}>
          {ticket.volume_m3&&<div style={{background:C.bg,borderRadius:10,padding:"12px 16px",flex:1}}><div style={{color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase"}}>Volume</div><div style={{color:C.accent,fontWeight:800,fontSize:20,fontFamily:"monospace"}}>{parseFloat(ticket.volume_m3).toFixed(2)} m³</div></div>}
          {ticket.pump_volume_m3&&<div style={{background:C.bg,borderRadius:10,padding:"12px 16px",flex:1}}><div style={{color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase"}}>Pumped</div><div style={{color:C.teal,fontWeight:800,fontSize:20,fontFamily:"monospace"}}>{parseFloat(ticket.pump_volume_m3).toFixed(2)} m³</div></div>}
          {ticket.pump_hours_charged&&<div style={{background:C.bg,borderRadius:10,padding:"12px 16px",flex:1}}><div style={{color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase"}}>Pump Hours</div><div style={{color:C.blue,fontWeight:800,fontSize:20,fontFamily:"monospace"}}>{parseFloat(ticket.pump_hours_charged)} hrs</div></div>}
          <div style={{background:C.bg,borderRadius:10,padding:"12px 16px",flex:1}}><div style={{color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase"}}>Ticket Mix</div><div style={{color:mismatch?C.red:C.text,fontWeight:800,fontSize:16}}>{ticket.mix_design||"—"}</div></div>
          {specMpa&&<div style={{background:C.bg,borderRadius:10,padding:"12px 16px",flex:1}}><div style={{color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase"}}>Spec MPa</div><div style={{color:C.green,fontWeight:800,fontSize:16}}>{specMpa}</div></div>}
        </div>
        <div style={{background:C.bg,borderRadius:10,padding:"14px 16px",marginBottom:16,fontSize:13}}>
          {ticket.area&&<div style={{marginBottom:6}}><span style={{color:C.muted}}>Location: </span><b>{ticket.area}{ticket.item?` — ${ticket.item}`:""}</b></div>}
          {ticket.supplier&&<div style={{marginBottom:6}}><span style={{color:C.muted}}>Supplier: </span><b>{ticket.supplier}</b></div>}
          {ticket.invoice_number&&<div style={{marginBottom:6}}><span style={{color:C.muted}}>Invoice #: </span><b>{ticket.invoice_number}</b></div>}
          {ticket.truck_number&&<div style={{marginBottom:6}}><span style={{color:C.muted}}>Truck: </span><b>{ticket.truck_number}</b></div>}
          {ticket.driver&&<div style={{marginBottom:6}}><span style={{color:C.muted}}>Driver: </span><b>{ticket.driver}</b></div>}
          {ticket.notes&&<div style={{marginTop:8,color:C.muted,fontStyle:"italic"}}>{ticket.notes}</div>}
        </div>
        </>}
        <div style={{display:"flex",gap:10}}>{editing?<><button onClick={async()=>{if(await onSave(ticket,draft)){setEditing(false);onClose();}}} style={{background:C.accent,color:"#fff",border:"none",borderRadius:9,padding:"10px 20px",fontWeight:800,cursor:"pointer",flex:1}}>Save Changes</button><button onClick={()=>{setDraft({...ticket});setEditing(false);}} style={{background:C.bg,color:C.muted,border:`1px solid ${C.border}`,borderRadius:9,padding:"10px 20px",fontWeight:700,cursor:"pointer"}}>Cancel</button></>:<><button onClick={()=>setEditing(true)} style={{background:C.blue,color:"#fff",border:"none",borderRadius:9,padding:"10px 20px",fontWeight:800,cursor:"pointer",flex:1}}>✏️ Edit Ticket</button><button onClick={onClose} style={{background:C.bg,color:C.muted,border:`1px solid ${C.border}`,borderRadius:9,padding:"10px 20px",fontWeight:700,cursor:"pointer"}}>Close</button></>}</div>
      </div>
    </div>);
  }

  function InvoiceModal({invoice,onClose}){
    const m=matchInvoiceToTickets(invoice,tickets);
    const hasIssues=m.unmatched.length>0||m.volumeMatch===false;
    return(<div style={{position:"fixed",inset:0,background:"#000c",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:16,padding:28,width:"94%",maxWidth:580,maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <div><div style={{fontWeight:800,fontSize:18}}>🧾 Invoice {invoice.invoice_number||"—"}</div><div style={{color:C.muted,fontSize:13}}>{invoice.supplier} · {invoice.invoice_date}</div></div>
          <Badge color={hasIssues?C.red:C.green}>{hasIssues?"⚠ Review":"✓ OK"}</Badge>
        </div>
        <div style={{display:"flex",gap:12,marginBottom:20,flexWrap:"wrap"}}>
          {invoice.total_amount>0&&<div style={{background:C.bg,borderRadius:10,padding:"12px 18px",flex:1}}><div style={{color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:1}}>Total</div><div style={{color:C.green,fontWeight:800,fontSize:20,fontFamily:"monospace"}}>{invoice.currency||""} {invoice.total_amount?.toLocaleString()}</div></div>}
          {invoice.total_volume_m3>0&&<div style={{background:C.bg,borderRadius:10,padding:"12px 18px",flex:1}}><div style={{color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:1}}>Invoice Volume</div><div style={{color:C.accent,fontWeight:800,fontSize:20,fontFamily:"monospace"}}>{fmt(invoice.total_volume_m3)} m³</div></div>}
          {m.ticketVolume>0&&<div style={{background:C.bg,borderRadius:10,padding:"12px 18px",flex:1}}><div style={{color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:1}}>Ticket Volume</div><div style={{color:m.volumeMatch===false?C.red:C.green,fontWeight:800,fontSize:20,fontFamily:"monospace"}}>{fmt(m.ticketVolume)} m³</div></div>}
        </div>
        {m.volumeMatch===false&&<div style={{background:"#450a0a",border:`1px solid ${C.red}`,borderRadius:10,padding:"12px 16px",marginBottom:14,color:"#fca5a5",fontSize:13}}>⚠ Volume mismatch — invoice shows {fmt(m.invoiceVolume)} m³ but matched tickets total {fmt(m.ticketVolume)} m³</div>}
        {m.usesConsolidatedBatch&&<div style={{background:"#082f49",border:`1px solid ${C.blue}`,borderRadius:10,padding:"12px 16px",marginBottom:14,color:"#bae6fd",fontSize:13}}>✓ Ocean consolidated invoice — reconciled against all {m.matched.length} concrete delivery tickets dated {invoice.invoice_date}. Ocean does not print every delivery ticket number on its invoice.</div>}
        {m.unmatched.length>0&&<div style={{background:"#451a03",border:`1px solid ${C.yellow}`,borderRadius:10,padding:"12px 16px",marginBottom:14,color:"#fde68a",fontSize:13}}>⚠ {m.unmatched.length} ticket{m.unmatched.length>1?"s":""} on invoice not in system: <b>{m.unmatched.join(", ")}</b></div>}
        {m.matched.length>0&&<div style={{marginBottom:16}}>
          <div style={{color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:10}}>Matched Tickets ({m.matched.length})</div>
          {m.matched.map(({ticket:t})=>{ const mm=checkMpaMismatch(t); return(<div key={t.id} style={{background:C.bg,borderRadius:9,padding:"10px 14px",marginBottom:8,border:`1px solid ${mm?C.red+"44":"transparent"}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:6}}><div><span style={{fontWeight:700}}>#{t.ticket_number}</span><span style={{color:C.muted,fontSize:12,marginLeft:8}}>{t.date}</span>{t.area&&<span style={{color:C.sub,fontSize:12,marginLeft:8}}>📍 {t.area}{t.item?` — ${t.item}`:""}</span>}</div><div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{t.mix_design&&<Badge color={mm?C.red:C.green}>{t.mix_design}</Badge>}{t.volume_m3&&<Badge color={C.accent}>{parseFloat(t.volume_m3).toFixed(2)} m³</Badge>}{mm&&<Badge color={C.red}>⚠ MPa</Badge>}</div></div></div>); })}
        </div>}
        <button onClick={onClose} style={{background:C.bg,color:C.muted,border:`1px solid ${C.border}`,borderRadius:9,padding:"10px 20px",fontWeight:700,cursor:"pointer",width:"100%"}}>Close</button>
      </div>
    </div>);
  }

  return(
    <div style={{minHeight:"100vh",background:C.bg,color:C.text,fontFamily:"'DM Sans','Segoe UI',sans-serif",paddingBottom:60}}>
      {toast&&<div style={{position:"fixed",top:18,right:18,zIndex:999,background:toast.type==="err"?"#450a0a":"#052e16",color:toast.type==="err"?"#fca5a5":"#86efac",border:`1px solid ${toast.type==="err"?C.red:C.green}`,borderRadius:10,padding:"12px 22px",fontWeight:600,fontSize:14,boxShadow:"0 8px 32px #0009"}}>{toast.msg}</div>}
      {selectedTicket&&<TicketModal ticket={selectedTicket} onClose={()=>setSelectedTicket(null)} onSave={saveTicketEdits}/>} 
      {selectedInvoice&&<InvoiceModal invoice={selectedInvoice} onClose={()=>setSelectedInvoice(null)}/>}
      {feedbackOpen&&<div style={{position:"fixed",inset:0,background:"#000c",zIndex:250,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={e=>e.target===e.currentTarget&&setFeedbackOpen(false)}>
        <div style={{background:C.card,border:`1px solid ${C.blue}66`,borderRadius:16,padding:26,width:"94%",maxWidth:560,maxHeight:"90vh",overflowY:"auto"}}>
          <div style={{fontWeight:800,fontSize:19,marginBottom:6}}>💬 Help Improve the Tracker</div>
          <div style={{color:C.sub,fontSize:13,lineHeight:1.6,marginBottom:16}}>If something looks wrong or could work better, please include the details below and attach a screenshot where possible. Do not delete or re-upload a record just to reproduce an issue.</div>
          <pre style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:10,padding:"14px 16px",whiteSpace:"pre-wrap",fontFamily:"inherit",fontSize:12,color:C.text,lineHeight:1.7,marginBottom:16}}>{`Fortuna Tracker Feedback — v${APP_VERSION}

Date/time:
Your name:
Page/tab:
What were you trying to do?
What did you expect to happen?
What actually happened?
Ticket or invoice number (if applicable):
Browser/device:
Screenshot attached: Yes / No`}</pre>
          <div style={{display:"flex",gap:9}}>
            <button onClick={copyFeedbackTemplate} style={{background:C.blue,color:"#fff",border:"none",borderRadius:9,padding:"10px 18px",fontWeight:800,cursor:"pointer",flex:1}}>📋 Copy Template</button>
            <button onClick={()=>setFeedbackOpen(false)} style={{background:C.bg,color:C.muted,border:`1px solid ${C.border}`,borderRadius:9,padding:"10px 18px",fontWeight:700,cursor:"pointer"}}>Close</button>
          </div>
        </div>
      </div>}
      {deletedOpen&&<div style={{position:"fixed",inset:0,background:"#000c",zIndex:250,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={e=>e.target===e.currentTarget&&setDeletedOpen(false)}>
        <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:16,padding:26,width:"94%",maxWidth:650,maxHeight:"90vh",overflowY:"auto"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:12,marginBottom:18}}>
            <div><div style={{fontWeight:800,fontSize:19}}>🗑 Recently Deleted</div><div style={{color:C.muted,fontSize:12,marginTop:3}}>The 25 most recent deleted tickets and invoices can be restored here.</div></div>
            <button onClick={()=>setDeletedOpen(false)} style={{background:"transparent",color:C.muted,border:"none",fontSize:22,cursor:"pointer"}}>×</button>
          </div>
          {deletedRecords.length===0?<div style={{color:C.muted,textAlign:"center",padding:"36px 0"}}>Nothing has been deleted.</div>:deletedRecords.map((entry,i)=>{
            const r=entry.record||{};
            const label=entry.kind==="ticket"?`Ticket #${r.ticket_number||"—"}`:`Invoice ${r.invoice_number||"—"}`;
            return <div key={`${entry.deleted_at}-${i}`} style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:10,padding:"12px 14px",marginBottom:9,display:"flex",justifyContent:"space-between",alignItems:"center",gap:12,flexWrap:"wrap"}}>
              <div><div style={{fontWeight:750}}>{entry.kind==="ticket"?"🧾":"💰"} {label}</div><div style={{color:C.muted,fontSize:11,marginTop:3}}>{r.date||r.invoice_date||"No record date"} · deleted {new Date(entry.deleted_at).toLocaleString("en-CA")}</div></div>
              <button onClick={()=>undoDelete(entry)} style={{background:C.green+"18",color:C.green,border:`1px solid ${C.green}55`,borderRadius:8,padding:"7px 13px",fontWeight:800,cursor:"pointer"}}>↶ Restore</button>
            </div>;
          })}
        </div>
      </div>}

      <div style={{background:C.card,borderBottom:`1px solid ${C.border}`,padding:"16px 28px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12}}>
        <div style={{display:"flex",alignItems:"center",gap:13}}>
          <button onClick={onBack} style={{background:"transparent",border:`1px solid ${C.border}`,color:C.muted,borderRadius:8,padding:"6px 13px",fontWeight:700,fontSize:12,cursor:"pointer"}}>← Back</button>
          <div style={{width:40,height:40,borderRadius:11,background:C.accent,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20}}>🏗️</div>
          <div>
            <div style={{fontWeight:800,fontSize:17}}>Concrete Tracker <span style={{color:C.muted,fontSize:10,fontWeight:700,verticalAlign:"middle",border:`1px solid ${C.border}`,borderRadius:5,padding:"2px 6px",marginLeft:5}}>v{APP_VERSION}</span></div>
            <div style={{color:C.muted,fontSize:12}}>{fmt(TOTAL_SCOPE_M3,1)} m³ scope · {tickets.length} tickets · {invoices.length} invoices{mpaMismatches.length>0?` · ⚠ ${mpaMismatches.length} MPa mismatch${mpaMismatches.length>1?"es":""}`:""}</div>
          </div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <span style={{fontSize:11,color:saveStatus==="error"?C.red:saveStatus==="saved"?C.green:C.muted,fontWeight:700}}>
            {saveStatus==="saved"?"💾 Saved":saveStatus==="saving"?"⏳ Saving...":saveStatus==="error"?"⚠ Save error":"⏳ Loading..."}
          </span>
          <button onClick={()=>setFeedbackOpen(true)} style={{background:"transparent",color:C.blue,border:`1px solid ${C.blue}55`,borderRadius:9,padding:"9px 13px",fontWeight:750,fontSize:12,cursor:"pointer"}}>💬 Feedback</button>
          <button onClick={()=>setDeletedOpen(true)} style={{background:"transparent",color:C.muted,border:`1px solid ${C.border}`,borderRadius:9,padding:"9px 13px",fontWeight:750,fontSize:12,cursor:"pointer"}}>🗑 Deleted{deletedRecords.length?` (${deletedRecords.length})`:""}</button>
          <button onClick={exportXLSX} style={{background:C.green,color:"#052e16",border:"none",borderRadius:9,padding:"10px 22px",fontWeight:800,fontSize:14,cursor:"pointer"}}>⬇ Export .xlsx</button>
        </div>
      </div>

      {mpaMismatches.length>0&&<div style={{background:"#450a0a",borderBottom:`1px solid ${C.red}`,padding:"10px 28px",display:"flex",alignItems:"center",gap:12}}>
        <span style={{color:C.red,fontWeight:800,fontSize:13}}>⚠ {mpaMismatches.length} ticket{mpaMismatches.length>1?"s":""} with MPa mismatch</span>
        <button onClick={()=>setTab("tickets")} style={{background:C.red+"22",color:C.red,border:`1px solid ${C.red}44`,borderRadius:6,padding:"4px 12px",fontSize:12,fontWeight:700,cursor:"pointer"}}>View Tickets →</button>
      </div>}

      <div style={{padding:"14px 28px 0",display:"flex",gap:4,borderBottom:`1px solid ${C.border}`,overflowX:"auto"}}>
        {TAB("dashboard","📊 Dashboard")}
        {TAB("tickets",`🧾 Tickets (${tickets.length})${mpaMismatches.length>0?" ⚠":""}`)}
        {TAB("invoices",`💰 Invoices (${invoices.length})${invoicesWithIssues>0?` ⚠${invoicesWithIssues}`:""}`)}
        {TAB("pumping","💧 Pumping")}
        {TAB("testing",`🔬 Testing (${tests.length})`)}
        {TAB("remaining","🔮 Remaining Works")}
        {TAB("mpa","🧪 By MPa")}
        {TAB("scope","📋 Full Scope")}
      </div>

      <div style={{padding:"26px 28px"}}>
        {tab==="dashboard"&&(
          <div>
            <div style={{display:"flex",gap:14,flexWrap:"wrap",marginBottom:26}}>
              <Stat label="Total Poured"   value={`${fmt(totalPoured)} m³`}  sub={`${fmt(totalYd3)} yd³`}             color={C.accent}/>
              <Stat label="Remaining"      value={`${fmt(remaining)} m³`}    sub={`${fmt(remaining*M3_TO_YD3)} yd³`}  color={remaining>0?C.yellow:C.green}/>
              <Stat label="Overage"        value={`${fmt(totalOverage)} m³`} sub={overageElements.length?`${overageElements.length} element${overageElements.length>1?"s":""} over scope`:"none recorded"} color={totalOverage>0?C.red:C.green}/>
              <Stat label="Tickets"        value={tickets.length}             sub="dockets scanned"                    color={C.blue}/>
              <Stat label="Pump Used"      value={`${fmt(totalPumpM3)} m³`}  sub={`$${totalPumpCost.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})} charged`} color={pumpPct>90?C.red:C.teal}/>
              <Stat label="MPa Mismatches" value={mpaMismatches.length}       sub={mpaMismatches.length>0?"⚠ review required":"✓ all clear"} color={mpaMismatches.length>0?C.red:C.green}/>
              <Stat label="Invoices"       value={invoices.length}            sub={invoicesWithIssues>0?`⚠ ${invoicesWithIssues} need review`:invoices.length>0?"✓ all matched":"none yet"} color={invoicesWithIssues>0?C.red:C.purple}/>
            </div>
            <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:14,padding:"20px 24px",marginBottom:24}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}><span style={{fontWeight:700}}>Overall Progress</span><span style={{color:C.accent,fontWeight:800,fontFamily:"monospace"}}>{fmt(pct,1)}%</span></div>
              <Bar pct={pct} color={pct>=100?C.green:C.accent}/>
              <div style={{display:"flex",justifyContent:"space-between",marginTop:8,color:C.muted,fontSize:12}}><span>{fmt(codedPoured)} m³ coded to scope</span><span>{fmt(TOTAL_SCOPE_M3,1)} m³ total scope</span></div>
            </div>
            <div style={{display:"flex",gap:14,marginBottom:16,flexWrap:"wrap"}}>
              <div onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);handleTicketFiles(e.dataTransfer.files);}} onClick={()=>fileRef.current.click()} style={{border:`2px dashed ${drag?C.accent:C.border}`,borderRadius:14,padding:"28px 20px",textAlign:"center",cursor:"pointer",flex:1,minWidth:200,background:drag?C.accent+"11":C.card,transition:"all .2s"}}>
                <div style={{fontSize:28,marginBottom:8}}>📎</div>
                <div style={{fontWeight:700,fontSize:14,marginBottom:4}}>Scan Ticket</div>
                <div style={{color:C.muted,fontSize:12}}>Photo or PDF · MPa auto-checked</div>
                <input ref={fileRef} type="file" multiple accept="image/*,application/pdf" style={{display:"none"}} onChange={e=>handleTicketFiles(e.target.files)}/>
              </div>
              <div onDragOver={e=>{e.preventDefault();setInvDrag(true);}} onDragLeave={()=>setInvDrag(false)} onDrop={e=>{e.preventDefault();setInvDrag(false);handleInvoiceFiles(e.dataTransfer.files);}} onClick={()=>invFileRef.current.click()} style={{border:`2px dashed ${invDrag?C.purple:C.border}`,borderRadius:14,padding:"28px 20px",textAlign:"center",cursor:"pointer",flex:1,minWidth:200,background:invDrag?C.purple+"11":C.card,transition:"all .2s"}}>
                <div style={{fontSize:28,marginBottom:8}}>💰</div>
                <div style={{fontWeight:700,fontSize:14,marginBottom:4}}>Scan Invoice</div>
                <div style={{color:C.muted,fontSize:12}}>Auto-matches to tickets</div>
                <input ref={invFileRef} type="file" multiple accept="image/*,application/pdf" style={{display:"none"}} onChange={e=>handleInvoiceFiles(e.target.files)}/>
              </div>
            </div>
            {loading&&<div style={{background:"#1e3a5f",border:`1px solid ${C.blue}`,borderRadius:11,padding:"13px 20px",color:"#93c5fd",fontWeight:600,marginBottom:16}}>⏳ {loadMsg||"Processing…"}</div>}
            <button onClick={()=>setManualOpen(true)} style={{background:C.card,border:`1px solid ${C.border}`,color:C.text,borderRadius:9,padding:"10px 20px",fontWeight:700,cursor:"pointer",fontSize:13}}>✏️ Add Ticket Manually</button>

            {/* REVIEW QUEUE */}
            {reviewQueue.length > 0 && (
              <div style={{marginTop:28}}>
                <div style={{background:"#1a2e1a",border:`2px solid ${C.green}`,borderRadius:16,padding:"20px 24px"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18,flexWrap:"wrap",gap:10}}>
                    <div>
                      <div style={{fontWeight:800,fontSize:17,color:C.green}}>📍 Confirm Ticket Locations</div>
                      <div style={{color:C.muted,fontSize:13,marginTop:3}}>{reviewQueue.length} ticket{reviewQueue.length>1?"s":""} scanned — assign each to the correct area and element before saving</div>
                    </div>
                    <button
                      onClick={async()=>{
                        const uncategorizedPump = reviewQueue.find(t=>parseFloat(t.pump_volume_m3)>0 && !PUMP_CATEGORIES.includes(t.pump_category));
                        if(uncategorizedPump){
                          showToast(`Select a pumping budget category for slip ${uncategorizedPump.ticket_number||"—"} before saving.`,"err");
                          return;
                        }
                        const codingProblem=reviewQueue.find(t=>ticketCodingIssue(t));
                        if(codingProblem){
                          showToast(`Ticket ${codingProblem.ticket_number||"—"} needs a ${ticketCodingIssue(codingProblem)} before it can be saved.`,"err");
                          return;
                        }
                        // Re-check the shared store at the final commit point in
                        // case somebody saved one of these tickets during review.
                        const latestSaved = await storageGet("concrete-data");
                        const latestTickets = latestSaved?.tickets || tickets;
                        const duplicateNumbers = reviewQueue
                          .filter(t=>ticketNumberKey(t.ticket_number) && latestTickets.some(saved=>ticketNumberKey(saved.ticket_number)===ticketNumberKey(t.ticket_number)))
                          .map(t=>t.ticket_number);
                        if(duplicateNumbers.length){
                          showToast(`Save blocked — ticket${duplicateNumbers.length>1?"s":""} already exist${duplicateNumbers.length===1?"s":""}: ${duplicateNumbers.join(", ")}`,"err");
                          return;
                        }
                        const mpaWarnings=[];
                        reviewQueue.forEach(t=>{const mm=checkMpaMismatch(t);if(mm)mpaWarnings.push(t);});
                        if(mpaWarnings.length&&!window.confirm(`${mpaWarnings.length} ticket${mpaWarnings.length>1?"s have":" has"} an MPa mismatch. Save the actual ticket data anyway?`)) return;
                        // Merge with the freshest shared copy so another user's
                        // recently saved tickets are not overwritten.
                        const ticketsToSave=reviewQueue;
                        setTickets([...latestTickets,...ticketsToSave]);
                        setReviewQueue([]);
                        if(mpaWarnings.length>0) showToast(`⚠ ${mpaWarnings.length} MPa mismatch${mpaWarnings.length>1?"es":""} detected!`,"err");
                        else showToast(`${reviewQueue.length} ticket${reviewQueue.length>1?"s":""} saved ✓`);
                        setTab("tickets");
                      }}
                      style={{background:C.green,color:"#052e16",border:"none",borderRadius:9,padding:"10px 22px",fontWeight:800,fontSize:14,cursor:"pointer"}}>
                      ✓ Save All Tickets
                    </button>
                  </div>

                  {reviewQueue.map((t,i)=>{
                    const specMpa = t.area && t.item ? MPA_SPEC[`${t.area}|||${t.item}`] : null;
                    const mismatch = checkMpaMismatch(t);
                    return(
                      <div key={t.id} style={{background:C.card,border:`1px solid ${mismatch?C.red+"66":C.border}`,borderRadius:12,padding:"16px 18px",marginBottom:12}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:8,marginBottom:12}}>
                          <div>
                            <span style={{fontWeight:800,fontSize:15}}>🧾 Ticket #{t.ticket_number||"—"}</span>
                            <span style={{color:C.muted,fontSize:12,marginLeft:10}}>{t.date}</span>
                            <span style={{color:C.muted,fontSize:12,marginLeft:10}}>🏭 {t.supplier}</span>
                          </div>
                          <div style={{display:"flex",gap:7,flexWrap:"wrap",alignItems:"center"}}>
                            {t.volume_m3&&<Badge color={C.accent}>{parseFloat(t.volume_m3).toFixed(2)} m³</Badge>}
                            {t.pump_volume_m3&&<Badge color={C.teal}>{parseFloat(t.pump_volume_m3).toFixed(2)} m³ pumped</Badge>}
                            {t.pump_hours_charged&&<Badge color={C.blue}>{parseFloat(t.pump_hours_charged)} pump hrs</Badge>}
                            {t.mix_design&&<Badge color={mismatch?C.red:specMpa?C.green:C.muted}>{t.mix_design}</Badge>}
                          </div>
                        </div>

                        <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-end"}}>
                          <div style={{flex:1,minWidth:160}}>
                            <label style={{display:"block",color:C.muted,fontSize:10,fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:4}}>
                              Area * {t._suggested&&t.area?<span style={{color:C.blue,fontWeight:600,fontSize:9,letterSpacing:.5}}> 🤖 suggested</span>:null}
                            </label>
                            <select value={t.area||""} onChange={e=>{
                              const val=e.target.value;
                              setReviewQueue(q=>q.map((x,j)=>j===i?{...x,area:val,item:"",_suggested:false}:x));
                            }} style={{width:"100%",background:C.bg,border:`1px solid ${t.area?C.blue:C.yellow}`,borderRadius:8,padding:"9px 12px",color:C.text,fontSize:14,boxSizing:"border-box"}}>
                              <option value="">— select area —</option>
                              {AREAS.map(a=><option key={a} value={a}>{a}</option>)}
                            </select>
                          </div>
                          {parseFloat(t.pump_volume_m3)>0&&<div style={{flex:1,minWidth:190}}>
                            <label style={{display:"block",color:C.muted,fontSize:10,fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:4}}>Pumping Budget Category *</label>
                            <select value={t.pump_category||""} onChange={e=>setReviewQueue(q=>q.map((x,j)=>j===i?{...x,pump_category:e.target.value}:x))} style={{width:"100%",background:C.bg,border:`1px solid ${PUMP_CATEGORIES.includes(t.pump_category)?C.teal:C.yellow}`,borderRadius:8,padding:"9px 12px",color:C.text,fontSize:14,boxSizing:"border-box"}}>
                              <option value="">— select pumping category —</option>
                              {PUMP_CATEGORIES.map(category=><option key={category} value={category}>{category}</option>)}
                            </select>
                          </div>}
                          <div style={{flex:1,minWidth:160}}>
                            <label style={{display:"block",color:C.muted,fontSize:10,fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:4}}>
                              Element * {t._suggested&&t.item?<span style={{color:C.blue,fontWeight:600,fontSize:9,letterSpacing:.5}}> 🤖 suggested</span>:null}
                            </label>
                            <select value={t.item||""} onChange={e=>{
                              const val=e.target.value;
                              setReviewQueue(q=>q.map((x,j)=>j===i?{...x,item:val,_suggested:false}:x));
                            }} style={{width:"100%",background:C.bg,border:`1px solid ${t.item?C.blue:C.yellow}`,borderRadius:8,padding:"9px 12px",color:C.text,fontSize:14,boxSizing:"border-box"}}>
                              <option value="">— select element —</option>
                              {ITEMS.map(it=><option key={it} value={it}>{it}</option>)}
                            </select>
                          </div>
                          <div style={{minWidth:160}}>
                            <label style={{display:"block",color:C.muted,fontSize:10,fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:4}}>Spec MPa</label>
                            <div style={{background:C.bg,border:`1px solid ${mismatch?C.red:specMpa?C.blue+"66":C.border}`,borderRadius:8,padding:"9px 12px",fontSize:14,fontWeight:700,color:mismatch?C.red:specMpa?C.blue:C.muted,minHeight:38,boxSizing:"border-box"}}>
                              {specMpa||"— select area & element —"}
                            </div>
                          </div>
                          {parseFloat(t.volume_m3)>0&&<div style={{flex:1,minWidth:180}}>
                            <label style={{display:"block",color:C.muted,fontSize:10,fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:4}}>Ticket Mix / MPa *</label>
                            <input value={t.mix_design||""} onChange={e=>setReviewQueue(q=>q.map((x,j)=>j===i?{...x,mix_design:e.target.value}:x))} placeholder="e.g. 35 MPa" style={{width:"100%",background:C.bg,border:`1px solid ${parseMpaNum(t.mix_design)?C.blue:C.yellow}`,borderRadius:8,padding:"9px 12px",color:C.text,fontSize:14,boxSizing:"border-box",minHeight:38}}/>
                          </div>}
                        </div>

                        {mismatch&&(
                          <div style={{marginTop:10,background:"#450a0a",border:`1px solid ${C.red}`,borderRadius:8,padding:"9px 14px",fontSize:13,color:"#fca5a5",fontWeight:600}}>
                            ⚠ MPa mismatch — ticket says <b>{mismatch.ticketMpa}</b> but spec requires <b>{mismatch.specMpa}</b> — verify before pouring
                          </div>
                        )}
                        {specMpa&&!mismatch&&(
                          <div style={{marginTop:10,background:"#052e16",border:`1px solid ${C.green}44`,borderRadius:8,padding:"9px 14px",fontSize:13,color:"#86efac",fontWeight:600}}>
                            ✓ Mix design matches spec
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}

        {tab==="tickets"&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:12}}>
              <div><div style={{fontWeight:700,fontSize:18}}>Ticket Log</div>{mpaMismatches.length>0&&<div style={{color:C.red,fontSize:13,marginTop:2}}>⚠ {mpaMismatches.length} ticket{mpaMismatches.length>1?"s":""} with MPa mismatch</div>}</div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
                <div style={{position:"relative"}}>
                  <span style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",color:C.muted,pointerEvents:"none"}}>⌕</span>
                  <input value={ticketSearch} onChange={e=>setTicketSearch(e.target.value)} placeholder="Search date, ticket #, supplier…" style={{width:280,maxWidth:"70vw",background:C.card,border:`1px solid ${C.border}`,borderRadius:9,padding:"9px 34px 9px 34px",color:C.text,fontSize:13,boxSizing:"border-box"}}/>
                  {ticketSearch&&<button onClick={()=>setTicketSearch("")} title="Clear search" style={{position:"absolute",right:9,top:"50%",transform:"translateY(-50%)",background:"transparent",border:"none",color:C.muted,cursor:"pointer",fontSize:16,padding:2}}>×</button>}
                </div>
                <button onClick={()=>fileRef.current.click()} style={{background:C.accent,color:"#fff",border:"none",borderRadius:9,padding:"9px 18px",fontWeight:700,cursor:"pointer",fontSize:13}}>+ Scan Ticket</button>
              </div>
              <input ref={fileRef} type="file" multiple accept="image/*,application/pdf" style={{display:"none"}} onChange={e=>handleTicketFiles(e.target.files)}/>
            </div>
            {tickets.length===0?<div style={{color:C.muted,textAlign:"center",padding:"60px 0"}}>No tickets yet.</div>
            :filteredTickets.length===0?<div style={{color:C.muted,textAlign:"center",padding:"60px 0"}}>No tickets match “{ticketSearch}”.</div>
            :(()=>{
              // Group tickets by date, sorted newest first
              const grouped = {};
              [...filteredTickets].sort((a,b)=>(b.date||"").localeCompare(a.date||"")).forEach(t=>{
                const key = t.date || "No date";
                if(!grouped[key]) grouped[key] = [];
                grouped[key].push(t);
              });
              return Object.entries(grouped).map(([date, dayTickets])=>{
                const dayVol = dayTickets.reduce((s,t)=>s+(parseFloat(t.volume_m3)||0),0);
                const dayMismatches = dayTickets.filter(t=>checkMpaMismatch(t)).length;
                const hasPump = dayTickets.some(t=>parseFloat(t.pump_volume_m3)>0);
                const pumpVol = dayTickets.reduce((s,t)=>s+(parseFloat(t.pump_volume_m3)||0),0);
                // Format date nicely
                let displayDate = date;
                if(date !== "No date") {
                  try {
                    const d = new Date(date + "T12:00:00");
                    displayDate = d.toLocaleDateString("en-CA", { weekday:"long", year:"numeric", month:"long", day:"numeric" });
                  } catch(e) {}
                }
                return (
                  <div key={date} style={{marginBottom:24}}>
                    {/* Date group header */}
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,paddingBottom:8,borderBottom:`2px solid ${C.border}`,flexWrap:"wrap",gap:8}}>
                      <div style={{display:"flex",alignItems:"center",gap:10}}>
                        <span style={{fontWeight:800,fontSize:15,color:C.text}}>📅 {displayDate}</span>
                        {dayMismatches>0&&<Badge color={C.red}>⚠ {dayMismatches} mismatch{dayMismatches>1?"es":""}</Badge>}
                      </div>
                      <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
                        <Badge color={C.accent}>{dayTickets.length} ticket{dayTickets.length>1?"s":""}</Badge>
                        <Badge color={C.blue}>{dayVol.toFixed(2)} m³</Badge>
                        {hasPump&&<Badge color={C.teal}>💧 {pumpVol.toFixed(2)} m³ pumped</Badge>}
                      </div>
                    </div>
                    {/* Tickets for this date */}
                    {dayTickets.map(t=>{
                      const mismatch=checkMpaMismatch(t);
                      const specMpa=t.area&&t.item?MPA_SPEC[`${t.area}|||${t.item}`]:null;
                      return(<div key={t.id} onClick={()=>setSelectedTicket(t)} style={{background:C.card,border:`1px solid ${mismatch?C.red+"88":C.border}`,borderRadius:12,padding:"15px 20px",marginBottom:10,cursor:"pointer"}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8,flexWrap:"wrap",gap:8}}>
                          <div><span style={{fontWeight:800,fontSize:15}}>{t.ticket_number||"No ticket #"}</span></div>
                          <div style={{display:"flex",gap:7,flexWrap:"wrap",alignItems:"center"}}>
                            {t.volume_m3&&<Badge color={C.accent}>{parseFloat(t.volume_m3).toFixed(2)} m³</Badge>}
                            {t.mix_design&&<Badge color={mismatch?C.red:C.green}>{t.mix_design}{mismatch?" ⚠":""}</Badge>}
                            {specMpa&&!mismatch&&<Badge color={C.muted}>spec: {specMpa}</Badge>}
                            {parseFloat(t.pump_volume_m3)>0&&<Badge color={C.teal}>💧 {parseFloat(t.pump_volume_m3).toFixed(2)} m³</Badge>}
                            {(t.file_url||t.originalFile)&&<button onClick={e=>{ e.stopPropagation(); const src=t.file_url||t.originalFile; const isImg=/^data:image|\.(jpg|jpeg|png|gif|webp|heic)/i.test(src); const w=window.open(); w.document.write(isImg?`<html><body style="margin:0;background:#000;display:flex;align-items:center;justify-content:center;min-height:100vh"><img src="${src}" style="max-width:100%;max-height:100vh;object-fit:contain"></body></html>`:`<iframe src="${src}" width="100%" height="100%" style="border:none;position:fixed;top:0;left:0"></iframe>`); }} style={{background:"transparent",border:`1px solid ${C.blue}44`,color:C.blue,borderRadius:6,padding:"3px 9px",fontSize:12,fontWeight:700,cursor:"pointer"}}>📄 View</button>}
                            <button onClick={e=>{ e.stopPropagation(); if(window.confirm(`Delete ticket #${t.ticket_number||"this ticket"}? This will restore its volume to the remaining-work totals. You can undo this from Recently Deleted.`)) deleteTicket(t); }} style={{background:"transparent",border:`1px solid ${C.red}44`,color:C.red,borderRadius:6,padding:"3px 9px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🗑 Delete</button>
                          </div>
                        </div>
                        <div style={{display:"flex",gap:16,flexWrap:"wrap",fontSize:13,color:C.sub}}>
                          {t.supplier&&<span>🏭 {t.supplier}</span>}
                          {t.area&&<span>📍 {t.area}{t.item?` — ${t.item}`:""}</span>}
                          {t.invoice_number&&<span>🧾 Inv: {t.invoice_number}</span>}
                        </div>
                        {mismatch&&<div style={{marginTop:8,background:"#450a0a",borderRadius:7,padding:"7px 12px",fontSize:12,color:"#fca5a5",fontWeight:600}}>⚠ MPa mismatch — ticket: {mismatch.ticketMpa} · spec: {mismatch.specMpa}</div>}
                      </div>);
                    })}
                  </div>
                );
              });
            })()}
          </div>
        )}

        {tab==="invoices"&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:12}}>
              <div><div style={{fontWeight:700,fontSize:18}}>Invoice Matching</div><div style={{color:C.muted,fontSize:13,marginTop:2}}>Auto-matched against logged tickets</div></div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
                <div style={{position:"relative"}}>
                  <span style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",color:C.muted,pointerEvents:"none"}}>⌕</span>
                  <input value={invoiceSearch} onChange={e=>setInvoiceSearch(e.target.value)} placeholder="Search date, invoice #, supplier…" style={{width:280,maxWidth:"70vw",background:C.card,border:`1px solid ${C.border}`,borderRadius:9,padding:"9px 34px 9px 34px",color:C.text,fontSize:13,boxSizing:"border-box"}}/>
                  {invoiceSearch&&<button onClick={()=>setInvoiceSearch("")} title="Clear search" style={{position:"absolute",right:9,top:"50%",transform:"translateY(-50%)",background:"transparent",border:"none",color:C.muted,cursor:"pointer",fontSize:16,padding:2}}>×</button>}
                </div>
                <button onClick={()=>invFileRef.current.click()} style={{background:C.purple,color:"#fff",border:"none",borderRadius:9,padding:"9px 18px",fontWeight:700,cursor:"pointer",fontSize:13}}>+ Scan Invoice</button>
              </div>
              <input ref={invFileRef} type="file" multiple accept="image/*,application/pdf" style={{display:"none"}} onChange={e=>handleInvoiceFiles(e.target.files)}/>
            </div>
            {invoices.length===0?<div style={{color:C.muted,textAlign:"center",padding:"60px 0"}}>No invoices yet.</div>
            :filteredInvoices.length===0?<div style={{color:C.muted,textAlign:"center",padding:"60px 0"}}>No invoices match “{invoiceSearch}”.</div>
            :filteredInvoices.map(inv=>{ const m=matchInvoiceToTickets(inv,tickets); const hasIssues=m.unmatched.length>0||m.volumeMatch===false;
              return(<div key={inv.id} onClick={()=>setSelectedInvoice(inv)} style={{background:C.card,border:`1px solid ${hasIssues?C.red+"66":C.border}`,borderRadius:12,padding:"16px 20px",marginBottom:12,cursor:"pointer"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10,flexWrap:"wrap",gap:8}}>
                  <div><span style={{fontWeight:800,fontSize:15}}>Invoice {inv.invoice_number||"—"}</span><span style={{color:C.muted,fontSize:12,marginLeft:10}}>{inv.invoice_date}</span></div>
                  <div style={{display:"flex",gap:7,flexWrap:"wrap",alignItems:"center"}}>{inv.total_amount>0&&<Badge color={C.green}>{inv.currency||""} {inv.total_amount?.toLocaleString()}</Badge>}<Badge color={hasIssues?C.red:C.green}>{hasIssues?"⚠ Review":"✓ Matched"}</Badge>
                    {(inv.file_url||inv.originalFile)&&<button onClick={e=>{ e.stopPropagation(); const src=inv.file_url||inv.originalFile; const isImg=/^data:image|\.(jpg|jpeg|png|gif|webp|heic)/i.test(src); const w=window.open(); w.document.write(isImg?`<html><body style="margin:0;background:#000;display:flex;align-items:center;justify-content:center;min-height:100vh"><img src="${src}" style="max-width:100%;max-height:100vh;object-fit:contain"></body></html>`:`<iframe src="${src}" width="100%" height="100%" style="border:none;position:fixed;top:0;left:0"></iframe>`); }} style={{background:"transparent",border:`1px solid ${C.blue}44`,color:C.blue,borderRadius:6,padding:"3px 9px",fontSize:12,fontWeight:700,cursor:"pointer"}}>📄 View</button>}
                    <button onClick={e=>{ e.stopPropagation(); if(window.confirm(`Delete invoice ${inv.invoice_number||"this invoice"}? You can undo this from Recently Deleted.`)) deleteInvoice(inv); }} style={{background:"transparent",border:`1px solid ${C.red}44`,color:C.red,borderRadius:6,padding:"3px 9px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🗑 Delete</button>
                  </div>
                </div>
                <div style={{display:"flex",gap:20,fontSize:13,color:C.sub,flexWrap:"wrap"}}>
                  {inv.supplier&&<span>🏭 {inv.supplier}</span>}
                  <span style={{color:m.matched.length>0?C.green:C.muted}}>✓ {m.matched.length} matched</span>
                  {m.unmatched.length>0&&<span style={{color:C.red}}>⚠ {m.unmatched.length} not found</span>}
                  {m.volumeMatch===false&&<span style={{color:C.red}}>⚠ Volume mismatch</span>}
                </div>
              </div>);
            })}
          </div>
        )}

        {tab==="pumping"&&(
          <div>
            <div style={{display:"flex",gap:14,flexWrap:"wrap",marginBottom:24}}>
              <Stat label="Total Pumped"    value={`${fmt(totalPumpM3)} m³`}        sub={`of ${fmt(TOTAL_PUMP_BUDGET_M3)} m³ budgeted`} color={pumpPct>100?C.red:C.teal}/>
              <Stat label="Pump Remaining"  value={`${fmt(pumpRemaining)} m³`}      sub="budget left"                                    color={pumpRemaining<50?C.red:C.yellow}/>
              <Stat label="Pump Hours Used" value={`${fmt(totalPumpHours)} hrs`}    sub={`of ${fmt(TOTAL_PUMP_BUDGET_HOURS)} hrs budgeted`} color={totalPumpHours>TOTAL_PUMP_BUDGET_HOURS?C.red:C.blue}/>
              <Stat label="Hours Remaining" value={`${fmt(pumpHoursRemaining)} hrs`} sub="budget left" color={pumpHoursRemaining<10?C.red:C.yellow}/>
              <Stat label="Pumping Charged" value={`$${totalPumpCost.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}`} sub="before HST · from invoices/tickets" color={C.green}/>
              <Stat label="Tickets w/ Pump" value={tickets.filter(t=>parseFloat(t.pump_volume_m3)>0).length} sub="of total tickets" color={C.blue}/>
            </div>
            <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:14,padding:"20px 24px",marginBottom:24}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}><span style={{fontWeight:700}}>Overall Pump Usage vs Budget</span><span style={{color:pumpPct>90?C.red:C.teal,fontWeight:800,fontFamily:"monospace"}}>{fmt(pumpPct,1)}%</span></div>
              <Bar pct={pumpPct} color={pumpPct>100?C.red:pumpPct>80?C.yellow:C.teal}/>
              <div style={{display:"flex",justifyContent:"space-between",marginTop:8,color:C.muted,fontSize:12}}><span>{fmt(totalPumpM3)} m³ used</span><span>{fmt(TOTAL_PUMP_BUDGET_M3)} m³ budgeted</span></div>
            </div>
            <div style={{fontWeight:800,fontSize:15,marginBottom:12}}>Budget Breakdown by Category</div>
            {PUMP_BUDGET.map(row => {
              const categoryTickets = tickets.filter(t => t.pump_category === row.category);
              const used = categoryTickets.reduce((s,t) => s + (parseFloat(t.pump_volume_m3)||0), 0);
              const hoursUsed = categoryTickets.reduce((s,t) => s + (parseFloat(t.pump_hours_charged)||0), 0);
              const pct2 = row.volume_m3 > 0 ? Math.min(100,(used/row.volume_m3)*100) : 0;
              const over = used > row.volume_m3;
              const hoursOver = hoursUsed > row.hours;
              return (
                <div key={row.category} style={{background:C.card,border:`1px solid ${over?C.red+"66":C.border}`,borderRadius:13,padding:"16px 20px",marginBottom:10}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8,flexWrap:"wrap",gap:8}}>
                    <span style={{fontWeight:700}}>{row.category}</span>
                    <div style={{display:"flex",gap:8,alignItems:"center"}}>
                      {(over||hoursOver) && <Badge color={C.red}>⚠ Over Budget</Badge>}
                      <Badge color={C.teal}>{fmt(used)} / {fmt(row.volume_m3)} m³</Badge>
                      <Badge color={hoursOver?C.red:C.blue}>{fmt(hoursUsed)} / {fmt(row.hours)} hrs</Badge>
                    </div>
                  </div>
                  <Bar pct={pct2} color={over?C.red:pct2>80?C.yellow:C.teal}/>
                  <div style={{display:"flex",justifyContent:"space-between",marginTop:6,color:C.muted,fontSize:12}}>
                    <span>{fmt(used)} m³ used</span>
                    <span style={{color:over?C.red:C.muted}}>{over ? `⚠ ${fmt(used-row.volume_m3)} m³ over` : `${fmt(row.volume_m3-used)} m³ remaining`}</span>
                  </div>
                  <div style={{display:"flex",justifyContent:"space-between",marginTop:4,color:C.muted,fontSize:12}}>
                    <span>{fmt(hoursUsed)} hrs charged</span>
                    <span style={{color:hoursOver?C.red:C.muted}}>{hoursOver ? `⚠ ${fmt(hoursUsed-row.hours)} hrs over` : `${fmt(row.hours-hoursUsed)} hrs remaining`}</span>
                  </div>
                </div>
              );
            })}
            <div style={{marginTop:20,fontWeight:800,fontSize:15,marginBottom:12}}>Tickets with Pumping</div>
            {tickets.filter(t=>parseFloat(t.pump_volume_m3)>0).length===0
              ? <div style={{color:C.muted,textAlign:"center",padding:"40px 0"}}>No tickets with pumping charges yet</div>
              : tickets.filter(t=>parseFloat(t.pump_volume_m3)>0).map(t=>(
                <div key={t.id} style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"14px 18px",marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
                  <div><span style={{fontWeight:700}}>#{t.ticket_number||"—"}</span><span style={{color:C.muted,fontSize:12,marginLeft:10}}>{t.date}</span>{t.area&&<span style={{color:C.sub,fontSize:12,marginLeft:10}}>📍 {t.area}{t.item?` — ${t.item}`:""}</span>}</div>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                    <Badge color={C.teal}>{fmt(parseFloat(t.pump_volume_m3))} m³ pumped</Badge>
                    {t.pump_hours_charged&&<Badge color={C.blue}>{fmt(parseFloat(t.pump_hours_charged))} hrs charged</Badge>}
                    {t.pump_cost&&<Badge color={C.green}>${parseFloat(t.pump_cost).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</Badge>}
                  </div>
                </div>
              ))
            }
          </div>
        )}

        {tab==="testing"&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:12}}>
              <div>
                <div style={{fontWeight:800,fontSize:18}}>Concrete Test Reports</div>
                <div style={{color:C.muted,fontSize:13,marginTop:2}}>{tests.length} report{tests.length!==1?"s":""} uploaded · cylinder break results</div>
              </div>
              <div onDragOver={e=>{e.preventDefault();}} onDrop={e=>{e.preventDefault();handleTestFiles(e.dataTransfer.files);}} onClick={()=>{ const i=document.createElement("input"); i.type="file"; i.multiple=true; i.accept="image/*,application/pdf"; i.onchange=e=>handleTestFiles(e.target.files); i.click(); }} style={{border:`2px dashed ${C.border}`,borderRadius:12,padding:"16px 24px",textAlign:"center",cursor:"pointer",background:C.card}}>
                <div style={{fontSize:22,marginBottom:4}}>🔬</div>
                <div style={{fontWeight:700,fontSize:13}}>Upload Test Reports</div>
                <div style={{color:C.muted,fontSize:11}}>PDF or photo · AI extracts results</div>
              </div>
            </div>
            {tests.length===0
              ? <div style={{color:C.muted,textAlign:"center",padding:"60px 0",fontSize:15}}>No test reports uploaded yet<br/><span style={{fontSize:13}}>Upload lab cylinder break reports to track 7/14/28 day results</span></div>
              : tests.map(test => {
                const allPass = test.results?.every(r=>r.result==="pass"||r.result==="pending");
                const anyFail = test.results?.some(r=>r.result==="fail");
                const latestBreak = test.results?.filter(r=>r.strength_mpa).sort((a,b)=>b.age_days-a.age_days)[0];
                return (
                  <div key={test.id} style={{background:C.card,border:`1px solid ${anyFail?C.red+"66":allPass?C.green+"33":C.border}`,borderRadius:14,padding:"18px 22px",marginBottom:14}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:10,marginBottom:14}}>
                      <div>
                        <div style={{fontWeight:800,fontSize:15}}>Report #{test.report_number||"—"}</div>
                        <div style={{color:C.muted,fontSize:12,marginTop:3}}>
                          Sampled: {test.date_sampled||"—"} · {test.lab_name||"Lab unknown"}
                          {test.pour_area&&<span style={{marginLeft:10,color:C.sub}}>📍 {test.pour_area}{test.pour_element?` — ${test.pour_element}`:""}</span>}
                        </div>
                      </div>
                      <div style={{display:"flex",gap:7,flexWrap:"wrap",alignItems:"center"}}>
                        {test.mix_design&&<Badge color={C.accent}>{test.mix_design}</Badge>}
                        {anyFail?<Badge color={C.red}>⚠ FAIL</Badge>:allPass&&test.results?.some(r=>r.result==="pass")?<Badge color={C.green}>✓ PASS</Badge>:<Badge color={C.yellow}>⏳ Pending</Badge>}
                        {test.file_url&&<button onClick={()=>window.open(test.file_url,"_blank")} style={{background:"transparent",border:`1px solid ${C.blue}44`,color:C.blue,borderRadius:6,padding:"3px 9px",fontSize:12,fontWeight:700,cursor:"pointer"}}>📄 View</button>}
                        <button onClick={()=>setTests(prev=>prev.filter(x=>x.id!==test.id))} style={{background:"transparent",border:`1px solid ${C.red}44`,color:C.red,borderRadius:6,padding:"3px 9px",fontSize:12,fontWeight:700,cursor:"pointer"}}>✕</button>
                      </div>
                    </div>
                    <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:12}}>
                      {test.slump_mm!=null&&<div style={{background:C.bg,borderRadius:8,padding:"8px 14px",fontSize:12}}><span style={{color:C.muted}}>Slump </span><span style={{fontWeight:700}}>{test.slump_mm} mm</span></div>}
                      {test.air_content_pct!=null&&<div style={{background:C.bg,borderRadius:8,padding:"8px 14px",fontSize:12}}><span style={{color:C.muted}}>Air </span><span style={{fontWeight:700}}>{test.air_content_pct}%</span></div>}
                      {test.specified_mpa!=null&&<div style={{background:C.bg,borderRadius:8,padding:"8px 14px",fontSize:12}}><span style={{color:C.muted}}>Spec </span><span style={{fontWeight:700}}>{test.specified_mpa} MPa</span></div>}
                      {latestBreak&&<div style={{background:C.bg,borderRadius:8,padding:"8px 14px",fontSize:12}}><span style={{color:C.muted}}>Latest ({latestBreak.age_days}d) </span><span style={{fontWeight:700,color:latestBreak.result==="fail"?C.red:C.green}}>{latestBreak.strength_mpa} MPa</span></div>}
                    </div>
                    <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                      {(test.results||[]).map((r,i)=>(
                        <div key={i} style={{background:r.result==="fail"?C.red+"22":r.result==="pass"?C.green+"22":C.bg,border:`1px solid ${r.result==="fail"?C.red+"66":r.result==="pass"?C.green+"44":C.border}`,borderRadius:10,padding:"10px 16px",minWidth:90,textAlign:"center"}}>
                          <div style={{color:C.muted,fontSize:11,fontWeight:700}}>{r.age_days} DAY</div>
                          <div style={{fontWeight:800,fontSize:18,color:r.result==="fail"?C.red:r.result==="pass"?C.green:C.muted,margin:"4px 0"}}>{r.strength_mpa!=null?`${r.strength_mpa}`:"—"}<span style={{fontSize:11,fontWeight:400}}> MPa</span></div>
                          <div style={{fontSize:11,color:r.result==="fail"?C.red:r.result==="pass"?C.green:C.muted,fontWeight:700}}>{r.result==="pending"?"PENDING":r.result?.toUpperCase()}</div>
                          {r.break_date&&<div style={{fontSize:10,color:C.muted,marginTop:3}}>{r.break_date}</div>}
                        </div>
                      ))}
                    </div>
                    {test.notes&&<div style={{marginTop:12,color:C.muted,fontSize:12,borderTop:`1px solid ${C.border}`,paddingTop:10}}>📝 {test.notes}</div>}
                  </div>
                );
              })
            }
          </div>
        )}

        {tab==="remaining"&&(()=>{
          const rate=parseFloat(ratePerM3)||null;
          const invoicedVolume=invoices.reduce((s,inv)=>s+(parseFloat(inv.total_volume_m3)||0),0);
          const invoicedAmount=invoices.reduce((s,inv)=>s+(parseFloat(inv.total_amount)||0),0);
          const derivedRate=invoicedVolume>0&&invoicedAmount>0?invoicedAmount/invoicedVolume:null;
          const totalEstCost=rate?remaining*rate:null;
          const areaRows=AREAS.map(area=>{ const lines=scopeProgress.filter(r=>r.area===area); const scope=lines.reduce((s,r)=>s+r.m3,0); const poured=lines.reduce((s,r)=>s+r.poured,0); const hasScope=scope>0; const rem=lines.reduce((s,r)=>s+r.remaining,0); const over=lines.reduce((s,r)=>s+r.overage,0); const p=hasScope?(poured/scope)*100:0; const status=over>0.01?"over":hasScope&&rem<=0.01?"complete":poured>0?"inprogress":"notstarted"; const estCost=rate&&rem>0?rem*rate:null; const areaMpas=[...new Set(Object.entries(MPA_SPEC).filter(([key])=>key.startsWith(`${area}|||`)).map(([,mpa])=>mpa))]; return{area,scope,poured,rem,over,p,status,estCost,areaMpas,hasScope}; });
          const STATUS_CONFIG={over:{label:"⚠️ Over Scope",color:C.red},complete:{label:"✅ Complete",color:C.green},inprogress:{label:"🟡 In Progress",color:C.yellow},notstarted:{label:"🔴 Not Started",color:C.red}};
          return(<div>
            <div style={{fontWeight:700,fontSize:18,marginBottom:6}}>Remaining Works</div>
            <div style={{color:C.muted,fontSize:13,marginBottom:22}}>Live picture of what's done, in progress, and still to pour</div>
            <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:24}}>
              {[{emoji:"⚠️",count:areaRows.filter(r=>r.status==="over").length,label:"Over Scope",color:C.red},{emoji:"✅",count:areaRows.filter(r=>r.status==="complete").length,label:"Complete",color:C.green},{emoji:"🟡",count:areaRows.filter(r=>r.status==="inprogress").length,label:"In Progress",color:C.yellow},{emoji:"🔴",count:areaRows.filter(r=>r.status==="notstarted").length,label:"Not Started",color:C.red},{emoji:"🏗️",count:fmt(remaining,1),label:"m³ left",color:C.accent}].map(({emoji,count,label,color})=>(<div key={label} style={{background:color+"18",border:`1px solid ${color}44`,borderRadius:12,padding:"14px 20px",flex:1,minWidth:110,textAlign:"center"}}><div style={{fontSize:28}}>{emoji}</div><div style={{fontWeight:800,fontSize:22,color}}>{count}</div><div style={{color:C.muted,fontSize:12,marginTop:2}}>{label}</div></div>))}
            </div>
            <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:14,padding:"18px 22px",marginBottom:24}}>
              <div style={{fontWeight:700,marginBottom:12}}>💲 Estimated Remaining Cost</div>
              <div style={{display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
                <div style={{flex:1,minWidth:200}}>
                  <label style={{color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:1,display:"block",marginBottom:5}}>Rate per m³</label>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <input type="number" placeholder="e.g. 220" value={ratePerM3} onChange={e=>setRatePerM3(e.target.value)} style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,padding:"9px 12px",color:C.text,fontSize:14,width:140}}/>
                    {derivedRate&&!ratePerM3&&<button onClick={()=>setRatePerM3(derivedRate.toFixed(2))} style={{background:C.purple+"22",color:C.purple,border:`1px solid ${C.purple}44`,borderRadius:8,padding:"8px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Use invoice rate (${derivedRate.toFixed(2)}/m³)</button>}
                  </div>
                </div>
                {totalEstCost!==null&&<div style={{background:C.bg,borderRadius:10,padding:"12px 20px"}}><div style={{color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:1}}>Est. Remaining Cost</div><div style={{color:C.green,fontWeight:800,fontSize:24,fontFamily:"monospace"}}>${totalEstCost.toLocaleString(undefined,{maximumFractionDigits:0})}</div></div>}
              </div>
            </div>
            {areaRows.map(row=>{ const sc=STATUS_CONFIG[row.status]; return(<div key={row.area} style={{background:C.card,border:`1px solid ${row.status==="over"?C.red+"88":row.status==="complete"?C.green+"44":row.status==="inprogress"?C.yellow+"44":C.border}`,borderRadius:13,padding:"15px 20px",marginBottom:10}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8,flexWrap:"wrap",gap:8}}>
                <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}><span style={{fontWeight:800,fontSize:15}}>📍 {row.area}</span><Badge color={sc.color}>{sc.label}</Badge>{row.areaMpas.map(m=><MpaBadge key={m} mpa={m}/>)}</div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>{row.hasScope?<Badge color={C.accent}>{fmt(row.p,1)}%</Badge>:<Badge color={C.yellow}>Scope Not Set</Badge>}{row.estCost!==null&&<Badge color={C.green}>${row.estCost.toLocaleString(undefined,{maximumFractionDigits:0})} est.</Badge>}</div>
              </div>
              <Bar pct={row.p} color={row.status==="over"?C.red:row.status==="complete"?C.green:row.status==="inprogress"?C.yellow:C.border}/>
              <div style={{display:"flex",gap:20,marginTop:9,fontSize:13,flexWrap:"wrap"}}>
                <span style={{color:C.muted}}>Poured: <b style={{color:C.text}}>{fmt(row.poured)} m³</b></span>
                <span style={{color:C.muted}}>Scope: <b style={{color:row.hasScope?C.text:C.yellow}}>{row.hasScope?`${fmt(row.scope)} m³`:"Not set"}</b></span>
                {!row.hasScope?null:row.over>0?<span style={{color:C.red,fontWeight:800}}>⚠ {fmt(row.over)} m³ over scope</span>:row.rem>0?<span style={{color:C.muted}}>Still to pour: <b style={{color:C.yellow}}>{fmt(row.rem)} m³</b></span>:<span style={{color:C.green,fontWeight:700}}>✓ All poured</span>}
              </div>
            </div>); })}
          </div>);
        })()}

        {tab==="mpa"&&(()=>{
          // Build MPa summary from scope and tickets
          const mpaSummary = {};
          SCOPE.forEach(r => {
            if (!r.mpa || !r.item) return;
            const num = parseMpaNum(r.mpa);
            if (!num) return;
            const key = `${num} MPa`;
            if (!mpaSummary[key]) mpaSummary[key] = { mpa: key, color: num>=35?C.purple:num>=32?C.blue:C.accent, scopeM3: 0, pouredM3: 0, tickets: [] };
            mpaSummary[key].scopeM3 += r.m3;
          });
          tickets.forEach(t => {
            if (!t.mix_design || !t.volume_m3) return;
            const num = parseMpaNum(t.mix_design);
            if (!num) return;
            const key = `${num} MPa`;
            if (!mpaSummary[key]) mpaSummary[key] = { mpa: key, color: num>=35?C.purple:num>=32?C.blue:C.accent, scopeM3: 0, pouredM3: 0, tickets: [] };
            mpaSummary[key].pouredM3 += parseFloat(t.volume_m3) || 0;
            mpaSummary[key].tickets.push(t);
          });
          const rows = Object.values(mpaSummary).sort((a,b) => parseMpaNum(a.mpa) - parseMpaNum(b.mpa));
          const totalPouredAllMpa = rows.reduce((s,r) => s + r.pouredM3, 0);
          const totalScopeAllMpa  = rows.reduce((s,r) => s + r.scopeM3, 0);

          return (
            <div>
              <div style={{fontWeight:700,fontSize:18,marginBottom:6}}>Summary by MPa</div>
              <div style={{color:C.muted,fontSize:13,marginBottom:24}}>Scope and poured volumes grouped by concrete strength class</div>

              {/* Overall mini stats */}
              <div style={{display:"flex",gap:14,flexWrap:"wrap",marginBottom:28}}>
                {rows.map(r => {
                  const pct = r.scopeM3 > 0 ? (r.pouredM3/r.scopeM3)*100 : 0;
                  return (
                    <div key={r.mpa} style={{background:r.color+"18",border:`1px solid ${r.color}44`,borderRadius:14,padding:"18px 22px",flex:1,minWidth:150,textAlign:"center"}}>
                      <div style={{color:r.color,fontWeight:800,fontSize:22,marginBottom:4}}>{r.mpa}</div>
                      <div style={{color:C.text,fontWeight:800,fontSize:18,fontFamily:"monospace"}}>{r.pouredM3.toFixed(1)} m³</div>
                      <div style={{color:C.muted,fontSize:12,marginTop:2}}>of {r.scopeM3.toFixed(1)} m³ scope</div>
                      <div style={{marginTop:10}}><Bar pct={pct} color={r.color}/></div>
                      <div style={{color:r.color,fontWeight:700,fontSize:13,marginTop:6}}>{pct.toFixed(1)}%</div>
                    </div>
                  );
                })}
              </div>

              {/* Detailed breakdown per MPa */}
              {rows.map(r => {
                const pct = r.scopeM3 > 0 ? (r.pouredM3/r.scopeM3)*100 : 0;
                const remaining = Math.max(0, r.scopeM3 - r.pouredM3);
                // Group scope lines by this MPa
                const scopeLines = SCOPE.filter(s => s.mpa && parseMpaNum(s.mpa) === parseMpaNum(r.mpa) && s.item);
                return (
                  <div key={r.mpa} style={{background:C.card,border:`1px solid ${r.color}44`,borderRadius:14,padding:"20px 24px",marginBottom:20}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap",gap:10}}>
                      <div style={{display:"flex",alignItems:"center",gap:12}}>
                        <div style={{width:14,height:14,borderRadius:99,background:r.color}}/>
                        <span style={{fontWeight:800,fontSize:18,color:r.color}}>{r.mpa}</span>
                        <Badge color={r.color}>{r.tickets.length} ticket{r.tickets.length!==1?"s":""}</Badge>
                      </div>
                      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                        <Badge color={C.accent}>{r.pouredM3.toFixed(2)} m³ poured</Badge>
                        <Badge color={C.muted}>{r.scopeM3.toFixed(2)} m³ scope</Badge>
                        <Badge color={remaining>0?C.yellow:C.green}>{remaining.toFixed(2)} m³ remaining</Badge>
                      </div>
                    </div>
                    <Bar pct={pct} color={r.color}/>
                    <div style={{color:C.muted,fontSize:12,marginTop:6,marginBottom:16}}>{pct.toFixed(1)}% complete</div>

                    {/* Areas that use this MPa */}
                    <div style={{marginBottom:14}}>
                      <div style={{color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:10}}>Scope Elements</div>
                      <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                        {scopeLines.map((s,i) => {
                          const poured = pouredMap[`${s.area}|||${s.item}`] || 0;
                          const rem = Math.max(0, s.m3 - poured);
                          return (
                            <div key={i} style={{background:C.bg,borderRadius:8,padding:"8px 14px",fontSize:12,border:`1px solid ${C.border}`}}>
                              <span style={{fontWeight:700,color:C.text}}>{s.area} — {s.item}</span>
                              <span style={{color:C.muted,marginLeft:8}}>{s.m3} m³</span>
                              {poured > 0 && <span style={{color:C.green,marginLeft:6}}>✓ {poured.toFixed(1)} poured</span>}
                            </div>
                          );
                        })}
                      </div>
                    </div>

                    {/* Tickets for this MPa */}
                    {r.tickets.length > 0 && (
                      <div>
                        <div style={{color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:10}}>Tickets</div>
                        <div style={{overflowX:"auto"}}>
                          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                            <thead>
                              <tr style={{color:C.muted}}>
                                {["Ticket #","Date","Area","Element","Volume (m³)","Supplier"].map(h=>(
                                  <th key={h} style={{padding:"6px 12px",borderBottom:`1px solid ${C.border}`,fontWeight:700,textAlign:"left",whiteSpace:"nowrap"}}>{h}</th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              {r.tickets.map(t => (
                                <tr key={t.id} style={{borderBottom:`1px solid ${C.border}22`}}>
                                  <td style={{padding:"7px 12px",fontWeight:700,color:r.color}}>{t.ticket_number||"—"}</td>
                                  <td style={{padding:"7px 12px",color:C.sub}}>{t.date||"—"}</td>
                                  <td style={{padding:"7px 12px"}}>{t.area||"—"}</td>
                                  <td style={{padding:"7px 12px"}}>{t.item||"—"}</td>
                                  <td style={{padding:"7px 12px",fontFamily:"monospace",color:C.accent}}>{t.volume_m3?parseFloat(t.volume_m3).toFixed(2):"—"}</td>
                                  <td style={{padding:"7px 12px",color:C.sub}}>{t.supplier||"—"}</td>
                                </tr>
                              ))}
                              <tr style={{borderTop:`2px solid ${C.border}`,fontWeight:800}}>
                                <td colSpan={4} style={{padding:"8px 12px",color:C.muted}}>TOTAL</td>
                                <td style={{padding:"8px 12px",fontFamily:"monospace",color:C.accent}}>{r.pouredM3.toFixed(2)}</td>
                                <td/>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          );
        })()}

        {tab==="scope"&&(<div>
          <div style={{fontWeight:700,fontSize:18,marginBottom:20}}>Full Project Scope</div>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr style={{color:C.muted,textAlign:"left"}}>{["Area","Element","Spec MPa","Scope (m³)","Poured (m³)","Remaining (m³)","Overage (m³)","% Done"].map(h=><th key={h} style={{padding:"8px 14px",borderBottom:`1px solid ${C.border}`,fontWeight:700,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
              <tbody>{SCOPE.map((r,i)=>{ const poured=pouredMap[`${r.area}|||${r.item}`]||0; const rem=Math.max(0,r.m3-poured); const over=Math.max(0,poured-r.m3); const p=r.m3>0?(poured/r.m3)*100:0; return(<tr key={i} style={{borderBottom:`1px solid ${C.border}22`,background:over>0.01?C.red+"12":i%2===0?"transparent":C.card+"88"}}>
                <td style={{padding:"9px 14px",color:C.sub}}>{r.area}</td>
                <td style={{padding:"9px 14px"}}>{r.item}</td>
                <td style={{padding:"9px 14px"}}>{r.mpa?<MpaBadge mpa={r.mpa}/>:"—"}</td>
                <td style={{padding:"9px 14px",fontFamily:"monospace"}}>{r.m3.toFixed(1)}</td>
                <td style={{padding:"9px 14px",fontFamily:"monospace",color:poured>0?C.green:C.muted}}>{poured>0?poured.toFixed(2):"—"}</td>
                <td style={{padding:"9px 14px",fontFamily:"monospace",color:rem>0?C.yellow:C.green}}>{rem>0?rem.toFixed(2):"✓"}</td>
                <td style={{padding:"9px 14px",fontFamily:"monospace",color:over>0?C.red:C.muted}}>{over>0?over.toFixed(2):"—"}</td>
                <td style={{padding:"9px 14px"}}>{poured>0?<Badge color={over>0?C.red:p>=100?C.green:C.accent}>{fmt(p,0)}%</Badge>:<span style={{color:C.muted}}>—</span>}</td>
              </tr>); })}</tbody>
              <tfoot><tr style={{borderTop:`2px solid ${C.border}`,fontWeight:800}}>
                <td colSpan={3} style={{padding:"11px 14px"}}>TOTAL</td>
                <td style={{padding:"11px 14px",fontFamily:"monospace"}}>{fmt(TOTAL_SCOPE_M3,1)}</td>
                <td style={{padding:"11px 14px",fontFamily:"monospace",color:C.green}}>{totalPoured>0?fmt(totalPoured):"—"}</td>
                <td style={{padding:"11px 14px",fontFamily:"monospace",color:C.yellow}}>{fmt(remaining)}</td>
                <td style={{padding:"11px 14px",fontFamily:"monospace",color:totalOverage>0?C.red:C.muted}}>{totalOverage>0?fmt(totalOverage):"—"}</td>
                <td style={{padding:"11px 14px"}}><Badge color={totalOverage>0?C.red:pct>=100?C.green:C.accent}>{fmt(pct,1)}%</Badge></td>
              </tr></tfoot>
            </table>
          </div>
        </div>)}
      </div>

      {manualOpen&&(
        <div style={{position:"fixed",inset:0,background:"#000b",zIndex:100,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={e=>e.target===e.currentTarget&&setManualOpen(false)}>
          <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:16,padding:30,width:"92%",maxWidth:500,maxHeight:"90vh",overflowY:"auto"}}>
            <div style={{fontWeight:800,fontSize:18,marginBottom:20}}>✏️ Add Ticket Manually</div>
            {INPUT("date","Date","date")}{INPUT("ticket_number","Ticket / Docket #")}{INPUT("supplier","Supplier")}{INPUT("mix_design","Mix Design / Strength (e.g. 35 MPa)")}{INPUT("volume_m3","Volume (m³)","number")}{INPUT("volume_yd3","Volume (yd³)","number")}{INPUT("area","Area","text",AREAS)}{INPUT("item","Element Type","text",ITEMS)}{INPUT("invoice_number","Invoice #")}{INPUT("notes","Notes")}
            {manual.area&&manual.item&&manual.mix_design&&(()=>{ const preview=checkMpaMismatch({area:manual.area,item:manual.item,mix_design:manual.mix_design}); const spec=MPA_SPEC[`${manual.area}|||${manual.item}`]; if(preview) return(<div style={{background:"#450a0a",border:`1px solid ${C.red}`,borderRadius:8,padding:"10px 14px",marginBottom:12,fontSize:13,color:"#fca5a5"}}>⚠ MPa mismatch — you entered <b>{preview.ticketMpa}</b> but this element requires <b>{preview.specMpa}</b></div>); if(spec) return(<div style={{background:"#052e16",border:`1px solid ${C.green}44`,borderRadius:8,padding:"10px 14px",marginBottom:12,fontSize:13,color:"#86efac"}}>✓ Mix design matches spec ({spec})</div>); return null; })()}
            <div style={{display:"flex",gap:10,marginTop:18}}><button onClick={addManual} style={{background:C.accent,color:"#fff",border:"none",borderRadius:9,padding:"11px 0",fontWeight:800,cursor:"pointer",flex:1}}>Add Ticket</button><button onClick={()=>setManualOpen(false)} style={{background:C.bg,color:C.muted,border:`1px solid ${C.border}`,borderRadius:9,padding:"11px 18px",fontWeight:700,cursor:"pointer"}}>Cancel</button></div>
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// TIME & MATERIALS MODULE
// ═══════════════════════════════════════════════════════════════════════════════
const TM_STAGES = ["Open","Work ongoing","Awaiting signed sheets","Awaiting pricing","Cost support complete","RCO submitted","Included in change order","Billed on progress invoice","Closed"];
const TM_DOC_TYPES = ["Auto-detect","Signed field sheet","Priced daily breakdown","Written direction","Summary of extra work","RCO","Change order","Progress invoice","Receipt / backup","Other"];
const emptyTM = () => ({
  project:"Fortuna", trade:"", title:"", location:"", description:"", stage:"Open",
  direction_ref:"", rco_number:"", change_order_number:"", notes:"", timesheets:[], invoices:[], documents:[]
});
const money = n => `$${(parseFloat(n)||0).toLocaleString("en-CA",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const tmSupported = item => (item.timesheets||[]).reduce((s,t)=>s+(parseFloat(t.total)||0),0);
const tmInvoiced = item => (item.invoices||[]).reduce((s,t)=>s+(parseFloat(t.amount)||0),0);
function tmFlags(item){
  const flags=[];
  const sheets=item.timesheets||[], docs=item.documents||[], invoices=item.invoices||[];
  const signedSheets=sheets.filter(t=>t.signed);
  sheets.forEach(t=>{
    if(!t.sheet_number) flags.push(`${t.date||"Timesheet"}: sheet number missing`);
    if(t.sheet_number&&t.priced_sheet_number&&String(t.sheet_number).trim()!==String(t.priced_sheet_number).trim()) flags.push(`${t.date||"Timesheet"}: field sheet #${t.sheet_number} does not match priced sheet #${t.priced_sheet_number}`);
    if(!t.description) flags.push(`${t.sheet_number||t.date||"Timesheet"}: work description missing`);
    if(!t.signed) flags.push(`${t.sheet_number||t.date||"Timesheet"}: signed field sheet missing`);
    const calc=(parseFloat(t.labour)||0)+(parseFloat(t.equipment)||0)+(parseFloat(t.material)||0)+(parseFloat(t.subcontractor)||0)+(parseFloat(t.markup)||0);
    if(t.total!==""&&Math.abs(calc-(parseFloat(t.total)||0))>.02) flags.push(`${t.sheet_number||t.date||"Timesheet"}: total differs from entered breakdown by ${money(Math.abs(calc-(parseFloat(t.total)||0)))}`);
  });
  const numbers=sheets.map(t=>String(t.sheet_number||"").trim()).filter(Boolean);
  const duplicates=numbers.filter((n,i)=>numbers.indexOf(n)!==i);
  if(duplicates.length) flags.push(`Duplicate timesheet number: ${[...new Set(duplicates)].join(", ")}`);
  const invoiceNumbers=invoices.map(v=>String(v.invoice_number||"").trim()).filter(Boolean);
  const duplicateInvoices=invoiceNumbers.filter((n,i)=>invoiceNumbers.indexOf(n)!==i);
  if(duplicateInvoices.length) flags.push(`Duplicate progress invoice reference: ${[...new Set(duplicateInvoices)].join(", ")}`);
  const supported=tmSupported(item), invoiced=tmInvoiced(item);
  if(supported>5000&&!item.direction_ref&&!docs.some(d=>d.type==="Written direction")) flags.push("T&M exceeds $5,000 with no written direction attached or referenced");
  if(invoices.length&&Math.abs(invoiced-supported)>.02) flags.push(`Progress billing amount differs from supported T&M by ${money(Math.abs(invoiced-supported))}`);
  if(invoices.length&&!signedSheets.length) flags.push("Progress billing is linked but no signed field sheets are recorded");
  if(item.rco_number&&!docs.some(d=>d.type==="RCO")) flags.push(`RCO ${item.rco_number} referenced but document not attached`);
  if(item.change_order_number&&!docs.some(d=>d.type==="Change order")) flags.push(`CO ${item.change_order_number} referenced but document not attached`);
  if(item.stage==="Included in change order"&&!item.change_order_number&&!docs.some(d=>d.type==="Change order")) flags.push("Stage says Included in change order but no CO number or document is linked");
  if(item.stage==="Billed on progress invoice"&&!invoices.length) flags.push("Stage says Billed on progress invoice but no billing reference is linked");
  return flags;
}

function exportSheet(rows,headers,widths){
  const ws=XLSX.utils.json_to_sheet(rows,{header:headers});
  ws["!autofilter"]={ref:ws["!ref"]||`A1:${XLSX.utils.encode_col(headers.length-1)}1`};
  ws["!cols"]=widths.map(w=>({wch:w}));
  ws["!freeze"]={xSplit:0,ySplit:1,topLeftCell:"A2",activePane:"bottomLeft",state:"frozen"};
  return ws;
}

function TMModule({onBack}){
  const [items,setItems]=useState([]), [labourEntries,setLabourEntries]=useState([]), [submittedWeeks,setSubmittedWeeks]=useState([]);
  const [ready,setReady]=useState(false), [tab,setTab]=useState("dashboard"), [selected,setSelected]=useState(null);
  const [newOpen,setNewOpen]=useState(false), [draft,setDraft]=useState(emptyTM());
  const [sheetDraft,setSheetDraft]=useState({date:"",sheet_number:"",description:"",labour:"",equipment:"",material:"",subcontractor:"",markup:"",total:"",signed:false,labour_entries:[],equipment_entries:[],material_entries:[]});
  const [invoiceDraft,setInvoiceDraft]=useState({invoice_number:"",date:"",amount:"",timesheet_refs:"",notes:""});
  const [docDraft,setDocDraft]=useState({type:"Auto-detect",reference:"",file:null});
  const [labourDraft,setLabourDraft]=useState({week_ending:"",date:"",labourer:"",regular_hours:"",overtime_hours:"",project:"Fortuna",task:"",location:"",tm_id:"",backcharge:false});
  const [toast,setToast]=useState(null), [loading,setLoading]=useState(false), [search,setSearch]=useState(""), [stageFilter,setStageFilter]=useState("All stages"), [labourSearch,setLabourSearch]=useState("");
  const [editingLabour,setEditingLabour]=useState(null);
  const [intakeOpen,setIntakeOpen]=useState(false), [intakeFile,setIntakeFile]=useState(null), [intakeReview,setIntakeReview]=useState(null), [intakeLoading,setIntakeLoading]=useState(false), [intakeNotice,setIntakeNotice]=useState("");
  const showToast=(m,type="ok")=>{setToast({m,type});setTimeout(()=>setToast(null),3200);};

  const cleanMatch=s=>String(s||"").toLowerCase().replace(/[^a-z0-9]+/g," ").trim();
  const sheetNumbers=t=>[t?.sheet_number,t?.priced_sheet_number,t?.field_sheet_number].map(x=>String(x||"").trim()).filter(Boolean);
  const findExistingMatch=(incomingSheets,contractor,description)=>{
    let best=null;
    for(const item of items){
      let score=0,reason="";
      const existingNumbers=(item.timesheets||[]).flatMap(sheetNumbers);
      const incomingNumbers=(incomingSheets||[]).flatMap(sheetNumbers);
      const matchedNumber=incomingNumbers.find(n=>existingNumbers.includes(n));
      if(matchedNumber){score+=100;reason=`sheet #${matchedNumber}`;}
      const tradeA=cleanMatch(item.trade),tradeB=cleanMatch(contractor);
      if(tradeA&&tradeB&&(tradeA.includes(tradeB)||tradeB.includes(tradeA))){score+=20;if(!reason)reason="same trade";}
      const dates=new Set((item.timesheets||[]).map(t=>t.date).filter(Boolean));
      if((incomingSheets||[]).some(t=>t.date&&dates.has(t.date)))score+=20;
      const words=new Set(cleanMatch(`${item.title} ${item.description}`).split(" ").filter(w=>w.length>3));
      const incomingWords=cleanMatch(description).split(" ").filter(w=>w.length>3);
      if(incomingWords.filter(w=>words.has(w)).length>=2)score+=15;
      if(!best||score>best.score)best={item,score,reason};
    }
    return best&&best.score>=70?best:null;
  };

  const mergeFieldRows=(original=[],priced=[],key)=>original.length?original.map(row=>{const match=priced.find(p=>cleanMatch(p[key])===cleanMatch(row[key]));return match?{...match,...row,rate:match.rate??row.rate,amount:match.amount??row.amount}:row}):priced;
  const mergeUploadedItem=(existing,incoming)=>{
    const mergedSheets=[...(existing.timesheets||[])];
    for(const added of incoming.timesheets||[]){
      const nums=sheetNumbers(added);
      const idx=mergedSheets.findIndex(old=>sheetNumbers(old).some(n=>nums.includes(n))||(old.date&&added.date&&old.date===added.date&&cleanMatch(old.description)===cleanMatch(added.description)));
      if(idx<0){mergedSheets.push(added);continue;}
      const old=mergedSheets[idx];
      mergedSheets[idx]={...old,
        date:old.date||added.date, sheet_number:old.sheet_number||added.sheet_number,
        priced_sheet_number:added.priced_sheet_number||old.priced_sheet_number,
        description:old.description||added.description, signed:old.signed||added.signed,
        labour:(+added.labour||0)||(+old.labour||0), equipment:(+added.equipment||0)||(+old.equipment||0),
        material:(+added.material||0)||(+old.material||0), subcontractor:(+added.subcontractor||0)||(+old.subcontractor||0),
        markup:(+added.markup||0)||(+old.markup||0), total:(+added.total||0)||(+old.total||0),
        labour_entries:mergeFieldRows(old.labour_entries,added.labour_entries,"name"),
        equipment_entries:mergeFieldRows(old.equipment_entries,added.equipment_entries,"machine_number"),
        material_entries:(old.material_entries||[]).length?old.material_entries:added.material_entries||[]
      };
    }
    const notes=[existing.notes,incoming.notes].filter(Boolean).filter((x,i,a)=>a.indexOf(x)===i).join("\n\n");
    return {...existing,timesheets:mergedSheets,documents:[...(existing.documents||[]),...(incoming.documents||[])],invoices:[...(existing.invoices||[]),...(incoming.invoices||[])],notes,
      direction_ref:existing.direction_ref||incoming.direction_ref,rco_number:existing.rco_number||incoming.rco_number,change_order_number:existing.change_order_number||incoming.change_order_number};
  };

  useEffect(()=>{(async()=>{const saved=await storageGet("tm-data");setItems((saved?.items||[]).filter(x=>x.id!=="TM-TEST-DEXTER"));setLabourEntries(saved?.labourEntries||[]);setSubmittedWeeks(saved?.submittedWeeks||[]);setReady(true);})()},[]);
  useEffect(()=>{if(!ready)return;const timer=setTimeout(()=>storageSet("tm-data",{items,labourEntries,submittedWeeks}),350);return()=>clearTimeout(timer)},[items,labourEntries,submittedWeeks,ready]);
  useEffect(()=>{if(selected){const fresh=items.find(x=>x.id===selected.id);if(fresh)setSelected(fresh)}},[items]);

  const saveItem=updated=>setItems(prev=>prev.map(x=>x.id===updated.id?updated:x));
  const deleteItem=item=>{if(!window.confirm(`Delete ${item.id} · ${item.title}? This removes its timesheets, progress-billing references and document links.`))return;setItems(prev=>prev.filter(x=>x.id!==item.id));setLabourEntries(prev=>prev.map(e=>e.tm_id===item.id?{...e,tm_id:""}:e));setSelected(null);showToast(`${item.id} deleted`)};
  const createItem=()=>{if(!draft.trade||!draft.title){showToast("Enter the trade and T&M title.","err");return}const next=Math.max(0,...items.map(x=>parseInt(String(x.id||"").replace(/\D/g,""))||0))+1;const id=`TM-${String(next).padStart(3,"0")}`;setItems([{...draft,id,created_at:new Date().toISOString()},...items]);setDraft(emptyTM());setNewOpen(false);showToast(`${id} created ✓`)};
  const addSheet=()=>{if(!selected||!sheetDraft.date){showToast("Enter the work date.","err");return}const calc=(+sheetDraft.labour||0)+(+sheetDraft.equipment||0)+(+sheetDraft.material||0)+(+sheetDraft.subcontractor||0)+(+sheetDraft.markup||0);const total=sheetDraft.total===""?calc:+sheetDraft.total;saveItem({...selected,timesheets:[...(selected.timesheets||[]),{...sheetDraft,id:Date.now(),total}]});setSheetDraft({date:"",sheet_number:"",description:"",labour:"",equipment:"",material:"",subcontractor:"",markup:"",total:"",signed:false,labour_entries:[],equipment_entries:[],material_entries:[]});showToast("Timesheet added ✓")};
  const addInvoice=()=>{if(!selected||!invoiceDraft.invoice_number){showToast("Enter the progress invoice number.","err");return}saveItem({...selected,invoices:[...(selected.invoices||[]),{...invoiceDraft,id:Date.now()}]});setInvoiceDraft({invoice_number:"",date:"",amount:"",timesheet_refs:"",notes:""});showToast("Progress billing linked ✓")};
  async function extractTMDocument(file){
    const base64=await new Promise((resolve,reject)=>{const r=new FileReader();r.onload=()=>resolve(r.result.split(",")[1]);r.onerror=reject;r.readAsDataURL(file)});
    const block=file.type==="application/pdf"?{type:"document",source:{type:"base64",media_type:"application/pdf",data:base64}}:{type:"image",source:{type:"base64",media_type:file.type||"image/jpeg",data:base64}};
    const prompt=`You are extracting construction Time & Materials documents for Southwest Construction Management. The upload may contain signed handwritten field slips, contractor-priced daily T&M breakdowns, a summary of extra work, an RCO, a change order, a regular progress invoice, receipts, or several of these in one PDF. Extract facts only; do not decide whether rates are commercially acceptable. Preserve every daily record separately and pair each priced breakdown with its corresponding signed field slip using date, description and sheet number. Read EVERY handwritten labour row and equipment row. Labour hours and equipment hours are operational quantities, not dollar amounts: never put hours into the labour or equipment cost fields. When a form has one Hours column, store it as regular_hours unless the form explicitly identifies overtime. A priced daily breakdown, summary, or RCO is cost support and is NOT an invoice. Populate invoice fields only when the document is clearly an actual progress invoice with an invoice number and billing date. Compare field-slip numbers with priced-sheet numbers when both are visible. Return ONLY one valid JSON object:
{"document_type":"Signed field sheet|Priced daily breakdown|Written direction|Summary of extra work|RCO|Change order|Progress invoice|Receipt / backup|Mixed package|Other","project":"","contractor":"","scope_title":"","location":"","description":"","direction_ref":"","rco_number":"","change_order_number":"","invoice_number":"","invoice_date":"YYYY-MM-DD or empty","invoice_amount_before_tax":number or null,"timesheets":[{"date":"YYYY-MM-DD or empty","field_sheet_number":"","priced_sheet_number":"","description":"","labour":number,"equipment":number,"material":number,"subcontractor":number,"markup":number,"total":number,"signed_by_southwest":true|false,"labour_entries":[{"name":"","classification":"","regular_hours":number,"overtime_hours":number,"rate":number|null,"amount":number|null}],"equipment_entries":[{"type":"","machine_number":"","hours":number,"rate":number|null,"amount":number|null}],"material_entries":[{"supplier":"","type":"","quantity":number,"unit":"","rate":number|null,"amount":number|null}]}],"notes":"brief extraction notes"}`;
    const res=await fetch("/api/claude",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:5000,messages:[{role:"user",content:[block,{type:"text",text:prompt}]}]})});
    const data=await res.json().catch(()=>({}));
    if(!res.ok||data.error) throw new Error(data.error?.message||`Document reader failed (${res.status})`);
    const text=data.content?.map(x=>x.text||"").join("")||"";
    if(!text) throw new Error("The document reader returned no text.");
    return extractJSON(text);
  }
  async function readNewTMForm(){
    if(!intakeFile){showToast("Choose the trade's T&M form first.","err");return}
    setIntakeLoading(true);setIntakeNotice("");
    let extracted=null,url=null;
    try{
      const base64=await new Promise((resolve,reject)=>{const r=new FileReader();r.onload=()=>resolve(r.result.split(",")[1]);r.onerror=reject;r.readAsDataURL(intakeFile)});
      const [readResult,uploadResult]=await Promise.all([
        extractTMDocument(intakeFile),
        fetch("/api/file-upload",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({base64,fileName:intakeFile.name,mimeType:intakeFile.type,folder:"tm"})}).then(r=>r.ok?r.json():{}).catch(()=>({}))
      ]);
      extracted=readResult;url=uploadResult.url||null;
    }catch(e){console.error(e);setIntakeNotice(`The form was stored, but the document reader could not extract it: ${e.message||"Unknown error"}`)}
    const x=extracted||{};
    const timesheets=(x.timesheets||[]).map((t,i)=>({id:Date.now()+i,date:t.date||"",sheet_number:t.field_sheet_number||"",priced_sheet_number:t.priced_sheet_number||"",description:t.description||x.description||"",labour:+t.labour||0,equipment:+t.equipment||0,material:+t.material||0,subcontractor:+t.subcontractor||0,markup:+t.markup||0,total:+t.total||0,signed:!!t.signed_by_southwest,labour_entries:(t.labour_entries||[]).map((r,j)=>({id:Date.now()+i*100+j,name:r.name||"",classification:r.classification||"",regular_hours:+r.regular_hours||0,overtime_hours:+r.overtime_hours||0,rate:r.rate??"",amount:r.amount??""})),equipment_entries:(t.equipment_entries||[]).map((r,j)=>({id:Date.now()+i*100+j+30,type:r.type||"",machine_number:r.machine_number||"",hours:+r.hours||0,rate:r.rate??"",amount:r.amount??""})),material_entries:(t.material_entries||[]).map((r,j)=>({id:Date.now()+i*100+j+60,supplier:r.supplier||"",type:r.type||"",quantity:+r.quantity||0,unit:r.unit||"",rate:r.rate??"",amount:r.amount??""}))}));
    const match=findExistingMatch(timesheets,x.contractor,x.scope_title||x.description);
    setIntakeReview({project:x.project||"Fortuna",trade:x.contractor||"",title:x.scope_title||x.description||"",location:x.location||"",description:x.description||"",stage:"Open",direction_ref:x.direction_ref||"",rco_number:x.rco_number||"",change_order_number:x.change_order_number||"",notes:x.notes||"",timesheets,invoices:x.invoice_number?[{id:Date.now()+999,invoice_number:x.invoice_number,date:x.invoice_date||"",amount:x.invoice_amount_before_tax||0,timesheet_refs:timesheets.map(t=>t.sheet_number).filter(Boolean).join(", "),notes:"Extracted from uploaded form — confirmed by user"}]:[],documents:[{id:Date.now()+1999,type:x.document_type||"Signed field sheet",reference:x.invoice_number||x.rco_number||x.change_order_number||timesheets[0]?.priced_sheet_number||timesheets[0]?.sheet_number||"",filename:intakeFile.name,file_url:url,added_at:new Date().toISOString(),extraction_notes:x.notes||""}],linked_item_id:match?.item.id||"",match_reason:match?.reason||""});
    setIntakeLoading(false);
  }
  function confirmNewTMForm(){
    if(!intakeReview?.trade||!intakeReview?.title){showToast("Confirm the trade and T&M scope before adding it.","err");return}
    if(intakeReview.linked_item_id){const existing=items.find(i=>i.id===intakeReview.linked_item_id);if(existing){const incoming={...intakeReview};delete incoming.linked_item_id;delete incoming.match_reason;const updated=mergeUploadedItem(existing,incoming);setItems(prev=>prev.map(i=>i.id===existing.id?updated:i));setSelected(updated);setIntakeOpen(false);setIntakeFile(null);setIntakeReview(null);setIntakeNotice("");showToast(`Pricing and document linked to ${existing.id} ✓`);return}}
    const next=Math.max(0,...items.map(x=>parseInt(String(x.id||"").replace(/\D/g,""))||0))+1,id=`TM-${String(next).padStart(3,"0")}`;
    const incoming={...intakeReview};delete incoming.linked_item_id;delete incoming.match_reason;setItems(prev=>[{...incoming,id,created_at:new Date().toISOString()},...prev]);setSelected(null);setIntakeOpen(false);setIntakeFile(null);setIntakeReview(null);setIntakeNotice("");showToast(`${id} added from uploaded form ✓`);
  }
  async function uploadTMDocument(){
    if(!selected||!docDraft.file){showToast("Choose a document to upload.","err");return}
    setLoading(true);let url=null,extracted=null;
    try{const base64=await new Promise((resolve,reject)=>{const r=new FileReader();r.onload=()=>resolve(r.result.split(",")[1]);r.onerror=reject;r.readAsDataURL(docDraft.file)});const [uploadResult,extractResult]=await Promise.all([fetch("/api/file-upload",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({base64,fileName:docDraft.file.name,mimeType:docDraft.file.type,folder:"tm"})}).then(r=>r.json()).catch(()=>({})),extractTMDocument(docDraft.file).catch(e=>{console.error(e);return null})]);url=uploadResult.url||null;extracted=extractResult}catch(e){console.error(e)}
    const type=docDraft.type==="Auto-detect"?(extracted?.document_type||"Other"):docDraft.type;
    const extractedSheets=(extracted?.timesheets||[]).map((t,i)=>({id:Date.now()+i,source_type:"priced",date:t.date||"",sheet_number:t.field_sheet_number||"",priced_sheet_number:t.priced_sheet_number||t.field_sheet_number||"",description:t.description||"",labour:+t.labour||0,equipment:+t.equipment||0,material:+t.material||0,subcontractor:+t.subcontractor||0,markup:+t.markup||0,total:+t.total||0,signed:false,labour_entries:(t.labour_entries||[]).map((r,j)=>({...r,id:Date.now()+i*100+j})),equipment_entries:(t.equipment_entries||[]).map((r,j)=>({...r,id:Date.now()+i*100+j+30})),material_entries:(t.material_entries||[]).map((r,j)=>({...r,id:Date.now()+i*100+j+60}))}));
    const extractedInvoices=extracted?.invoice_number?[{id:Date.now()+999,invoice_number:extracted.invoice_number,date:extracted.invoice_date||"",amount:extracted.invoice_amount_before_tax||0,timesheet_refs:extractedSheets.map(t=>t.sheet_number).filter(Boolean).join(", "),notes:"AI-extracted; review against source"}]:[];
    const incomingRefs=extractedSheets.flatMap(sheetNumbers),isOldPriced=t=>t.source_type==="priced"||(+t.labour||0)+(+t.equipment||0)+(+t.material||0)+(+t.subcontractor||0)+(+t.markup||0)+(+t.total||0)>0||(!t.signed&&!!t.priced_sheet_number);
    const retained=(selected.timesheets||[]).filter(old=>!isOldPriced(old)||(!sheetNumbers(old).some(n=>incomingRefs.includes(n))&&!extractedSheets.some(s=>s.date&&old.date===s.date)));
    const document={id:Date.now(),type,reference:docDraft.reference||extracted?.invoice_number||extracted?.rco_number||extracted?.change_order_number||extractedSheets[0]?.priced_sheet_number||"",filename:docDraft.file.name,file_url:url,added_at:new Date().toISOString(),extraction_notes:extracted?.notes||""};
    const notes=[selected.notes,extracted?.notes].filter(Boolean).filter((x,i,a)=>a.indexOf(x)===i).join("\n\n");
    const updated={...selected,timesheets:[...retained,...extractedSheets],invoices:[...(selected.invoices||[]),...extractedInvoices],documents:[...(selected.documents||[]),document],notes,direction_ref:selected.direction_ref||extracted?.direction_ref||"",rco_number:selected.rco_number||extracted?.rco_number||"",change_order_number:selected.change_order_number||extracted?.change_order_number||""};
    saveItem(updated);setDocDraft({type:"Auto-detect",reference:"",file:null});setLoading(false);showToast(extracted?`Priced breakdown added to ${selected.id} for comparison ✓`:`Document stored in ${selected.id} — AI extraction needs manual review`,extracted?"ok":"err");
  }
  const addLabour=()=>{if(!labourDraft.labourer||!labourDraft.date){showToast("Enter the labourer and work date.","err");return}setLabourEntries([{...labourDraft,id:Date.now()},...labourEntries]);setLabourDraft({...labourDraft,date:"",labourer:"",regular_hours:"",overtime_hours:"",task:"",location:"",tm_id:"",backcharge:false});showToast("Labour hours added ✓")};
  const exportTM=()=>{const wb=XLSX.utils.book_new();const register=items.map(i=>({"T&M #":i.id,Project:i.project,Trade:i.trade,Scope:i.title,Location:i.location,Stage:i.stage,"Written Direction":i.direction_ref,"RCO #":i.rco_number,"CO #":i.change_order_number,"Supported T&M":tmSupported(i),"Progress Billed":tmInvoiced(i),"Billing Variance":tmInvoiced(i)-tmSupported(i),Flags:tmFlags(i).join(" | ")}));const sheets=items.flatMap(i=>(i.timesheets||[]).map(t=>({"T&M #":i.id,Project:i.project,Trade:i.trade,Date:t.date,"Field Sheet #":t.sheet_number,"Priced Sheet #":t.priced_sheet_number,Description:t.description,"Field Labour Hours":(t.labour_entries||[]).reduce((s,r)=>s+(+r.regular_hours||0)+(+r.overtime_hours||0),0),"Field Equipment Hours":(t.equipment_entries||[]).reduce((s,r)=>s+(+r.hours||0),0),Labour:+t.labour||0,Equipment:+t.equipment||0,Materials:+t.material||0,Subcontractors:+t.subcontractor||0,Markup:+t.markup||0,Total:+t.total||0,"Signed Sheet":t.signed?"Yes":"No"})));const fieldLabour=items.flatMap(i=>(i.timesheets||[]).flatMap(t=>(t.labour_entries||[]).map(r=>({"T&M #":i.id,Project:i.project,Trade:i.trade,Date:t.date,"Field Sheet #":t.sheet_number,Name:r.name,Classification:r.classification,"Regular Hours":+r.regular_hours||0,"Overtime Hours":+r.overtime_hours||0,"Total Hours":(+r.regular_hours||0)+(+r.overtime_hours||0),Rate:r.rate===""?"":+r.rate||0,Amount:r.amount===""?"":+r.amount||0}))));const fieldEquipment=items.flatMap(i=>(i.timesheets||[]).flatMap(t=>(t.equipment_entries||[]).map(r=>({"T&M #":i.id,Project:i.project,Trade:i.trade,Date:t.date,"Field Sheet #":t.sheet_number,Equipment:r.type,"Machine #":r.machine_number,Hours:+r.hours||0,Rate:r.rate===""?"":+r.rate||0,Amount:r.amount===""?"":+r.amount||0}))));const invoices=items.flatMap(i=>(i.invoices||[]).map(v=>({"T&M #":i.id,Trade:i.trade,"Progress Invoice #":v.invoice_number,Date:v.date,"T&M / Sheets Referenced":v.timesheet_refs,Amount:+v.amount||0,"Supported T&M":tmSupported(i),Variance:(+v.amount||0)-tmSupported(i),Notes:v.notes})));XLSX.utils.book_append_sheet(wb,exportSheet(register,["T&M #","Project","Trade","Scope","Location","Stage","Written Direction","RCO #","CO #","Supported T&M","Progress Billed","Billing Variance","Flags"],[12,18,22,32,20,24,24,12,12,16,16,14,48]),"T&M Register");XLSX.utils.book_append_sheet(wb,exportSheet(sheets,["T&M #","Project","Trade","Date","Field Sheet #","Priced Sheet #","Description","Field Labour Hours","Field Equipment Hours","Labour","Equipment","Materials","Subcontractors","Markup","Total","Signed Sheet"],[12,18,22,13,15,16,45,18,20,14,14,14,16,14,14,13]),"Priced T&M");XLSX.utils.book_append_sheet(wb,exportSheet(fieldLabour,["T&M #","Project","Trade","Date","Field Sheet #","Name","Classification","Regular Hours","Overtime Hours","Total Hours","Rate","Amount"],[12,18,22,13,15,22,20,14,15,13,12,14]),"Field Labour");XLSX.utils.book_append_sheet(wb,exportSheet(fieldEquipment,["T&M #","Project","Trade","Date","Field Sheet #","Equipment","Machine #","Hours","Rate","Amount"],[12,18,22,13,15,28,16,12,12,14]),"Field Equipment");XLSX.utils.book_append_sheet(wb,exportSheet(invoices,["T&M #","Trade","Progress Invoice #","Date","T&M / Sheets Referenced","Amount","Supported T&M","Variance","Notes"],[12,22,18,13,30,16,16,14,40]),"Progress Billing");XLSX.writeFile(wb,`Southwest_TM_Register_${new Date().toISOString().slice(0,10)}.xlsx`)};
  const exportLabour=()=>{const rows=visibleLabour.map(e=>({"Week Ending":e.week_ending,Date:e.date,Labourer:e.labourer,"Regular Hours":+e.regular_hours||0,"Overtime Hours":+e.overtime_hours||0,"Total Hours":(+e.regular_hours||0)+(+e.overtime_hours||0),Project:e.project,Task:e.task,Location:e.location,"Related T&M":e.tm_id,"Potential Back-charge":e.backcharge?"Yes":"No"}));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,exportSheet(rows,["Week Ending","Date","Labourer","Regular Hours","Overtime Hours","Total Hours","Project","Task","Location","Related T&M","Potential Back-charge"],[14,14,22,15,16,14,18,38,22,16,20]),"Labour Hours");XLSX.writeFile(wb,`Southwest_Labour_Hours_${currentWeek||"All"}.xlsx`)};

  const allFlags=items.flatMap(i=>tmFlags(i).map(f=>({id:i.id,trade:i.trade,title:i.title,flag:f})));
  const invoiceOwners={};items.forEach(i=>(i.invoices||[]).forEach(v=>{const n=String(v.invoice_number||"").trim();if(n)(invoiceOwners[n]??=[]).push(i.id)}));Object.entries(invoiceOwners).filter(([,owners])=>new Set(owners).size>1).forEach(([n,owners])=>allFlags.push({id:[...new Set(owners)].join(" / "),trade:"Multiple T&M items",title:"Possible duplicate billing",flag:`Progress invoice ${n} is linked to more than one T&M item`}));
  const openItems=items.filter(i=>i.stage!=="Closed");
  const supported=items.reduce((s,i)=>s+tmSupported(i),0), invoiced=items.reduce((s,i)=>s+tmInvoiced(i),0);
  const weekOptions=[...new Set(labourEntries.map(e=>e.week_ending).filter(Boolean))].sort().reverse();
  const currentWeek=labourDraft.week_ending||weekOptions[0]||"";
  const weekly=labourEntries.filter(e=>!currentWeek||e.week_ending===currentWeek);
  const visibleLabour=weekly.filter(e=>[e.week_ending,e.date,e.labourer,e.project,e.task,e.location,e.tm_id,e.backcharge?"back-charge":""].join(" ").toLowerCase().includes(labourSearch.toLowerCase()));
  const labourSummary=Object.values(visibleLabour.reduce((a,e)=>{a[e.labourer]??={name:e.labourer,regular:0,ot:0};a[e.labourer].regular+=+e.regular_hours||0;a[e.labourer].ot+=+e.overtime_hours||0;return a},{}));
  const filtered=items.filter(i=>(stageFilter==="All stages"||i.stage===stageFilter)&&[i.id,i.trade,i.title,i.project,i.location,i.stage,i.direction_ref,i.rco_number,i.change_order_number,...(i.timesheets||[]).flatMap(t=>[t.sheet_number,t.priced_sheet_number,t.date,t.description]),...(i.invoices||[]).flatMap(v=>[v.invoice_number,v.date,v.timesheet_refs,v.notes])].join(" ").toLowerCase().includes(search.toLowerCase()));
  const field=(label,key,type="text",options=null)=><label style={{display:"block",marginBottom:12}}><span style={{display:"block",color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:.8,marginBottom:5}}>{label}</span>{options?<select value={draft[key]||""} onChange={e=>setDraft({...draft,[key]:e.target.value})} style={inputStyle}>{options.map(o=><option key={o}>{o}</option>)}</select>:<input type={type} value={draft[key]||""} onChange={e=>setDraft({...draft,[key]:e.target.value})} style={inputStyle}/>}</label>;

  return <div style={{minHeight:"100vh",background:C.bg,color:C.text,fontFamily:"'DM Sans','Segoe UI',sans-serif"}}>
    {toast&&<div style={{position:"fixed",top:20,right:20,zIndex:500,background:toast.type==="err"?"#450a0a":"#052e16",border:`1px solid ${toast.type==="err"?C.red:C.green}`,borderRadius:10,padding:"12px 18px",fontWeight:700}}>{toast.m}</div>}
    <header style={{height:64,borderBottom:`1px solid ${C.border}`,display:"flex",alignItems:"center",padding:"0 24px",gap:15,background:C.card}}><button onClick={onBack} style={ghostBtn}>← Modules</button><div style={{width:38,height:38,borderRadius:10,background:C.yellow+"22",display:"grid",placeItems:"center",fontSize:21}}>⏱️</div><div><div style={{fontWeight:800,fontSize:18}}>Time & Materials <span style={{fontSize:10,color:C.muted}}>v{APP_VERSION}</span></div><div style={{fontSize:11,color:C.muted}}>Tracking, document relationships and reconciliation</div></div></header>
    <div style={{display:"flex",minHeight:"calc(100vh - 65px)"}}>
      <nav style={{width:210,borderRight:`1px solid ${C.border}`,padding:16,background:C.card}}>{[["dashboard","Dashboard"],["items","Trade T&M"],["labour","SW Labour Hours"]].map(([id,label])=><button key={id} onClick={()=>setTab(id)} style={{...navBtn,background:tab===id?C.yellow+"22":"transparent",color:tab===id?C.yellow:C.sub}}>{label}</button>)}</nav>
      <main style={{flex:1,padding:"28px 30px",minWidth:0}}>
        {tab==="dashboard"&&<><div style={{display:"flex",justifyContent:"space-between",gap:15,alignItems:"center",marginBottom:18,flexWrap:"wrap"}}><div><h1 style={{margin:0,fontSize:25}}>T&M overview</h1><div style={{color:C.muted,fontSize:13,marginTop:5}}>Upload the trade's form, review what was read, then confirm it into the register.</div></div><div style={{display:"flex",gap:9,flexWrap:"wrap"}}><button onClick={exportTM} style={ghostBtn}>⬇ Export Register</button><button onClick={()=>setNewOpen(true)} style={ghostBtn}>Manual Entry</button><button onClick={()=>setIntakeOpen(true)} style={primaryBtn}>⬆ Upload T&M Form</button></div></div>
          <div style={{display:"flex",gap:10,marginBottom:22,flexWrap:"wrap"}}><input aria-label="Search T&M records" value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search trade, scope, sheet, progress invoice, RCO, CO…" style={{...inputStyle,maxWidth:480}}/><select aria-label="Filter by T&M stage" value={stageFilter} onChange={e=>setStageFilter(e.target.value)} style={{...inputStyle,width:210}}><option>All stages</option>{TM_STAGES.map(s=><option key={s}>{s}</option>)}</select>{(search||stageFilter!=="All stages")&&<button onClick={()=>{setSearch("");setStageFilter("All stages")}} style={ghostBtn}>Clear</button>}</div>
          <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:25}}><Stat label="Open items" value={openItems.length} color={C.yellow}/><Stat label="Supported T&M" value={money(supported)} color={C.blue}/><Stat label="Progress billed" value={money(invoiced)} color={C.teal}/><Stat label="Open flags" value={allFlags.length} color={allFlags.length?C.red:C.green}/></div>
          <div style={{display:"grid",gridTemplateColumns:"minmax(0,1.4fr) minmax(280px,.8fr)",gap:18}}><section style={panel}><div style={panelTitle}>Current T&M items <span style={{color:C.muted,fontSize:11,fontWeight:600}}>({filtered.length})</span></div>{filtered.length?filtered.slice(0,20).map(i=><TMRow key={i.id} item={i} onOpen={()=>setSelected(i)}/>):<Empty text={items.length?"No T&M items match your search or stage filter.":"No T&M items yet. Create your first item when testing begins."}/>}</section><section style={panel}><div style={panelTitle}>Items needing attention</div>{allFlags.filter(f=>!search||[f.id,f.trade,f.title,f.flag].join(" ").toLowerCase().includes(search.toLowerCase())).length?allFlags.filter(f=>!search||[f.id,f.trade,f.title,f.flag].join(" ").toLowerCase().includes(search.toLowerCase())).slice(0,12).map((f,i)=><div key={i} style={{padding:"10px 0",borderBottom:`1px solid ${C.border}`,fontSize:12}}><div style={{color:C.red,fontWeight:700}}>{f.id} · {f.trade}</div><div style={{color:C.sub,marginTop:3}}>{f.flag}</div></div>):<div style={{color:C.green,padding:"18px 0"}}>✓ No matching reconciliation flags</div>}</section></div>
        </>}
        {tab==="items"&&<><div style={{display:"flex",justifyContent:"space-between",gap:12,flexWrap:"wrap",marginBottom:20}}><div><h1 style={{margin:0,fontSize:24}}>Trade T&M register</h1><div style={{color:C.muted,fontSize:12,marginTop:4}}>Priced T&M support, signed slips, RCOs and change orders organized by scope. Progress billing is optional and linked later.</div></div><div style={{display:"flex",gap:8}}><button onClick={()=>setNewOpen(true)} style={ghostBtn}>Manual Entry</button><button onClick={()=>setIntakeOpen(true)} style={primaryBtn}>⬆ Upload T&M Form</button></div></div><div style={{display:"flex",gap:10,marginBottom:15,flexWrap:"wrap"}}><input aria-label="Search Trade T&M" value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search trade, scope, sheet, progress invoice, RCO, CO…" style={{...inputStyle,maxWidth:480}}/><select aria-label="Filter Trade T&M by stage" value={stageFilter} onChange={e=>setStageFilter(e.target.value)} style={{...inputStyle,width:210}}><option>All stages</option>{TM_STAGES.map(s=><option key={s}>{s}</option>)}</select>{(search||stageFilter!=="All stages")&&<button onClick={()=>{setSearch("");setStageFilter("All stages")}} style={ghostBtn}>Clear</button>}</div><section style={panel}><div style={{color:C.muted,fontSize:11,marginBottom:7}}>{filtered.length} item{filtered.length===1?"":"s"}</div>{filtered.length?filtered.map(i=><TMRow key={i.id} item={i} onOpen={()=>setSelected(i)}/>):<Empty text={(search||stageFilter!=="All stages")?"No matching T&M items.":"No T&M items yet. Upload a trade form to begin."}/>}</section></>}
        {tab==="labour"&&<><div style={{display:"flex",justifyContent:"space-between",gap:12,alignItems:"center",flexWrap:"wrap"}}><div><h1 style={{margin:0,fontSize:24}}>Southwest labour hours</h1><div style={{color:C.muted,fontSize:12,marginTop:4}}>Weekly union labourer hours only — no payroll calculations.</div></div><button onClick={exportLabour} style={ghostBtn}>⬇ Export Labour Hours</button></div><div style={{display:"grid",gridTemplateColumns:"minmax(320px,.8fr) minmax(0,1.3fr)",gap:18,marginTop:20}}><section style={panel}><div style={panelTitle}>Add hours</div><LabourForm draft={labourDraft} setDraft={setLabourDraft} items={items} onAdd={addLabour}/></section><section style={panel}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:10,flexWrap:"wrap"}}><div style={panelTitle}>Weekly summary</div><div style={{display:"flex",gap:8,flexWrap:"wrap"}}><input aria-label="Search labour hours" value={labourSearch} onChange={e=>setLabourSearch(e.target.value)} placeholder="Search labourer, task, location, T&M…" style={{...inputStyle,width:300}}/><select aria-label="Filter labour hours by week" value={currentWeek} onChange={e=>setLabourDraft({...labourDraft,week_ending:e.target.value})} style={{...inputStyle,width:170}}><option value="">All entries</option>{weekOptions.map(w=><option key={w}>{w}</option>)}</select>{labourSearch&&<button onClick={()=>setLabourSearch("")} style={ghostBtn}>Clear</button>}</div></div><table style={tableStyle}><thead><tr>{["Labourer","Regular","Overtime","Total"].map(h=><th key={h} style={thStyle}>{h}</th>)}</tr></thead><tbody>{labourSummary.map(r=><tr key={r.name}><td style={tdStyle}>{r.name}</td><td style={tdNum}>{r.regular}</td><td style={tdNum}>{r.ot}</td><td style={{...tdNum,fontWeight:800,color:C.yellow}}>{r.regular+r.ot}</td></tr>)}</tbody></table>{!labourSummary.length&&<div style={{color:C.muted,padding:"25px 0"}}>{labourSearch?"No labour hours match your search.":"No labour hours entered for this week."}</div>}</section></div><section style={{...panel,marginTop:18}}><div style={panelTitle}>Hour entries <span style={{color:C.muted,fontSize:11,fontWeight:600}}>({visibleLabour.length})</span></div>{visibleLabour.length?visibleLabour.map(e=><div key={e.id} style={{display:"grid",gridTemplateColumns:"110px 1fr 90px 90px 1.4fr 145px",gap:10,padding:"9px 0",borderBottom:`1px solid ${C.border}`,fontSize:12}}><span>{e.date}</span><b>{e.labourer}</b><span>{e.regular_hours||0} reg</span><span>{e.overtime_hours||0} OT</span><span style={{color:C.sub}}>{e.task||"—"}{e.tm_id?` · ${e.tm_id}`:""}</span><span style={{display:"flex",gap:5}}><button onClick={()=>setEditingLabour({...e})} style={miniEdit}>Edit</button><button onClick={()=>{if(window.confirm(`Delete hours for ${e.labourer} on ${e.date}?`))setLabourEntries(labourEntries.filter(x=>x.id!==e.id))}} style={miniDanger}>Delete</button></span></div>):<div style={{color:C.muted,padding:"24px 0"}}>{labourSearch?"No hour entries match your search.":"No hour entries yet."}</div>}</section></>}
      </main>
    </div>
    {intakeOpen&&<Modal title={intakeReview?"Review Uploaded T&M":"Upload Trade T&M Form"} onClose={()=>{setIntakeOpen(false);setIntakeReview(null);setIntakeFile(null);setIntakeNotice("")}} wide={!!intakeReview}>{!intakeReview?<><div style={{background:C.yellow+"12",border:`1px solid ${C.yellow}44`,borderRadius:12,padding:"16px 18px",marginBottom:18,color:C.sub,fontSize:13,lineHeight:1.6}}><b style={{color:C.yellow}}>How it works:</b> upload a PDF, scan or photo from the trade. The app reads the details and automatically links priced backup to an existing signed field sheet when the numbers match.</div><label style={{...formLabel,border:`2px dashed ${C.border}`,borderRadius:12,padding:"30px 20px",textAlign:"center",cursor:"pointer",color:C.sub}}><div style={{fontSize:30,marginBottom:9}}>📄</div><div style={{fontSize:14,color:C.text,textTransform:"none",letterSpacing:0}}>{intakeFile?intakeFile.name:"Choose a PDF or image"}</div><input type="file" accept="image/*,application/pdf" onChange={e=>setIntakeFile(e.target.files?.[0]||null)} style={{marginTop:14,color:C.sub}}/></label><div style={modalActions}><button onClick={readNewTMForm} disabled={!intakeFile||intakeLoading} style={{...primaryBtn,opacity:!intakeFile||intakeLoading?.6:1}}>{intakeLoading?"Reading Form…":"Read Form & Review"}</button><button onClick={()=>setIntakeOpen(false)} style={ghostBtn}>Cancel</button></div></>:<TMIntakeReview draft={intakeReview} setDraft={setIntakeReview} items={items} notice={intakeNotice} onBack={()=>setIntakeReview(null)} onConfirm={confirmNewTMForm}/>}</Modal>}
    {newOpen&&<Modal title="Manual T&M Entry" onClose={()=>setNewOpen(false)}><div style={{color:C.muted,fontSize:12,marginBottom:16}}>Use this only when there is no form to upload or the form cannot be read.</div>{field("Project","project")}{field("Trade / contractor","trade")}{field("T&M title / scope","title")}{field("Location","location")}{field("Stage","stage","text",TM_STAGES)}{field("Written direction reference","direction_ref")}{field("RCO number","rco_number")}{field("Change order number","change_order_number")}<label style={formLabel}>Description<textarea value={draft.description} onChange={e=>setDraft({...draft,description:e.target.value})} style={{...inputStyle,minHeight:80,resize:"vertical"}}/></label><div style={modalActions}><button onClick={createItem} style={primaryBtn}>Create T&M Item</button><button onClick={()=>setNewOpen(false)} style={ghostBtn}>Cancel</button></div></Modal>}
    {editingLabour&&<Modal title="Edit Labour Hours" onClose={()=>setEditingLabour(null)}><LabourForm draft={editingLabour} setDraft={setEditingLabour} items={items} submitLabel="Save Changes" onAdd={()=>{setLabourEntries(prev=>prev.map(e=>e.id===editingLabour.id?editingLabour:e));setEditingLabour(null);showToast("Labour entry updated ✓")}}/></Modal>}
    {selected&&<TMDetail item={selected} onClose={()=>setSelected(null)} onSave={saveItem} onDelete={deleteItem} sheetDraft={sheetDraft} setSheetDraft={setSheetDraft} onAddSheet={addSheet} invoiceDraft={invoiceDraft} setInvoiceDraft={setInvoiceDraft} onAddInvoice={addInvoice} docDraft={docDraft} setDocDraft={setDocDraft} onUpload={uploadTMDocument} loading={loading}/>} 
  </div>
}

const inputStyle={width:"100%",background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,padding:"9px 11px",color:C.text,fontSize:13,boxSizing:"border-box"};
const primaryBtn={background:C.yellow,color:"#1c1600",border:"none",borderRadius:8,padding:"10px 16px",fontWeight:800,cursor:"pointer"};
const ghostBtn={background:"transparent",color:C.sub,border:`1px solid ${C.border}`,borderRadius:8,padding:"8px 12px",fontWeight:700,cursor:"pointer"};
const navBtn={display:"block",width:"100%",border:"none",borderRadius:8,padding:"11px 12px",textAlign:"left",fontWeight:750,cursor:"pointer",marginBottom:6};
const panel={background:C.card,border:`1px solid ${C.border}`,borderRadius:14,padding:"18px 20px"};
const panelTitle={fontWeight:800,fontSize:15,marginBottom:12};
const tableStyle={width:"100%",borderCollapse:"collapse",fontSize:12};
const thStyle={textAlign:"left",color:C.muted,padding:"8px 7px",borderBottom:`1px solid ${C.border}`,fontSize:10,textTransform:"uppercase"};
const tdStyle={padding:"9px 7px",borderBottom:`1px solid ${C.border}88`};
const tdNum={...tdStyle,textAlign:"right",fontFamily:"monospace"};
const formLabel={display:"block",color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:.7,marginBottom:12};
const modalActions={display:"flex",gap:10,justifyContent:"flex-end",marginTop:18};
const miniDanger={background:C.red+"18",color:C.red,border:`1px solid ${C.red}44`,borderRadius:6,padding:"4px 8px",cursor:"pointer",fontSize:10,fontWeight:700};
const miniEdit={background:C.blue+"18",color:C.blue,border:`1px solid ${C.blue}44`,borderRadius:6,padding:"4px 8px",cursor:"pointer",fontSize:10,fontWeight:700};

function FieldRecords({sheet,onChange,editable=false}){
  const labour=sheet.labour_entries||[], equipment=sheet.equipment_entries||[];
  const update=(group,id,key,value)=>onChange?.({...sheet,[group]:(sheet[group]||[]).map(r=>r.id===id?{...r,[key]:value}:r)});
  const remove=(group,id)=>onChange?.({...sheet,[group]:(sheet[group]||[]).filter(r=>r.id!==id)});
  const addLabour=()=>onChange?.({...sheet,labour_entries:[...labour,{id:Date.now()+Math.random(),name:"",classification:"",regular_hours:"",overtime_hours:"",rate:"",amount:""}]});
  const addEquipment=()=>onChange?.({...sheet,equipment_entries:[...equipment,{id:Date.now()+Math.random(),type:"",machine_number:"",hours:"",rate:"",amount:""}]});
  const cell=(group,row,key,type="text",width=110)=>editable?<input type={type} value={row[key]??""} onChange={e=>update(group,row.id,key,e.target.value)} style={{...inputStyle,minWidth:width,padding:"7px 8px"}}/>:(row[key]!==""&&row[key]!=null?row[key]:"—");
  if(!editable&&!labour.length&&!equipment.length)return <div style={{color:C.muted,fontSize:11}}>No field labour or equipment hours recorded.</div>;
  return <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(360px,1fr))",gap:12}}>
    <div><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><b style={{fontSize:12}}>Field labour hours</b>{editable&&<button onClick={addLabour} style={miniEdit}>+ Labour Row</button>}</div><div style={{overflowX:"auto"}}><table style={tableStyle}><thead><tr>{["Name","Classification","Regular","OT",...(editable?[""]:[])].map(h=><th key={h} style={thStyle}>{h}</th>)}</tr></thead><tbody>{labour.map(r=><tr key={r.id}><td style={tdStyle}>{cell("labour_entries",r,"name","text",135)}</td><td style={tdStyle}>{cell("labour_entries",r,"classification","text",120)}</td><td style={tdStyle}>{cell("labour_entries",r,"regular_hours","number",75)}</td><td style={tdStyle}>{cell("labour_entries",r,"overtime_hours","number",65)}</td>{editable&&<td style={tdStyle}><button onClick={()=>remove("labour_entries",r.id)} style={miniDanger}>×</button></td>}</tr>)}</tbody></table></div>{!labour.length&&<div style={{color:C.muted,fontSize:11,padding:"8px 0"}}>No labour rows found.</div>}</div>
    <div><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><b style={{fontSize:12}}>Field equipment hours</b>{editable&&<button onClick={addEquipment} style={miniEdit}>+ Equipment Row</button>}</div><div style={{overflowX:"auto"}}><table style={tableStyle}><thead><tr>{["Equipment","Machine #","Hours",...(editable?[""]:[])].map(h=><th key={h} style={thStyle}>{h}</th>)}</tr></thead><tbody>{equipment.map(r=><tr key={r.id}><td style={tdStyle}>{cell("equipment_entries",r,"type","text",150)}</td><td style={tdStyle}>{cell("equipment_entries",r,"machine_number","text",100)}</td><td style={tdStyle}>{cell("equipment_entries",r,"hours","number",75)}</td>{editable&&<td style={tdStyle}><button onClick={()=>remove("equipment_entries",r.id)} style={miniDanger}>×</button></td>}</tr>)}</tbody></table></div>{!equipment.length&&<div style={{color:C.muted,fontSize:11,padding:"8px 0"}}>No equipment rows found.</div>}</div>
  </div>;
}

function Empty({text,action,onAction}){return <div style={{padding:"34px 10px",textAlign:"center",color:C.muted}}><div>{text}</div>{action&&<button onClick={onAction} style={{...ghostBtn,marginTop:12,color:C.yellow}}>{action}</button>}</div>}
function TMRow({item,onOpen}){const flags=tmFlags(item);return <button onClick={onOpen} style={{width:"100%",display:"grid",gridTemplateColumns:"90px minmax(180px,1.4fr) minmax(130px,.8fr) 130px 120px 95px",gap:10,alignItems:"center",textAlign:"left",background:"transparent",color:C.text,border:"none",borderBottom:`1px solid ${C.border}`,padding:"13px 5px",cursor:"pointer"}}><b style={{color:C.yellow}}>{item.id}</b><span><b>{item.title}</b><small style={{display:"block",color:C.muted,marginTop:3}}>{item.project} · {item.location||"No location"}</small></span><span>{item.trade}</span><span>{money(tmSupported(item))}</span><Badge color={item.stage==="Closed"?C.green:C.blue}>{item.stage}</Badge><span style={{color:flags.length?C.red:C.green,fontWeight:700,fontSize:12}}>{flags.length?`⚠ ${flags.length} flag${flags.length>1?"s":""}`:"✓ Matched"}</span></button>}
function Modal({title,onClose,children,wide=false}){return <div style={{position:"fixed",inset:0,background:"#000b",zIndex:400,display:"flex",alignItems:"center",justifyContent:"center",padding:20}} onClick={e=>e.target===e.currentTarget&&onClose()}><div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:16,width:"100%",maxWidth:wide?1100:560,maxHeight:"92vh",overflowY:"auto",padding:24}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}><div style={{fontWeight:850,fontSize:19}}>{title}</div><button onClick={onClose} style={ghostBtn}>×</button></div>{children}</div></div>}
function LabourForm({draft,setDraft,items,onAdd,submitLabel="Add Hours"}){const f=(label,key,type="text")=><label style={formLabel}>{label}<input type={type} value={draft[key]||""} onChange={e=>setDraft({...draft,[key]:e.target.value})} style={inputStyle}/></label>;return <>{f("Week ending","week_ending","date")}{f("Work date","date","date")}{f("Labourer","labourer")}<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>{f("Regular hours","regular_hours","number")}{f("Overtime hours","overtime_hours","number")}</div>{f("Project","project")}{f("Task / description","task")}{f("Location","location")}<label style={formLabel}>Related T&M item<select value={draft.tm_id||""} onChange={e=>setDraft({...draft,tm_id:e.target.value})} style={inputStyle}><option value="">Not related to T&M</option>{items.map(i=><option key={i.id} value={i.id}>{i.id} · {i.title}</option>)}</select></label><label style={{display:"flex",gap:8,color:C.sub,fontSize:12,marginBottom:14}}><input type="checkbox" checked={draft.backcharge} onChange={e=>setDraft({...draft,backcharge:e.target.checked})}/>Potential back-charge</label><button onClick={onAdd} style={{...primaryBtn,width:"100%"}}>{submitLabel}</button></>}

function TMDetail({item,onClose,onSave,onDelete,sheetDraft,setSheetDraft,onAddSheet,invoiceDraft,setInvoiceDraft,onAddInvoice,docDraft,setDocDraft,onUpload,loading}){
  const [view,setView]=useState("summary"), [editingItem,setEditingItem]=useState(false), [itemDraft,setItemDraft]=useState({...item});
  const [editingSheet,setEditingSheet]=useState(null), [editingInvoice,setEditingInvoice]=useState(null), [editingDoc,setEditingDoc]=useState(null);
  const flags=tmFlags(item), supported=tmSupported(item), invoiced=tmInvoiced(item);
  const sf=(label,key,type="number")=><label style={formLabel}>{label}<input type={type} value={sheetDraft[key]||""} onChange={e=>setSheetDraft({...sheetDraft,[key]:e.target.value})} style={inputStyle}/></label>;
  const inf=(label,key,type="text")=><label style={formLabel}>{label}<input type={type} value={invoiceDraft[key]||""} onChange={e=>setInvoiceDraft({...invoiceDraft,[key]:e.target.value})} style={inputStyle}/></label>;
  return <Modal title={`${item.id} · ${item.title}`} onClose={onClose} wide><div style={{display:"flex",gap:7,flexWrap:"wrap",marginBottom:18}}>{[["summary","Summary"],["field","Field Sheet"],["priced","Priced Breakdown"],["compare","Hours Comparison"],["documents","Upload Pricing & Documents"],["invoices","Progress Billing"]].map(([v,label])=><button key={v} onClick={()=>{setView(v);if(v==="documents")setDocDraft({...docDraft,type:"Priced daily breakdown"})}} style={{...ghostBtn,color:view===v?C.yellow:C.sub,borderColor:view===v?C.yellow:C.border}}>{label}</button>)}</div>{view==="field"&&<TMFieldSheetView item={item}/>} {view==="priced"&&<TMPricedSheetView item={item}/>} {view==="compare"&&<TMHoursComparison item={item}/>} 
    {view==="summary"&&<><div style={{display:"flex",justifyContent:"flex-end",gap:8,marginBottom:10}}><button onClick={()=>{setItemDraft({...item});setEditingItem(!editingItem)}} style={miniEdit}>{editingItem?"Cancel Edit":"Edit T&M Item"}</button><button onClick={()=>onDelete(item)} style={miniDanger}>Delete T&M Item</button></div><div style={{display:"grid",gridTemplateColumns:"minmax(0,1.3fr) minmax(300px,.7fr)",gap:18}}><section style={panel}>{editingItem?<TMItemEditor draft={itemDraft} setDraft={setItemDraft} onSave={()=>{onSave(itemDraft);setEditingItem(false)}}/>:<><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>{[["Project",item.project],["Trade",item.trade],["Location",item.location||"—"],["Stage",item.stage],["Written direction",item.direction_ref||"—"],["RCO / CO",[item.rco_number&&`RCO ${item.rco_number}`,item.change_order_number&&`CO ${item.change_order_number}`].filter(Boolean).join(" · ")||"—"],["Description",item.description||"—"],["Notes",item.notes||"—"]].map(([a,b])=><div key={a}><small style={{color:C.muted,textTransform:"uppercase",fontWeight:700}}>{a}</small><div style={{marginTop:4}}>{b}</div></div>)}</div><div style={{marginTop:18}}><label style={formLabel}>Update stage<select value={item.stage} onChange={e=>onSave({...item,stage:e.target.value})} style={inputStyle}>{TM_STAGES.map(s=><option key={s}>{s}</option>)}</select></label></div></>}</section><section style={panel}><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}><Stat label="Supported T&M" value={money(supported)} color={C.blue}/><Stat label="Progress billed" value={money(invoiced)} color={C.teal}/></div><div style={{color:C.muted,fontSize:11,marginTop:9}}>A progress invoice is optional and may be linked after the T&M is included in a change order.</div><div style={{marginTop:15,fontWeight:800}}>Reconciliation</div>{flags.length?flags.map((f,i)=><div key={i} style={{color:"#fca5a5",fontSize:12,padding:"8px 0",borderBottom:`1px solid ${C.border}`}}>⚠ {f}</div>):<div style={{color:C.green,marginTop:10}}>✓ No discrepancies found</div>}</section></div></>}
    {view==="timesheets"&&<><table style={tableStyle}><thead><tr>{["Date","Field sheet #","Priced sheet #","Labour cost","Equipment cost","Total","Signed",""].map(h=><th key={h} style={thStyle}>{h}</th>)}</tr></thead><tbody>{(item.timesheets||[]).map(t=><><tr key={`${t.id}-main`}><td style={tdStyle}>{t.date}</td><td style={{...tdStyle,color:t.priced_sheet_number&&t.priced_sheet_number!==t.sheet_number?C.red:C.text}}>{t.sheet_number||"—"}</td><td style={{...tdStyle,color:t.priced_sheet_number&&t.priced_sheet_number!==t.sheet_number?C.red:C.text}}>{t.priced_sheet_number||"—"}</td><td style={tdNum}>{money(t.labour)}</td><td style={tdNum}>{money(t.equipment)}</td><td style={{...tdNum,fontWeight:800}}>{money(t.total)}</td><td style={tdStyle}>{t.signed?"✓":"⚠"}</td><td style={tdStyle}><span style={{display:"flex",gap:4}}><button onClick={()=>setEditingSheet({...t})} style={miniEdit}>Edit</button><button onClick={()=>{if(window.confirm(`Delete timesheet ${t.sheet_number||t.date}?`))onSave({...item,timesheets:item.timesheets.filter(x=>x.id!==t.id)})}} style={miniDanger}>Delete</button></span></td></tr><tr key={`${t.id}-field`}><td colSpan={8} style={{padding:"10px 12px 18px",background:C.bg+"88",borderBottom:`1px solid ${C.border}`}}><FieldRecords sheet={t}/></td></tr></>)}</tbody></table><section style={{...panel,marginTop:18}}><div style={panelTitle}>Add field sheet / priced breakdown</div><div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(120px,1fr))",gap:10}}>{sf("Work date","date","date")}{sf("Field sheet #","sheet_number","text")}{sf("Priced sheet #","priced_sheet_number","text")}{sf("Labour cost","labour")}{sf("Equipment cost","equipment")}{sf("Materials cost","material")}{sf("Subcontractors","subcontractor")}{sf("Markup","markup")}{sf("Entered total (optional)","total")}</div>{sf("Work description","description","text")}<FieldRecords sheet={sheetDraft} editable onChange={setSheetDraft}/><label style={{display:"flex",gap:8,color:C.sub,fontSize:12,margin:"14px 0"}}><input type="checkbox" checked={sheetDraft.signed} onChange={e=>setSheetDraft({...sheetDraft,signed:e.target.checked})}/>Signed field sheet is present</label><button onClick={onAddSheet} style={primaryBtn}>Add Timesheet</button></section></>}
    {view==="documents"&&<><div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(250px,1fr))",gap:10}}>{(item.documents||[]).map(d=><div key={d.id} style={panel}><Badge color={C.yellow}>{d.type}</Badge><div style={{fontWeight:750,marginTop:9}}>{d.filename}</div><div style={{color:C.muted,fontSize:11,marginTop:4}}>{d.reference||"No reference"}</div>{d.file_url&&<a href={d.file_url} target="_blank" rel="noreferrer" style={{color:C.blue,fontSize:12,display:"inline-block",marginTop:9}}>Open document ↗</a>}<div style={{display:"flex",gap:5,justifyContent:"flex-end",marginTop:8}}><button onClick={()=>setEditingDoc({...d})} style={miniEdit}>Edit</button><button onClick={()=>{if(window.confirm(`Delete ${d.filename}?`))onSave({...item,documents:item.documents.filter(x=>x.id!==d.id)})}} style={miniDanger}>Delete</button></div></div>)}</div><section style={{...panel,marginTop:18}}><div style={panelTitle}>Attach document</div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}><label style={formLabel}>Document type<select value={docDraft.type} onChange={e=>setDocDraft({...docDraft,type:e.target.value})} style={inputStyle}>{TM_DOC_TYPES.map(d=><option key={d}>{d}</option>)}</select></label><label style={formLabel}>Reference / number<input value={docDraft.reference} onChange={e=>setDocDraft({...docDraft,reference:e.target.value})} style={inputStyle}/></label></div><input type="file" accept="image/*,application/pdf" onChange={e=>setDocDraft({...docDraft,file:e.target.files?.[0]||null})} style={{color:C.sub,marginBottom:14}}/><br/><button onClick={onUpload} disabled={loading} style={primaryBtn}>{loading?"Uploading…":"Upload Document"}</button></section></>}
    {view==="invoices"&&<><div style={{background:C.blue+"12",border:`1px solid ${C.blue}44`,borderRadius:9,padding:"11px 14px",marginBottom:14,color:C.sub,fontSize:12}}>Optional: use this only when the approved change-order amount appears on a trade's regular progress invoice.</div><table style={tableStyle}><thead><tr>{["Progress invoice #","Date","T&M / sheets referenced","Amount billed","Variance","Notes",""].map(h=><th key={h} style={thStyle}>{h}</th>)}</tr></thead><tbody>{(item.invoices||[]).map(v=><tr key={v.id}><td style={tdStyle}>{v.invoice_number}</td><td style={tdStyle}>{v.date||"—"}</td><td style={tdStyle}>{v.timesheet_refs||"—"}</td><td style={tdNum}>{money(v.amount)}</td><td style={{...tdNum,color:Math.abs((+v.amount||0)-supported)>.02?C.red:C.green}}>{money((+v.amount||0)-supported)}</td><td style={tdStyle}>{v.notes||"—"}</td><td style={tdStyle}><span style={{display:"flex",gap:4}}><button onClick={()=>setEditingInvoice({...v})} style={miniEdit}>Edit</button><button onClick={()=>{if(window.confirm(`Delete progress billing reference ${v.invoice_number}?`))onSave({...item,invoices:item.invoices.filter(x=>x.id!==v.id)})}} style={miniDanger}>Delete</button></span></td></tr>)}</tbody></table><section style={{...panel,marginTop:18}}><div style={panelTitle}>Link progress billing</div><div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>{inf("Progress invoice number","invoice_number")}{inf("Invoice date","date","date")}{inf("T&M amount billed before HST","amount","number")}</div>{inf("T&M item / sheet references","timesheet_refs")}{inf("Notes","notes")}<button onClick={onAddInvoice} style={primaryBtn}>Link Progress Billing</button></section></>}
    {editingSheet&&<TimesheetEditor draft={editingSheet} setDraft={setEditingSheet} onClose={()=>setEditingSheet(null)} onSave={()=>{onSave({...item,timesheets:item.timesheets.map(x=>x.id===editingSheet.id?editingSheet:x)});setEditingSheet(null)}}/>}
    {editingInvoice&&<RecordEditor title="Edit Progress Billing" draft={editingInvoice} setDraft={setEditingInvoice} fields={[["Progress invoice number","invoice_number"],["Invoice date","date","date"],["T&M amount billed before HST","amount","number"],["T&M / sheets referenced","timesheet_refs"],["Notes","notes"]]} onClose={()=>setEditingInvoice(null)} onSave={()=>{onSave({...item,invoices:item.invoices.map(x=>x.id===editingInvoice.id?editingInvoice:x)});setEditingInvoice(null)}}/>}
    {editingDoc&&<RecordEditor title="Edit Document Details" draft={editingDoc} setDraft={setEditingDoc} fields={[["File name","filename"],["Reference / number","reference"]]} select={["Document type","type",TM_DOC_TYPES.filter(x=>x!=="Auto-detect")]} onClose={()=>setEditingDoc(null)} onSave={()=>{onSave({...item,documents:item.documents.map(x=>x.id===editingDoc.id?editingDoc:x)});setEditingDoc(null)}}/>}
  </Modal>
}

const tmMoneyValue=t=>(+t.labour||0)+(+t.equipment||0)+(+t.material||0)+(+t.subcontractor||0)+(+t.markup||0)+(+t.total||0);
const tmIsPriced=t=>t?.source_type==="priced"||tmMoneyValue(t)>0||(!t?.signed&&!!t?.priced_sheet_number);
const tmIsField=t=>t?.source_type==="field"||!!t?.signed||!tmIsPriced(t);
const tmHours=r=>(+r?.regular_hours||0)+(+r?.overtime_hours||0);
const tmPairFor=(field,priced)=>{const refs=sheetNumbersForDisplay(field);return priced.find(p=>sheetNumbersForDisplay(p).some(n=>refs.includes(n)))||priced.find(p=>field.date&&p.date===field.date)||null};
const sheetNumbersForDisplay=t=>[t?.sheet_number,t?.priced_sheet_number].map(x=>String(x||"").trim()).filter(Boolean);
const matchRows=(fieldRows,pricedRows,key)=>fieldRows.map((f,i)=>({field:f,priced:pricedRows.find(p=>cleanTM(p[key])===cleanTM(f[key]))||pricedRows[i]||null}));
const cleanTM=s=>String(s||"").toLowerCase().replace(/[^a-z0-9]+/g," ").trim();

function TMFieldSheetView({item}){
  const sheets=(item.timesheets||[]).filter(tmIsField);
  if(!sheets.length)return <Empty text="No signed field sheet has been uploaded for this T&M."/>;
  return <>{sheets.map(s=><section key={s.id} style={{...panel,marginBottom:14}}><div style={{display:"flex",justifyContent:"space-between",gap:10,marginBottom:14}}><div><div style={panelTitle}>Field sheet #{s.sheet_number||"—"}</div><div style={{color:C.muted,fontSize:12}}>{s.date||"No date"} · {s.description||item.description||"No description"}</div></div><Badge color={s.signed?C.green:C.yellow}>{s.signed?"Signed":"Signature not recorded"}</Badge></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:18}}><div><div style={panelTitle}>Labour hours</div><table style={tableStyle}><thead><tr>{["Name","Classification","Regular","OT","Total"].map(h=><th key={h} style={thStyle}>{h}</th>)}</tr></thead><tbody>{(s.labour_entries||[]).map(r=><tr key={r.id}><td style={tdStyle}>{r.name||"—"}</td><td style={tdStyle}>{r.classification||"—"}</td><td style={tdNum}>{+r.regular_hours||0}</td><td style={tdNum}>{+r.overtime_hours||0}</td><td style={tdNum}>{tmHours(r)}</td></tr>)}</tbody></table></div><div><div style={panelTitle}>Equipment hours</div><table style={tableStyle}><thead><tr>{["Equipment","Machine #","Hours"].map(h=><th key={h} style={thStyle}>{h}</th>)}</tr></thead><tbody>{(s.equipment_entries||[]).map(r=><tr key={r.id}><td style={tdStyle}>{r.type||"—"}</td><td style={tdStyle}>{r.machine_number||"—"}</td><td style={tdNum}>{+r.hours||0}</td></tr>)}</tbody></table></div></div></section>)}</>;
}

function TMPricedSheetView({item}){
  const sheets=(item.timesheets||[]).filter(tmIsPriced);
  if(!sheets.length)return <Empty text="No priced breakdown has been uploaded yet. Use Upload Pricing & Documents to add it."/>;
  return <>{sheets.map(s=><section key={s.id} style={{...panel,marginBottom:14}}><div style={{display:"flex",justifyContent:"space-between",gap:10,marginBottom:14}}><div><div style={panelTitle}>Priced breakdown #{s.priced_sheet_number||s.sheet_number||"—"}</div><div style={{color:C.muted,fontSize:12}}>{s.date||"No date"} · {s.description||item.description||"No description"}</div></div><b style={{color:C.green,fontSize:18}}>{money(s.total||tmMoneyValue(s))}</b></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:18}}><div><div style={panelTitle}>Priced labour</div><table style={tableStyle}><thead><tr>{["Name","Class","Hours","Rate","Amount"].map(h=><th key={h} style={thStyle}>{h}</th>)}</tr></thead><tbody>{(s.labour_entries||[]).map(r=><tr key={r.id}><td style={tdStyle}>{r.name||"—"}</td><td style={tdStyle}>{r.classification||"—"}</td><td style={tdNum}>{tmHours(r)}</td><td style={tdNum}>{r.rate===""||r.rate==null?"—":money(r.rate)}</td><td style={tdNum}>{r.amount===""||r.amount==null?"—":money(r.amount)}</td></tr>)}</tbody></table></div><div><div style={panelTitle}>Priced equipment</div><table style={tableStyle}><thead><tr>{["Equipment","Machine #","Hours","Rate","Amount"].map(h=><th key={h} style={thStyle}>{h}</th>)}</tr></thead><tbody>{(s.equipment_entries||[]).map(r=><tr key={r.id}><td style={tdStyle}>{r.type||"—"}</td><td style={tdStyle}>{r.machine_number||"—"}</td><td style={tdNum}>{+r.hours||0}</td><td style={tdNum}>{r.rate===""||r.rate==null?"—":money(r.rate)}</td><td style={tdNum}>{r.amount===""||r.amount==null?"—":money(r.amount)}</td></tr>)}</tbody></table></div></div><div style={{display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:8,marginTop:16}}>{[["Labour",s.labour],["Equipment",s.equipment],["Materials",s.material],["Subtrades",s.subcontractor],["Markup",s.markup],["Total",s.total]].map(([a,b])=><div key={a} style={{background:C.bg,borderRadius:8,padding:10}}><small style={{color:C.muted}}>{a}</small><div style={{fontWeight:800,marginTop:4}}>{money(b)}</div></div>)}</div></section>)}</>;
}

function TMHoursComparison({item}){
  const fields=(item.timesheets||[]).filter(tmIsField),priced=(item.timesheets||[]).filter(tmIsPriced);
  if(!fields.length)return <Empty text="Upload the signed field sheet before comparing hours."/>;
  return <>{fields.map(f=>{const p=tmPairFor(f,priced);if(!p)return <section key={f.id} style={panel}><div style={{color:C.yellow}}>⚠ No priced breakdown is paired with field sheet #{f.sheet_number||"—"} yet.</div></section>;const labour=matchRows(f.labour_entries||[],p.labour_entries||[],"name"),equipment=matchRows(f.equipment_entries||[],p.equipment_entries||[],"machine_number"),fieldLabour=(f.labour_entries||[]).reduce((s,r)=>s+tmHours(r),0),pricedLabour=(p.labour_entries||[]).reduce((s,r)=>s+tmHours(r),0),fieldEquip=(f.equipment_entries||[]).reduce((s,r)=>s+(+r.hours||0),0),pricedEquip=(p.equipment_entries||[]).reduce((s,r)=>s+(+r.hours||0),0);return <section key={f.id} style={{...panel,marginBottom:14}}><div style={{display:"flex",justifyContent:"space-between",gap:10,marginBottom:14}}><div><div style={panelTitle}>Field #{f.sheet_number||"—"} versus priced #{p.priced_sheet_number||p.sheet_number||"—"}</div><div style={{color:C.muted,fontSize:12}}>{f.date||p.date||"No date"}</div></div><Badge color={fieldLabour===pricedLabour&&fieldEquip===pricedEquip?C.green:C.red}>{fieldLabour===pricedLabour&&fieldEquip===pricedEquip?"Hours match":"Hours differ"}</Badge></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:18}}><div><div style={panelTitle}>Labour comparison</div><table style={tableStyle}><thead><tr>{["Worker","Field","Priced","Difference"].map(h=><th key={h} style={thStyle}>{h}</th>)}</tr></thead><tbody>{labour.map(({field:r,priced:q},i)=>{const a=tmHours(r),b=q?tmHours(q):0,d=b-a;return <tr key={r.id||i}><td style={tdStyle}>{r.name||q?.name||"—"}</td><td style={tdNum}>{a}</td><td style={tdNum}>{q?b:"Missing"}</td><td style={{...tdNum,color:d===0?C.green:C.red}}>{d===0?"✓ 0":d>0?`+${d}`:d}</td></tr>})}<tr><td style={{...tdStyle,fontWeight:800}}>Total</td><td style={tdNum}>{fieldLabour}</td><td style={tdNum}>{pricedLabour}</td><td style={{...tdNum,color:pricedLabour-fieldLabour===0?C.green:C.red}}>{pricedLabour-fieldLabour}</td></tr></tbody></table></div><div><div style={panelTitle}>Equipment comparison</div><table style={tableStyle}><thead><tr>{["Equipment","Field","Priced","Difference"].map(h=><th key={h} style={thStyle}>{h}</th>)}</tr></thead><tbody>{equipment.map(({field:r,priced:q},i)=>{const a=+r.hours||0,b=q?(+q.hours||0):0,d=b-a;return <tr key={r.id||i}><td style={tdStyle}>{r.type||q?.type||"—"}</td><td style={tdNum}>{a}</td><td style={tdNum}>{q?b:"Missing"}</td><td style={{...tdNum,color:d===0?C.green:C.red}}>{d===0?"✓ 0":d>0?`+${d}`:d}</td></tr>})}<tr><td style={{...tdStyle,fontWeight:800}}>Total</td><td style={tdNum}>{fieldEquip}</td><td style={tdNum}>{pricedEquip}</td><td style={{...tdNum,color:pricedEquip-fieldEquip===0?C.green:C.red}}>{pricedEquip-fieldEquip}</td></tr></tbody></table></div></div></section>})}</>;
}

function TMMatchBanner({draft,setDraft,items}){
  if(!draft.linked_item_id)return null;
  return <div style={{background:C.green+"14",border:`1px solid ${C.green}66`,borderRadius:10,padding:"12px 14px",marginBottom:14,display:"flex",justifyContent:"space-between",alignItems:"center",gap:12,flexWrap:"wrap"}}><div><b style={{color:C.green}}>✓ Matched to {draft.linked_item_id}</b><div style={{color:C.sub,fontSize:11,marginTop:3}}>Matched by {draft.match_reason}. Confirming will update the existing T&M and preserve its signed field hours.</div></div><select value={draft.linked_item_id} onChange={e=>setDraft({...draft,linked_item_id:e.target.value,match_reason:e.target.value?"manual selection":""})} style={{...inputStyle,width:280}}><option value="">Create a new T&M instead</option>{items.map(i=><option key={i.id} value={i.id}>{i.id} · {i.title}</option>)}</select></div>;
}

function TMIntakeReview({draft,setDraft,items,notice,onBack,onConfirm}){
  const f=(label,key,type="text")=><label style={formLabel}>{label}<input type={type} value={draft[key]||""} onChange={e=>setDraft({...draft,[key]:e.target.value})} style={inputStyle}/></label>;
  const updateSheet=(id,key,value)=>setDraft({...draft,timesheets:draft.timesheets.map(t=>t.id===id?{...t,[key]:value}:t)});
  const removeSheet=id=>setDraft({...draft,timesheets:draft.timesheets.filter(t=>t.id!==id)});
  const addBlankSheet=()=>setDraft({...draft,timesheets:[...draft.timesheets,{id:Date.now(),date:"",sheet_number:"",priced_sheet_number:"",description:"",labour:0,equipment:0,material:0,subcontractor:0,markup:0,total:0,signed:false,labour_entries:[],equipment_entries:[],material_entries:[]}]});
  return <><TMMatchBanner draft={draft} setDraft={setDraft} items={items}/>{notice&&<div style={{background:"#451a03",border:`1px solid ${C.yellow}`,color:"#fde68a",borderRadius:9,padding:"10px 13px",marginBottom:15,fontSize:12}}>⚠ {notice}</div>}<div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(150px,1fr))",gap:10}}>{f("Project","project")}{f("Trade / contractor","trade")}{f("T&M scope / title","title")}{f("Location","location")}{f("Written direction","direction_ref")}{f("RCO number","rco_number")}{f("Change order number","change_order_number")}<label style={formLabel}>Stage<select value={draft.stage} onChange={e=>setDraft({...draft,stage:e.target.value})} style={inputStyle}>{TM_STAGES.map(s=><option key={s}>{s}</option>)}</select></label></div><label style={formLabel}>Work description<textarea value={draft.description||""} onChange={e=>setDraft({...draft,description:e.target.value})} style={{...inputStyle,minHeight:65,resize:"vertical"}}/></label><section style={{...panel,marginTop:12,padding:14}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:10,marginBottom:10}}><div><div style={panelTitle}>Field sheets and pricing read from form</div><div style={{color:C.muted,fontSize:11}}>Check worker hours, equipment hours, sheet numbers and priced amounts against the original.</div></div><button onClick={addBlankSheet} style={ghostBtn}>+ Add Missing Sheet</button></div>{draft.timesheets.length?draft.timesheets.map(t=><div key={t.id} style={{border:`1px solid ${C.border}`,borderRadius:10,padding:12,marginBottom:12}}><div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(130px,1fr))",gap:8}}>{[["Date","date","date"],["Field sheet #","sheet_number","text"],["Priced sheet #","priced_sheet_number","text"],["Description","description","text"],["Labour cost","labour","number"],["Equipment cost","equipment","number"],["Materials cost","material","number"],["Subtrades","subcontractor","number"],["Markup","markup","number"],["Total","total","number"]].map(([label,key,type])=><label key={key} style={formLabel}>{label}<input type={type} value={t[key]??""} onChange={e=>updateSheet(t.id,key,e.target.value)} style={inputStyle}/></label>)}</div><FieldRecords sheet={t} editable onChange={updated=>setDraft({...draft,timesheets:draft.timesheets.map(x=>x.id===t.id?updated:x)})}/><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:10}}><label style={{display:"flex",gap:8,color:C.sub,fontSize:12}}><input type="checkbox" checked={!!t.signed} onChange={e=>updateSheet(t.id,"signed",e.target.checked)}/>Signed field sheet is present</label><button onClick={()=>removeSheet(t.id)} style={miniDanger}>Remove Sheet</button></div></div>):<div style={{color:C.muted,padding:"20px 0",textAlign:"center"}}>No daily rows were found. Add a missing sheet or confirm the header details and add it later.</div>}</section>{draft.invoices?.length>0&&<section style={{...panel,marginTop:12,padding:14}}><div style={panelTitle}>Progress invoice found</div>{draft.invoices.map(v=><div key={v.id} style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 2fr",gap:10}}><label style={formLabel}>Invoice #<input value={v.invoice_number||""} onChange={e=>setDraft({...draft,invoices:draft.invoices.map(x=>x.id===v.id?{...x,invoice_number:e.target.value}:x)})} style={inputStyle}/></label><label style={formLabel}>Date<input type="date" value={v.date||""} onChange={e=>setDraft({...draft,invoices:draft.invoices.map(x=>x.id===v.id?{...x,date:e.target.value}:x)})} style={inputStyle}/></label><label style={formLabel}>Amount before HST<input type="number" value={v.amount??""} onChange={e=>setDraft({...draft,invoices:draft.invoices.map(x=>x.id===v.id?{...x,amount:e.target.value}:x)})} style={inputStyle}/></label><label style={formLabel}>Timesheets covered<input value={v.timesheet_refs||""} onChange={e=>setDraft({...draft,invoices:draft.invoices.map(x=>x.id===v.id?{...x,timesheet_refs:e.target.value}:x)})} style={inputStyle}/></label></div>)}</section>}<div style={{background:C.blue+"12",border:`1px solid ${C.blue}44`,borderRadius:9,padding:"10px 13px",marginTop:14,color:C.sub,fontSize:12}}>Source file: <b style={{color:C.text}}>{draft.documents?.[0]?.filename||"Uploaded form"}</b>. The original stays attached to this T&M item for later checking.</div><div style={modalActions}><button onClick={onConfirm} style={primaryBtn}>{draft.linked_item_id ? `Confirm & Update ${draft.linked_item_id}` : "Confirm & Add New T&M"}</button><button onClick={onBack} style={ghostBtn}>Choose Different Form</button></div></>
}

function TimesheetEditor({draft,setDraft,onClose,onSave}){const f=(label,key,type="text")=><label style={formLabel}>{label}<input type={type} value={draft[key]??""} onChange={e=>setDraft({...draft,[key]:e.target.value})} style={inputStyle}/></label>;return <div style={{position:"fixed",inset:0,zIndex:700,background:"#000c",display:"grid",placeItems:"center",padding:20}}><div style={{...panel,width:"100%",maxWidth:900,maxHeight:"90vh",overflowY:"auto"}}><div style={{fontWeight:850,fontSize:18,marginBottom:16}}>Edit Field Sheet / Priced T&M</div><div style={{display:"grid",gridTemplateColumns:"repeat(3,minmax(0,1fr))",gap:10}}>{f("Work date","date","date")}{f("Field sheet #","sheet_number")}{f("Priced sheet #","priced_sheet_number")}{f("Description","description")}{f("Labour cost","labour","number")}{f("Equipment cost","equipment","number")}{f("Materials cost","material","number")}{f("Subcontractors","subcontractor","number")}{f("Markup","markup","number")}{f("Total","total","number")}</div><FieldRecords sheet={draft} editable onChange={setDraft}/><label style={{display:"flex",gap:8,color:C.sub,fontSize:12,marginTop:14}}><input type="checkbox" checked={!!draft.signed} onChange={e=>setDraft({...draft,signed:e.target.checked})}/>Signed field sheet is present</label><div style={modalActions}><button onClick={onSave} style={primaryBtn}>Save Changes</button><button onClick={onClose} style={ghostBtn}>Cancel</button></div></div></div>}

function TMItemEditor({draft,setDraft,onSave}){const f=(label,key,options)=><label style={formLabel}>{label}{options?<select value={draft[key]||""} onChange={e=>setDraft({...draft,[key]:e.target.value})} style={inputStyle}>{options.map(x=><option key={x}>{x}</option>)}</select>:<input value={draft[key]||""} onChange={e=>setDraft({...draft,[key]:e.target.value})} style={inputStyle}/>}</label>;return <>{f("Project","project")}{f("Trade / contractor","trade")}{f("T&M title / scope","title")}{f("Location","location")}{f("Stage","stage",TM_STAGES)}{f("Written direction reference","direction_ref")}{f("RCO number","rco_number")}{f("Change order number","change_order_number")}<label style={formLabel}>Description<textarea value={draft.description||""} onChange={e=>setDraft({...draft,description:e.target.value})} style={{...inputStyle,minHeight:70}}/></label><label style={formLabel}>Notes<textarea value={draft.notes||""} onChange={e=>setDraft({...draft,notes:e.target.value})} style={{...inputStyle,minHeight:60}}/></label><button onClick={onSave} style={primaryBtn}>Save Changes</button></>}

function RecordEditor({title,draft,setDraft,fields,checkbox,select,onClose,onSave}){return <div style={{position:"fixed",inset:0,zIndex:700,background:"#000c",display:"grid",placeItems:"center",padding:20}}><div style={{...panel,width:"100%",maxWidth:600,maxHeight:"88vh",overflowY:"auto"}}><div style={{fontWeight:850,fontSize:18,marginBottom:16}}>{title}</div>{select&&<label style={formLabel}>{select[0]}<select value={draft[select[1]]||""} onChange={e=>setDraft({...draft,[select[1]]:e.target.value})} style={inputStyle}>{select[2].map(x=><option key={x}>{x}</option>)}</select></label>}<div style={{display:"grid",gridTemplateColumns:"repeat(2,minmax(0,1fr))",gap:10}}>{fields.map(([label,key,type="text"])=><label key={key} style={formLabel}>{label}<input type={type} value={draft[key]??""} onChange={e=>setDraft({...draft,[key]:e.target.value})} style={inputStyle}/></label>)}</div>{checkbox&&<label style={{display:"flex",gap:8,color:C.sub,fontSize:12,marginBottom:14}}><input type="checkbox" checked={!!draft[checkbox[1]]} onChange={e=>setDraft({...draft,[checkbox[1]]:e.target.checked})}/>{checkbox[0]}</label>}<div style={modalActions}><button onClick={onSave} style={primaryBtn}>Save Changes</button><button onClick={onClose} style={ghostBtn}>Cancel</button></div></div></div>}

// ═══════════════════════════════════════════════════════════════════════════════
// ROOT — ROUTER
// ═══════════════════════════════════════════════════════════════════════════════
export default function App() {
  const [module, setModule] = useState(null);
  if (module === "concrete")  return <ConcreteModule  onBack={() => setModule(null)} />;
  if (module === "certs")     return <CertsModule     onBack={() => setModule(null)} />;
  if (module === "tradedocs") return <TradeDocsModule onBack={() => setModule(null)} />;
  if (module === "tm")        return <TMModule        onBack={() => setModule(null)} />;
  return <LandingScreen onSelect={setModule} />;
}
