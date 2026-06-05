import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── COLORS ───────────────────────────────────────────────────────────────────
const C = {
  bg: "#0A0F1E", card: "#111827", card2: "#161f30", border: "#1F2937",
  accent: "#F97316", blue: "#3B82F6", green: "#22C55E",
  yellow: "#EAB308", red: "#EF4444", purple: "#A855F7",
  muted: "#6B7280", text: "#F9FAFB", sub: "#9CA3AF", teal: "#14B8A6",
};

// ─── UPSTASH STORAGE HELPERS ─────────────────────────────────────────────────
// Uses Upstash Redis REST API — set these two env vars in your Vercel project:
//   VITE_UPSTASH_REDIS_REST_URL  (e.g. https://xxxx.upstash.io)
//   VITE_UPSTASH_REDIS_REST_TOKEN
const UPSTASH_URL   = import.meta.env.VITE_UPSTASH_REDIS_REST_KV_REST_API_URL;
const UPSTASH_TOKEN = import.meta.env.VITE_UPSTASH_REDIS_REST_KV_REST_API_TOKEN;

async function upstashCmd(...args) {
  if (!UPSTASH_URL || !UPSTASH_TOKEN) return null;
  try {
    const res = await fetch(`${UPSTASH_URL}/${args.map(encodeURIComponent).join("/")}`, {
      headers: { Authorization: `Bearer ${UPSTASH_TOKEN}` }
    });
    const data = await res.json();
    return data.result ?? null;
  } catch { return null; }
}

async function storageGet(key) {
  try {
    const raw = await upstashCmd("GET", key);
    return raw ? JSON.parse(raw) : null;
  } catch { return null; }
}
async function storageSet(key, value) {
  try {
    await upstashCmd("SET", key, JSON.stringify(value));
    return true;
  } catch { return false; }
}
async function storageDel(key) {
  try {
    await upstashCmd("DEL", key);
    return true;
  } catch { return false; }
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

// ─── CONCRETE SCOPE ───────────────────────────────────────────────────────────
const SCOPE = [
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
const M3_TO_YD3 = 1.30795;
const AREAS = [...new Set(SCOPE.map(r => r.area))];
const ITEMS = [...new Set(SCOPE.map(r => r.item).filter(Boolean))];
const MPA_SPEC = {};
SCOPE.forEach(r => { if (r.area && r.item) MPA_SPEC[`${r.area}|||${r.item}`] = r.mpa; });

function parseMpaNum(str) {
  if (!str) return null;
  const m = String(str).match(/(\d+)\s*[Mm][Pp][Aa]/);
  if (m) return parseInt(m[1]);
  const n = parseFloat(str);
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
      id:"certs", emoji:"📋", title:"Training Certificates",
      desc:"Scan worker certifications, track expiry dates, get alerts before certifications lapse.",
      color: C.teal,
      stats: ["25 cert types","Expiry tracking","Worker roster","Email intake ready"],
    },
  ];

  return (
    <div style={{ minHeight:"100vh", background:C.bg, color:C.text, fontFamily:"'DM Sans','Segoe UI',sans-serif", display:"flex", flexDirection:"column" }}>
      <div style={{ background:C.card, borderBottom:`1px solid ${C.border}`, padding:"20px 32px", display:"flex", alignItems:"center", gap:14 }}>
        <div style={{ width:44, height:44, borderRadius:12, background:`linear-gradient(135deg,${C.accent},${C.teal})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:22 }}>⚡</div>
        <div>
          <div style={{ fontWeight:800, fontSize:19 }}>Site Document Intelligence</div>
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

        <div style={{ marginTop:48, background:C.card, border:`1px solid ${C.border}`, borderRadius:16, padding:"20px 28px", maxWidth:580, width:"100%", textAlign:"center" }}>
          <div style={{ color:C.muted, fontSize:12, fontWeight:700, letterSpacing:1, textTransform:"uppercase", marginBottom:8 }}>Coming Soon</div>
          <div style={{ display:"flex", gap:10, justifyContent:"center", flexWrap:"wrap" }}>
            {["📦 Delivery Slips","📝 RFI / Deficiency Logs","🦺 Safety Observations","🔍 Inspection Records"].map(m => (
              <span key={m} style={{ background:C.border+"44", color:C.muted, borderRadius:8, padding:"6px 14px", fontSize:12, fontWeight:600 }}>{m}</span>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// TRAINING CERTIFICATES MODULE
// ═══════════════════════════════════════════════════════════════════════════════
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

  // ── Load from storage on mount ──
  useEffect(() => {
    storageGet("certs-data").then(saved => {
      if (saved?.certs) setCerts(saved.certs);
      setStorageReady(true);
    });
  }, []);

  // ── Save to storage whenever certs change ──
  useEffect(() => {
    if (!storageReady) return;
    storageSet("certs-data", { certs });
  }, [certs, storageReady]);

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

    const res = await fetch("https://api.anthropic.com/v1/messages", {
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
    setCerts(prev => [...prev, { id:Date.now(), filename:"Manual entry", added_at:new Date().toISOString(), ...manual }]);
    setManual({ worker_name:"", employer:"", cert_type:"", issued_date:"", expiry_date:"", cert_number:"", notes:"" });
    setManualOpen(false);
    showToast("Certificate added ✓");
  }

  function exportXLSX() {
    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.json_to_sheet(certs.map((c,i) => {
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
    const expiring = certs.filter(c => { const d = daysUntilExpiry(c.expiry_date); return d !== null && d <= 60; })
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
    const workers = [...new Set(certs.map(c=>c.worker_name).filter(Boolean))];
    const rosterRows = workers.map(w => {
      const wCerts = certs.filter(c => c.worker_name === w);
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
  const expired  = certs.filter(c => { const d = daysUntilExpiry(c.expiry_date); return d!==null && d<0; });
  const critical = certs.filter(c => { const d = daysUntilExpiry(c.expiry_date); return d!==null && d>=0 && d<=30; });
  const warning  = certs.filter(c => { const d = daysUntilExpiry(c.expiry_date); return d!==null && d>30 && d<=60; });
  const workers  = [...new Set(certs.map(c=>c.worker_name).filter(Boolean))];

  // Filtered cert list
  const filteredCerts = certs.filter(c => {
    if (filterExpiry === "expired")  return daysUntilExpiry(c.expiry_date) !== null && daysUntilExpiry(c.expiry_date) < 0;
    if (filterExpiry === "critical") return daysUntilExpiry(c.expiry_date) !== null && daysUntilExpiry(c.expiry_date) >= 0 && daysUntilExpiry(c.expiry_date) <= 30;
    if (filterExpiry === "warning")  return daysUntilExpiry(c.expiry_date) !== null && daysUntilExpiry(c.expiry_date) > 30 && daysUntilExpiry(c.expiry_date) <= 60;
    return true;
  });

  // Worker detail modal
  function WorkerModal({ name, onClose }) {
    const wCerts = certs.filter(c => c.worker_name === name);
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
            <div style={{ fontWeight:800, fontSize:17 }}>Training Certificates</div>
            <div style={{ color:C.muted, fontSize:12 }}>{certs.length} certificates · {workers.length} workers{expired.length>0?` · ⚠ ${expired.length} expired`:""}{ critical.length>0?` · ⚠ ${critical.length} expiring soon`:""}</div>
          </div>
        </div>
        <div style={{ display:"flex", gap:8, alignItems:"center" }}>
          <span style={{ fontSize:11, color:storageReady?C.green:C.muted, fontWeight:700 }}>
            {storageReady ? "💾 Auto-saved" : "⏳ Loading..."}
          </span>
          <button onClick={async()=>{ if(!window.confirm("Clear ALL certificates? This cannot be undone.")) return; await storageDel("certs-data"); setCerts([]); showToast("All data cleared."); }} style={{ background:"transparent", color:C.red, border:`1px solid ${C.red}44`, borderRadius:7, padding:"5px 11px", fontSize:11, fontWeight:700, cursor:"pointer" }}>🗑 Clear Data</button>
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
        {TAB("certs",     `📋 Certificates (${certs.length})`)}
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
  const [storageReady, setStorageReady] = useState(false);
  const fileRef    = useRef();
  const invFileRef = useRef();

  // ── Load from storage on mount ──
  useEffect(() => {
    storageGet("concrete-data").then(saved => {
      if (saved?.tickets)  setTickets(saved.tickets);
      if (saved?.invoices) setInvoices(saved.invoices);
      setStorageReady(true);
    });
  }, []);

  // ── Save to storage whenever tickets or invoices change ──
  useEffect(() => {
    if (!storageReady) return;
    storageSet("concrete-data", { tickets, invoices });
  }, [tickets, invoices, storageReady]);

  function showToast(msg, type="ok") { setToast({msg,type}); setTimeout(()=>setToast(null),3500); }
  async function toB64(file) { return new Promise((res,rej)=>{ const r=new FileReader(); r.onload=()=>res(r.result.split(",")[1]); r.onerror=rej; r.readAsDataURL(file); }); }
  async function toDataURL(file) { return new Promise((res,rej)=>{ const r=new FileReader(); r.onload=()=>res(r.result); r.onerror=rej; r.readAsDataURL(file); }); }

  async function extractTicket(file) {
    const b64 = await toB64(file);
    const isPDF = file.type==="application/pdf";
    const block = isPDF ? {type:"document",source:{type:"base64",media_type:"application/pdf",data:b64}} : {type:"image",source:{type:"base64",media_type:file.type,data:b64}};
    const mpaRef = SCOPE.filter(r=>r.item).map(r=>`${r.area} ${r.item}: ${r.mpa}`).join(", ");
    const prompt = `You are a construction data extraction assistant. This file may contain ONE or MULTIPLE concrete delivery dockets/tickets. Extract ALL tickets found.
Project areas: ${AREAS.join(", ")}. Element types: ${ITEMS.join(", ")}.

CRITICAL FIELD EXTRACTION RULES — read carefully:

1. ticket_number: Read ONLY from the box explicitly labelled "TICKET NO", "NO. BILLET", or "TICKET NO / NO. BILLET". On Quality Concrete tickets this is near the bottom of the form above "AT PLANT TIME". It is typically a 7-8 digit number like 11477789. Do NOT use ORDER NO, PO number, or customer number.

2. mix_design (MOST IMPORTANT): You must find the concrete STRENGTH in MPa. Look for ANY of these:
   - A field labelled "PRODUCT CODE" or "CODE DU PRODUIT" — will contain codes like "Q35NA1A", "Q25NB1A", "HRWR10"
   - A field labelled "MIX DESIGN" or "DESIGNATION"
   - Text anywhere on the ticket showing a number followed by "MPA" or "MPa" such as "35MPA N 20MM", "25 MPA", "32MPa"
   - On Quality Concrete tickets, look in the middle section of the ticket for the product/mix code row
   - The strength will be a number like 25, 32, or 35 followed by MPa
   - DO NOT use "ULTRA SLUMP", slump values in mm, or admixture descriptions as the mix_design
   - If you find a product code like "Q35NA1A" extract it AND note the 35 MPa strength
   - Return the full mix description you find, e.g. "35 MPa N 20mm" or "Q35NA1A — 35 MPa"

3. volume_m3: Read from "QUANTITY" or "QUANTITE" field — a number like 8.00, 7.50 in m³. Do NOT use yd³ values here.

4. date: The delivery/load date on the ticket in YYYY-MM-DD format.

Return ONLY a valid JSON array (even if only one ticket). No markdown, no explanation:
[{"date":"YYYY-MM-DD","ticket_number":"7-8 digit from TICKET NO field","supplier":"supplier name","mix_design":"MPa strength and mix code e.g. 35 MPa N 20mm or Q35NA1A","volume_m3":number or null,"volume_yd3":number or null,"area":"best match from area list or null","item":"best match from element list or null","invoice_number":"string or null","driver":"string or null","truck_number":"string or null","notes":"string or null"}]`;
    const res = await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:4000,messages:[{role:"user",content:[block,{type:"text",text:prompt}]}]})});
    const data = await res.json();
    if(data.error) throw new Error("API error: "+(data.error.message||JSON.stringify(data.error)));
    const text = data.content?.map(b=>b.text||"").join("")||"";
    if(!text) throw new Error("Empty API response (HTTP "+res.status+")");
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
    const res = await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:1500,messages:[{role:"user",content:[block,{type:"text",text:prompt}]}]})});
    const data = await res.json();
    if(data.error) throw new Error("API error: "+(data.error.message||JSON.stringify(data.error)));
    const text = data.content?.map(b=>b.text||"").join("")||"";
    if(!text) throw new Error("Empty API response (HTTP "+res.status+")");
    return extractJSON(text);
  }

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
    const matched=[]; const unmatched=[];
    const ticketsOnInvoice = allTickets.filter(t=>{
      const tNum=String(t.ticket_number||"").trim().toLowerCase();
      const tInv=String(t.invoice_number||"").trim().toLowerCase();
      const invNum=String(invoice.invoice_number||"").trim().toLowerCase();
      return (invoiceTicketNums.includes(tNum)&&tNum!=="")||(tInv!==""&&invNum!==""&&tInv===invNum);
    });
    invoiceTicketNums.forEach(n=>{
      const found=allTickets.find(t=>String(t.ticket_number||"").trim().toLowerCase()===n);
      if(found) matched.push({invoiceRef:n,ticket:found}); else unmatched.push(n);
    });
    const ticketVolume=ticketsOnInvoice.reduce((s,t)=>s+(parseFloat(t.volume_m3)||0),0);
    const invoiceVolume=invoice.total_volume_m3||0;
    const volumeMatch=invoiceVolume>0?Math.abs(ticketVolume-invoiceVolume)<0.5:null;
    return {matched,unmatched,ticketsOnInvoice,ticketVolume,invoiceVolume,volumeMatch};
  }

  async function handleTicketFiles(files) {
    if(!files?.length) return; setLoading(true);
    const pending = [];
    for(const file of Array.from(files)){
      setLoadMsg(`Reading "${file.name}"…`);
      try{
        const [extractedArr,dataURL]=await Promise.all([extractTicket(file),toDataURL(file)]);
        for(const extracted of extractedArr){
          if(extracted.volume_m3&&!extracted.volume_yd3) extracted.volume_yd3=+(extracted.volume_m3*M3_TO_YD3).toFixed(3);
          if(extracted.volume_yd3&&!extracted.volume_m3) extracted.volume_m3=+(extracted.volume_yd3/M3_TO_YD3).toFixed(3);
          // Pass growing pending array so each ticket in batch accounts for previous ones
          const allSoFar = [...tickets, ...pending];
          const suggestion = suggestLocation(extracted.mix_design, allSoFar);
          extracted.area = suggestion.area || extracted.area || "";
          extracted.item = suggestion.item || extracted.item || "";
          pending.push({id:Date.now()+Math.random(),filename:file.name,fileType:file.type,originalFile:dataURL,added_at:new Date().toISOString(),...extracted,_suggested:!!(suggestion.area)});
        }
      }catch(e){ showToast(`Could not read "${file.name}": ${e.message}`,"err"); }
    }
    setLoading(false); setLoadMsg("");
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
        const [extracted,dataURL]=await Promise.all([extractInvoice(file),toDataURL(file)]);
        setInvoices(prev=>[...prev,{id:Date.now()+Math.random(),filename:file.name,fileType:file.type,originalFile:dataURL,added_at:new Date().toISOString(),...extracted}]);
        added++;
      }catch(e){ showToast(`Could not read invoice "${file.name}": ${e.message}`,"err"); }
    }
    setLoading(false); setLoadMsg(""); if(added){showToast(`${added} invoice${added>1?"s":""} scanned ✓`);setTab("invoices");}
  }

  const totalPoured=tickets.reduce((s,t)=>s+(parseFloat(t.volume_m3)||0),0);
  const totalYd3=tickets.reduce((s,t)=>s+(parseFloat(t.volume_yd3)||0),0);
  const remaining=Math.max(0,TOTAL_SCOPE_M3-totalPoured);
  const pct=(totalPoured/TOTAL_SCOPE_M3)*100;
  const mpaMismatches=tickets.filter(t=>checkMpaMismatch(t));
  const pouredMap={};
  tickets.forEach(t=>{ const key=`${t.area||"Unknown"}|||${t.item||""}`; pouredMap[key]=(pouredMap[key]||0)+(parseFloat(t.volume_m3)||0); });
  const areaTotals={};
  SCOPE.forEach(r=>{ if(!areaTotals[r.area]) areaTotals[r.area]={scope:0,poured:0}; areaTotals[r.area].scope+=r.m3; });
  Object.entries(pouredMap).forEach(([key,val])=>{ const area=key.split("|||")[0]; if(!areaTotals[area]) areaTotals[area]={scope:0,poured:0}; areaTotals[area].poured+=val; });
  const invoicesWithIssues=invoices.filter(inv=>{ const m=matchInvoiceToTickets(inv,tickets); return m.unmatched.length>0||m.volumeMatch===false; }).length;
  const totalInvoiced=invoices.reduce((s,inv)=>s+(parseFloat(inv.total_amount)||0),0);

  function exportXLSX() {
    const wb=XLSX.utils.book_new();
    const ws1=XLSX.utils.json_to_sheet(tickets.map((t,i)=>{
      const mismatch=checkMpaMismatch(t);
      return {"#":i+1,"Date":t.date||"","Ticket #":t.ticket_number||"","Supplier":t.supplier||"","Mix Design (Ticket)":t.mix_design||"","Spec MPa":t.area&&t.item?(MPA_SPEC[`${t.area}|||${t.item}`]||""):"","MPa Status":mismatch?`⚠ MISMATCH (spec: ${mismatch.specMpa})`:t.mix_design?"✓ OK":"—","Area":t.area||"","Element":t.item||"","Volume (m³)":parseFloat(t.volume_m3)||"","Volume (yd³)":parseFloat(t.volume_yd3)||"","Invoice #":t.invoice_number||"","Driver":t.driver||"","Truck #":t.truck_number||"","Notes":t.notes||""};
    }));
    ws1["!cols"]=[4,12,16,22,18,16,20,14,14,14,14,14,14,12,22].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb,ws1,"Ticket Log");
    const ws2=XLSX.utils.json_to_sheet(SCOPE.map(r=>{ const poured=pouredMap[`${r.area}|||${r.item}`]||0; const rem=Math.max(0,r.m3-poured); return {"Area":r.area,"Element":r.item,"Spec MPa":r.mpa||"","Scope (m³)":r.m3,"Poured (m³)":poured||"","Remaining (m³)":rem||"","% Complete":r.m3>0?((poured/r.m3)*100).toFixed(1)+"%":"0%"}; }));
    ws2["!cols"]=[14,22,14,14,14,16,12].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb,ws2,"Progress by Element");
    if(mpaMismatches.length>0){ const wsm=XLSX.utils.json_to_sheet(mpaMismatches.map(t=>({ "Ticket #":t.ticket_number||"","Date":t.date||"","Area":t.area||"","Element":t.item||"","Ticket Mix":t.mix_design||"","Spec MPa":MPA_SPEC[`${t.area}|||${t.item}`]||"","Supplier":t.supplier||"","Volume (m³)":parseFloat(t.volume_m3)||"" }))); wsm["!cols"]=[14,12,14,16,20,16,20,14].map(w=>({wch:w})); XLSX.utils.book_append_sheet(wb,wsm,"⚠ MPa Mismatches"); }
    XLSX.writeFile(wb,`concrete-tracker-${new Date().toISOString().slice(0,10)}.xlsx`);
    showToast("Spreadsheet downloaded ✓");
  }

  function addManual() {
    if(!manual.ticket_number&&!manual.date){showToast("Enter at least a date or ticket number.","err");return;}
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

  const TAB=(t,label)=>(<button onClick={()=>setTab(t)} style={{padding:"8px 16px",borderRadius:8,cursor:"pointer",fontWeight:700,fontSize:12,border:"none",background:tab===t?C.accent:"transparent",color:tab===t?"#fff":C.muted,transition:"all .15s",whiteSpace:"nowrap"}}>{label}</button>);
  const INPUT=(key,label,type="text",opts=null)=>(<div style={{marginBottom:12}}><label style={{display:"block",color:C.muted,fontSize:10,fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:4}}>{label}</label>{opts?<select value={manual[key]} onChange={e=>setManual(m=>({...m,[key]:e.target.value}))} style={{width:"100%",background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,padding:"9px 12px",color:C.text,fontSize:14,boxSizing:"border-box"}}><option value="">— select —</option>{opts.map(o=><option key={o} value={o}>{o}</option>)}</select>:<input type={type} value={manual[key]} onChange={e=>setManual(m=>({...m,[key]:e.target.value}))} style={{width:"100%",background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,padding:"9px 12px",color:C.text,fontSize:14,boxSizing:"border-box"}}/>}</div>);

  function TicketModal({ticket,onClose}){
    const mismatch=checkMpaMismatch(ticket);
    const specMpa=ticket.area&&ticket.item?MPA_SPEC[`${ticket.area}|||${ticket.item}`]:null;
    return(<div style={{position:"fixed",inset:0,background:"#000c",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={{background:C.card,border:`1px solid ${mismatch?C.red:C.border}`,borderRadius:16,padding:28,width:"94%",maxWidth:520,maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <div><div style={{fontWeight:800,fontSize:18}}>🧾 Ticket #{ticket.ticket_number||"—"}</div><div style={{color:C.muted,fontSize:13}}>{ticket.supplier} · {ticket.date}</div></div>
          <Badge color={mismatch?C.red:C.green}>{mismatch?"⚠ MPa Mismatch":"✓ OK"}</Badge>
        </div>
        {mismatch&&<div style={{background:"#450a0a",border:`1px solid ${C.red}`,borderRadius:10,padding:"14px 16px",marginBottom:16}}><div style={{color:C.red,fontWeight:800,fontSize:14,marginBottom:6}}>⚠ Mix Design Mismatch — Do Not Pour</div><div style={{color:"#fca5a5",fontSize:13}}>Ticket shows <b>{mismatch.ticketMpa}</b> but this element requires <b>{mismatch.specMpa}</b>.<br/>Verify with supplier before proceeding.</div></div>}
        <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:16}}>
          {ticket.volume_m3&&<div style={{background:C.bg,borderRadius:10,padding:"12px 16px",flex:1}}><div style={{color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase"}}>Volume</div><div style={{color:C.accent,fontWeight:800,fontSize:20,fontFamily:"monospace"}}>{parseFloat(ticket.volume_m3).toFixed(2)} m³</div></div>}
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
        <button onClick={onClose} style={{background:C.bg,color:C.muted,border:`1px solid ${C.border}`,borderRadius:9,padding:"10px 20px",fontWeight:700,cursor:"pointer",width:"100%"}}>Close</button>
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
      {selectedTicket&&<TicketModal ticket={selectedTicket} onClose={()=>setSelectedTicket(null)}/>}
      {selectedInvoice&&<InvoiceModal invoice={selectedInvoice} onClose={()=>setSelectedInvoice(null)}/>}

      <div style={{background:C.card,borderBottom:`1px solid ${C.border}`,padding:"16px 28px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12}}>
        <div style={{display:"flex",alignItems:"center",gap:13}}>
          <button onClick={onBack} style={{background:"transparent",border:`1px solid ${C.border}`,color:C.muted,borderRadius:8,padding:"6px 13px",fontWeight:700,fontSize:12,cursor:"pointer"}}>← Back</button>
          <div style={{width:40,height:40,borderRadius:11,background:C.accent,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20}}>🏗️</div>
          <div>
            <div style={{fontWeight:800,fontSize:17}}>Concrete Tracker</div>
            <div style={{color:C.muted,fontSize:12}}>9133.5 m³ scope · {tickets.length} tickets · {invoices.length} invoices{mpaMismatches.length>0?` · ⚠ ${mpaMismatches.length} MPa mismatch${mpaMismatches.length>1?"es":""}`:""}</div>
          </div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          <span style={{fontSize:11,color:storageReady?C.green:C.muted,fontWeight:700}}>
            {storageReady ? "💾 Auto-saved" : "⏳ Loading..."}
          </span>
          <button onClick={async()=>{ if(!window.confirm("Clear ALL tickets and invoices? This cannot be undone.")) return; await storageDel("concrete-data"); setTickets([]); setInvoices([]); showToast("All data cleared."); }} style={{background:"transparent",color:C.red,border:`1px solid ${C.red}44`,borderRadius:7,padding:"5px 11px",fontSize:11,fontWeight:700,cursor:"pointer"}}>🗑 Clear Data</button>
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
              <Stat label="Tickets"        value={tickets.length}             sub="dockets scanned"                    color={C.blue}/>
              <Stat label="MPa Mismatches" value={mpaMismatches.length}       sub={mpaMismatches.length>0?"⚠ review required":"✓ all clear"} color={mpaMismatches.length>0?C.red:C.green}/>
              <Stat label="Invoices"       value={invoices.length}            sub={invoicesWithIssues>0?`⚠ ${invoicesWithIssues} need review`:invoices.length>0?"✓ all matched":"none yet"} color={invoicesWithIssues>0?C.red:C.purple}/>
            </div>
            <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:14,padding:"20px 24px",marginBottom:24}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}><span style={{fontWeight:700}}>Overall Progress</span><span style={{color:C.accent,fontWeight:800,fontFamily:"monospace"}}>{fmt(pct,1)}%</span></div>
              <Bar pct={pct} color={pct>=100?C.green:C.accent}/>
              <div style={{display:"flex",justifyContent:"space-between",marginTop:8,color:C.muted,fontSize:12}}><span>{fmt(totalPoured)} m³ poured</span><span>{fmt(TOTAL_SCOPE_M3,1)} m³ total scope</span></div>
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
                      onClick={()=>{
                        const mpaWarnings=[];
                        reviewQueue.forEach(t=>{const mm=checkMpaMismatch(t);if(mm)mpaWarnings.push(t);});
                        setTickets(prev=>[...prev,...reviewQueue]);
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
              <button onClick={()=>fileRef.current.click()} style={{background:C.accent,color:"#fff",border:"none",borderRadius:9,padding:"9px 18px",fontWeight:700,cursor:"pointer",fontSize:13}}>+ Scan Ticket</button>
              <input ref={fileRef} type="file" multiple accept="image/*,application/pdf" style={{display:"none"}} onChange={e=>handleTicketFiles(e.target.files)}/>
            </div>
            {tickets.length===0?<div style={{color:C.muted,textAlign:"center",padding:"60px 0"}}>No tickets yet.</div>
            :[...tickets].reverse().map(t=>{
              const mismatch=checkMpaMismatch(t);
              const specMpa=t.area&&t.item?MPA_SPEC[`${t.area}|||${t.item}`]:null;
              return(<div key={t.id} onClick={()=>setSelectedTicket(t)} style={{background:C.card,border:`1px solid ${mismatch?C.red+"88":C.border}`,borderRadius:12,padding:"15px 20px",marginBottom:12,cursor:"pointer"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8,flexWrap:"wrap",gap:8}}>
                  <div><span style={{fontWeight:800,fontSize:15}}>{t.ticket_number||"No ticket #"}</span><span style={{color:C.muted,fontSize:12,marginLeft:10}}>{t.date||"No date"}</span></div>
                  <div style={{display:"flex",gap:7,flexWrap:"wrap",alignItems:"center"}}>
                    {t.volume_m3&&<Badge color={C.accent}>{parseFloat(t.volume_m3).toFixed(2)} m³</Badge>}
                    {t.mix_design&&<Badge color={mismatch?C.red:C.green}>{t.mix_design}{mismatch?" ⚠":""}</Badge>}
                    {specMpa&&!mismatch&&<Badge color={C.muted}>spec: {specMpa}</Badge>}
                    <button onClick={e=>{ e.stopPropagation(); if(window.confirm(`Delete ticket #${t.ticket_number||"this ticket"}?`)) setTickets(prev=>prev.filter(x=>x.id!==t.id)); }} style={{background:"transparent",border:`1px solid ${C.red}44`,color:C.red,borderRadius:6,padding:"3px 9px",fontSize:12,fontWeight:700,cursor:"pointer"}}>✕</button>
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
        )}

        {tab==="invoices"&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:12}}>
              <div><div style={{fontWeight:700,fontSize:18}}>Invoice Matching</div><div style={{color:C.muted,fontSize:13,marginTop:2}}>Auto-matched against logged tickets</div></div>
              <button onClick={()=>invFileRef.current.click()} style={{background:C.purple,color:"#fff",border:"none",borderRadius:9,padding:"9px 18px",fontWeight:700,cursor:"pointer",fontSize:13}}>+ Scan Invoice</button>
              <input ref={invFileRef} type="file" multiple accept="image/*,application/pdf" style={{display:"none"}} onChange={e=>handleInvoiceFiles(e.target.files)}/>
            </div>
            {invoices.length===0?<div style={{color:C.muted,textAlign:"center",padding:"60px 0"}}>No invoices yet.</div>
            :invoices.map(inv=>{ const m=matchInvoiceToTickets(inv,tickets); const hasIssues=m.unmatched.length>0||m.volumeMatch===false;
              return(<div key={inv.id} onClick={()=>setSelectedInvoice(inv)} style={{background:C.card,border:`1px solid ${hasIssues?C.red+"66":C.border}`,borderRadius:12,padding:"16px 20px",marginBottom:12,cursor:"pointer"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10,flexWrap:"wrap",gap:8}}>
                  <div><span style={{fontWeight:800,fontSize:15}}>Invoice {inv.invoice_number||"—"}</span><span style={{color:C.muted,fontSize:12,marginLeft:10}}>{inv.invoice_date}</span></div>
                  <div style={{display:"flex",gap:7,flexWrap:"wrap",alignItems:"center"}}>{inv.total_amount>0&&<Badge color={C.green}>{inv.currency||""} {inv.total_amount?.toLocaleString()}</Badge>}<Badge color={hasIssues?C.red:C.green}>{hasIssues?"⚠ Review":"✓ Matched"}</Badge>
                    <button onClick={e=>{ e.stopPropagation(); if(window.confirm(`Delete invoice ${inv.invoice_number||"this invoice"}?`)) setInvoices(prev=>prev.filter(x=>x.id!==inv.id)); }} style={{background:"transparent",border:`1px solid ${C.red}44`,color:C.red,borderRadius:6,padding:"3px 9px",fontSize:12,fontWeight:700,cursor:"pointer"}}>✕</button>
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

        {tab==="remaining"&&(()=>{
          const rate=parseFloat(ratePerM3)||null;
          const invoicedVolume=invoices.reduce((s,inv)=>s+(parseFloat(inv.total_volume_m3)||0),0);
          const invoicedAmount=invoices.reduce((s,inv)=>s+(parseFloat(inv.total_amount)||0),0);
          const derivedRate=invoicedVolume>0&&invoicedAmount>0?invoicedAmount/invoicedVolume:null;
          const totalEstCost=rate?remaining*rate:null;
          const areaRows=AREAS.map(area=>{ const at=areaTotals[area]||{scope:0,poured:0}; const rem=Math.max(0,at.scope-at.poured); const p=at.scope>0?(at.poured/at.scope)*100:0; const status=p>=100?"complete":p>0?"inprogress":"notstarted"; const estCost=rate&&rem>0?rem*rate:null; const areaMpas=[...new Set(SCOPE.filter(r=>r.area===area&&r.mpa).map(r=>r.mpa))]; return{area,scope:at.scope,poured:at.poured,rem,p,status,estCost,areaMpas}; });
          const STATUS_CONFIG={complete:{label:"✅ Complete",color:C.green},inprogress:{label:"🟡 In Progress",color:C.yellow},notstarted:{label:"🔴 Not Started",color:C.red}};
          return(<div>
            <div style={{fontWeight:700,fontSize:18,marginBottom:6}}>Remaining Works</div>
            <div style={{color:C.muted,fontSize:13,marginBottom:22}}>Live picture of what's done, in progress, and still to pour</div>
            <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:24}}>
              {[{emoji:"✅",count:areaRows.filter(r=>r.status==="complete").length,label:"Complete",color:C.green},{emoji:"🟡",count:areaRows.filter(r=>r.status==="inprogress").length,label:"In Progress",color:C.yellow},{emoji:"🔴",count:areaRows.filter(r=>r.status==="notstarted").length,label:"Not Started",color:C.red},{emoji:"🏗️",count:fmt(remaining,1),label:"m³ left",color:C.accent}].map(({emoji,count,label,color})=>(<div key={label} style={{background:color+"18",border:`1px solid ${color}44`,borderRadius:12,padding:"14px 20px",flex:1,minWidth:110,textAlign:"center"}}><div style={{fontSize:28}}>{emoji}</div><div style={{fontWeight:800,fontSize:22,color}}>{count}</div><div style={{color:C.muted,fontSize:12,marginTop:2}}>{label}</div></div>))}
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
            {areaRows.map(row=>{ const sc=STATUS_CONFIG[row.status]; return(<div key={row.area} style={{background:C.card,border:`1px solid ${row.status==="complete"?C.green+"44":row.status==="inprogress"?C.yellow+"44":C.border}`,borderRadius:13,padding:"15px 20px",marginBottom:10}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8,flexWrap:"wrap",gap:8}}>
                <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}><span style={{fontWeight:800,fontSize:15}}>📍 {row.area}</span><Badge color={sc.color}>{sc.label}</Badge>{row.areaMpas.map(m=><MpaBadge key={m} mpa={m}/>)}</div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap"}}><Badge color={C.accent}>{fmt(row.p,1)}%</Badge>{row.estCost!==null&&<Badge color={C.green}>${row.estCost.toLocaleString(undefined,{maximumFractionDigits:0})} est.</Badge>}</div>
              </div>
              <Bar pct={row.p} color={row.status==="complete"?C.green:row.status==="inprogress"?C.yellow:C.border}/>
              <div style={{display:"flex",gap:20,marginTop:9,fontSize:13,flexWrap:"wrap"}}>
                <span style={{color:C.muted}}>Poured: <b style={{color:C.text}}>{fmt(row.poured)} m³</b></span>
                <span style={{color:C.muted}}>Scope: <b style={{color:C.text}}>{fmt(row.scope)} m³</b></span>
                {row.rem>0?<span style={{color:C.muted}}>Still to pour: <b style={{color:C.yellow}}>{fmt(row.rem)} m³</b></span>:<span style={{color:C.green,fontWeight:700}}>✓ All poured</span>}
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
              <thead><tr style={{color:C.muted,textAlign:"left"}}>{["Area","Element","Spec MPa","Scope (m³)","Poured (m³)","Remaining (m³)","% Done"].map(h=><th key={h} style={{padding:"8px 14px",borderBottom:`1px solid ${C.border}`,fontWeight:700,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
              <tbody>{SCOPE.map((r,i)=>{ const poured=pouredMap[`${r.area}|||${r.item}`]||0; const rem=Math.max(0,r.m3-poured); const p=r.m3>0?(poured/r.m3)*100:0; return(<tr key={i} style={{borderBottom:`1px solid ${C.border}22`,background:i%2===0?"transparent":C.card+"88"}}>
                <td style={{padding:"9px 14px",color:C.sub}}>{r.area}</td>
                <td style={{padding:"9px 14px"}}>{r.item}</td>
                <td style={{padding:"9px 14px"}}>{r.mpa?<MpaBadge mpa={r.mpa}/>:"—"}</td>
                <td style={{padding:"9px 14px",fontFamily:"monospace"}}>{r.m3.toFixed(1)}</td>
                <td style={{padding:"9px 14px",fontFamily:"monospace",color:poured>0?C.green:C.muted}}>{poured>0?poured.toFixed(2):"—"}</td>
                <td style={{padding:"9px 14px",fontFamily:"monospace",color:rem>0?C.yellow:C.green}}>{rem>0?rem.toFixed(2):"✓"}</td>
                <td style={{padding:"9px 14px"}}>{poured>0?<Badge color={p>=100?C.green:C.accent}>{fmt(p,0)}%</Badge>:<span style={{color:C.muted}}>—</span>}</td>
              </tr>); })}</tbody>
              <tfoot><tr style={{borderTop:`2px solid ${C.border}`,fontWeight:800}}>
                <td colSpan={3} style={{padding:"11px 14px"}}>TOTAL</td>
                <td style={{padding:"11px 14px",fontFamily:"monospace"}}>{fmt(TOTAL_SCOPE_M3,1)}</td>
                <td style={{padding:"11px 14px",fontFamily:"monospace",color:C.green}}>{totalPoured>0?fmt(totalPoured):"—"}</td>
                <td style={{padding:"11px 14px",fontFamily:"monospace",color:C.yellow}}>{fmt(remaining)}</td>
                <td style={{padding:"11px 14px"}}><Badge color={pct>=100?C.green:C.accent}>{fmt(pct,1)}%</Badge></td>
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
// ROOT — ROUTER
// ═══════════════════════════════════════════════════════════════════════════════
export default function App() {
  const [module, setModule] = useState(null);
  if (module === "concrete") return <ConcreteModule onBack={() => setModule(null)} />;
  if (module === "certs")    return <CertsModule    onBack={() => setModule(null)} />;
  return <LandingScreen onSelect={setModule} />;
}
