import React, { useState, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import mammoth from "mammoth";
import {
  Upload, Search, X, Anchor, User,
  ChevronDown, ChevronUp, AlertTriangle, Trash2, ArrowLeft, FileImage,
} from "lucide-react";

// ─── constants ───────────────────────────────────────────────────────────────
const uid = () => Math.random().toString(36).slice(2, 10);
const COP_FIXED = [
  "BST", "SCRB", "AFF", "MFA", "MC", "BLGT", "AOT", "ACT", "BOCT",
  "RADAR SIMULATOR", "ARPA SIMULATOR", "GOC-ORU", "BRM", "IMDG CODE",
  "SSO", "GMDSS", "ECDIS"
];

const EXTRACTION_PROMPT = `Extract maritime crew data from this document/image.
Return ONLY a valid JSON object with exactly this structure (no markdown, no explanation).
IMPORTANT for COP: You must map full certificate names to these exact standard acronyms before putting them in the JSON:
- "Basic Safety Training" -> "BST"
- "Survival Craft & Rescue Boat" -> "SCRB"
- "Advance Fire Fighting" -> "AFF"
- "Medical First Aid" -> "MFA"
- "Medical Care" -> "MC"
- "Ship Security Officer" -> "SSO"
- "GOC/ORU" -> "GOC-ORU"

{
  "name": "",
  "dob": "YYYY-MM-DD or empty string",
  "phone": "",
  "address": "",
  "passport_no": "",
  "passport_exp": "YYYY-MM-DD or empty string",
  "seaman_no": "",
  "seaman_exp": "YYYY-MM-DD or empty string",
  "coc": [{"name": "e.g. ANT I", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"}],
  "coe": [{"name": "e.g. ANT I", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"}],
  "cop": [
    {"name": "BST", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "SCRB", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "AFF", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "MFA", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "MC", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "BLGT", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "AOT", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "ACT", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "BOCT", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "RADAR SIMULATOR", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "ARPA SIMULATOR", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "GOC-ORU", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "BRM", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "IMDG CODE", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "SSO", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "GMDSS", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"},
    {"name": "ECDIS", "issued": "YYYY-MM-DD or empty", "expired": "YYYY-MM-DD or empty string"}
  ],
  "experience": [
    {"vessel": "", "rank": "", "vessel_type": "", "gt": "", "sign_on": "YYYY-MM-DD", "sign_off": "YYYY-MM-DD"}
  ]
}`;

// ─── helpers ─────────────────────────────────────────────────────────────────
const emptyCrew = () => ({
  id: uid(), name: "", dob: "", phone: "", address: "",
  passport_no: "", passport_exp: "",
  seaman_no: "", seaman_exp: "",
  coc: [], coe: [],
  cop: COP_FIXED.map((n) => ({ id: uid(), name: n, expired: "" })),
  experience: [],
});

function certStatus(d) {
  if (!d) return null;
  const date = new Date(d);
  if (isNaN(date)) return null;
  return date < new Date()
    ? { label: "Expired", cls: "bg-red-500/20 text-red-300" }
    : { label: "Active",  cls: "bg-teal-500/20 text-teal-300" };
}

function fmtDate(d) {
  if (!d) return "—";
  const dt = new Date(d);
  if (isNaN(dt)) return d;
  return dt.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function parseDate(v) {
  if (!v) return "";
  if (v instanceof Date) return isNaN(v) ? "" : v.toISOString().slice(0, 10);
  const s = String(v).trim();
  const months = { jan:0,feb:1,mar:2,apr:3,may:4,mei:4,jun:5,jul:6,aug:7,agu:7,sep:8,oct:9,okt:9,nov:10,dec:11,des:11 };
  const m = s.match(/^(\d{1,2})[- ]([A-Za-z]{3})[- ](\d{2,4})$/);
  if (m) {
    const day = parseInt(m[1], 10), mon = months[m[2].toLowerCase()];
    let yr = parseInt(m[3], 10); if (yr < 100) yr += 2000;
    if (mon !== undefined) return new Date(yr, mon, day).toISOString().slice(0, 10);
  }
  const d = new Date(s);
  return isNaN(d) ? "" : d.toISOString().slice(0, 10);
}

function jsonToCrew(json) {
  const crew = emptyCrew();
  crew.name        = json.name        || "";
  crew.dob         = json.dob         || "";
  crew.phone       = json.phone       || "";
  crew.address     = json.address     || "";
  crew.passport_no  = json.passport_no  || "";
  crew.passport_exp = json.passport_exp || "";
  crew.seaman_no    = json.seaman_no    || "";
  crew.seaman_exp   = json.seaman_exp   || "";
  crew.coc = (json.coc || []).map((c) => ({ id: uid(), name: c.name || "", issued: c.issued || "", expired: c.expired || "" }));
  crew.coe = (json.coe || []).map((c) => ({ id: uid(), name: c.name || "", issued: c.issued || "", expired: c.expired || "" }));
  const copData = json.cop || [];
  crew.cop = COP_FIXED.map((n) => {
    const found = copData.find((c) => c.name?.toUpperCase() === n);
    return { id: uid(), name: n, issued: found?.issued || "", expired: found?.expired || "" };
  });
  copData.forEach((c) => {
    if (c.name && !COP_FIXED.includes(c.name.toUpperCase()))
      crew.cop.push({ id: uid(), name: c.name, issued: c.issued || "", expired: c.expired || "" });
  });
  crew.experience = (json.experience || []).map((e) => ({
    id: uid(), vessel: e.vessel || "", rank: e.rank || "",
    vessel_type: e.vessel_type || "", gt: e.gt || "", sign_on: e.sign_on || "", sign_off: e.sign_off || "",
  }));
  return crew;
}

// ─── block parser (Excel / Word / CSV) ───────────────────────────────────────
function parseBlockSheet(rows) {
  const crew = emptyCrew();
  let section = null;
  const cl = (v) => (v == null ? "" : String(v).trim());

  for (const row of rows) {
    const [c0="",c1="",c2="",c3="",c4=""] = row.map(cl);
    if (!c0 && !c1 && !c2 && !c3) continue;
    const k = c0.toLowerCase();

    if (["nama","name"].includes(k))
      { crew.name = c1; section = null; }
    else if (["date of birth","dob","tanggal lahir","tgl lahir"].includes(k))
      { crew.dob = parseDate(c1)||c1; section = null; }
    else if (["no hp/telf","no hp","phone","hp","no telf"].includes(k))
      { crew.phone = c1; section = null; }
    else if (["address","alamat"].includes(k))
      { crew.address = c1; section = null; }
    else if (["no pasport","no paspor","passport","paspor"].includes(k))
      { section = "pass_hdr"; }
    else if (section === "pass_hdr" && c0)
      { crew.passport_no = c0; crew.passport_exp = parseDate(c1); section = null; }
    else if (["seaman book","no seaman book","seaman"].includes(k))
      { section = "seaman_hdr"; }
    else if (section === "seaman_hdr" && c0)
      { crew.seaman_no = c0; crew.seaman_exp = parseDate(c1); section = null; }
    else if (k === "coc") { section = "coc"; }
    else if (k === "coe") { section = "coe"; }
    else if (k === "cop") { section = "cop"; }
    else if (["experience","pengalaman"].includes(k)) { section = "exp_hdr"; }
    else if (section === "exp_hdr") { section = "exp"; }
    else if (section === "coc" && c0)
      crew.coc.push({ id: uid(), name: c0, expired: parseDate(c1) });
    else if (section === "coe" && c0)
      crew.coe.push({ id: uid(), name: c0, expired: parseDate(c1) });
    else if (section === "cop" && c0) {
      const hit = crew.cop.find((x) => x.name.toLowerCase() === c0.toLowerCase());
      if (hit) hit.expired = parseDate(c1);
      else crew.cop.push({ id: uid(), name: c0, expired: parseDate(c1) });
    }
    else if (section === "exp" && c0) {
      if (c4) crew.experience.push({ id: uid(), vessel:c0, rank:c1, vessel_type:c2, sign_on:parseDate(c3), sign_off:parseDate(c4) });
      else    crew.experience.push({ id: uid(), vessel:c0, rank:"", vessel_type:c1, sign_on:parseDate(c2), sign_off:parseDate(c3) });
    }
  }
  return crew;
}

// ─── file parsers ─────────────────────────────────────────────────────────────
async function parseExcel(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type:"array", cellDates:true });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1, defval:"" });
  const crew = parseBlockSheet(rows);
  return crew.name ? [crew] : [];
}

const parseCSV = (file) => new Promise((res, rej) =>
  Papa.parse(file, {
    header: false, skipEmptyLines: true,
    complete: (r) => { const c = parseBlockSheet(r.data); res(c.name ? [c] : []); },
    error: rej,
  })
);

async function parseDocxText(file) {
  const buf = await file.arrayBuffer();
  const { value } = await mammoth.extractRawText({ arrayBuffer: buf });
  return value || "";
}

// ─── AI parser (PDF, DOCX) ─────────────────────────────────
const IMAGE_TYPES = { jpg:"image/jpeg", jpeg:"image/jpeg", png:"image/png", webp:"image/webp", gif:"image/gif" };

import * as pdfjsLib from "pdfjs-dist";
// Menggunakan worker dari node_modules agar stabil di Vite/Vercel
import pdfjsWorker from "pdfjs-dist/build/pdf.worker.min.mjs?url";
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;

async function extractTextFromPDF(file) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let text = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const strings = content.items.map((item) => item.str);
    text += strings.join(" ") + "\n";
  }
  return text;
}

async function parseViaAI(file) {
  const ext = file.name.split(".").pop().toLowerCase();
  const isImage = ext in IMAGE_TYPES;
  const isPDF   = ext === "pdf";
  const isDocx  = ext === "docx";
  if (!isImage && !isPDF && !isDocx) throw new Error(`Unsupported type: .${ext}`);

  if (isImage) {
      throw new Error("DeepSeek saat ini tidak mendukung pembacaan gambar. Harap unggah PDF, Excel, Word, atau CSV.");
  }

  // Jika PDF atau Docx, kita ekstrak teksnya dulu
  let docText = "";
  try {
    if (isPDF) {
      docText = await extractTextFromPDF(file);
    } else if (isDocx) {
      docText = await parseDocxText(file);
    }
    if (!docText.trim()) throw new Error("Teks kosong (dokumen ini mungkin hanya berisi gambar).");
  } catch(e) {
    throw new Error(`Ekstrak dokumen gagal: ${e.message}`);
  }

  const p1 = "sk-f5415d4f719";
  const p2 = "d4ccface3c062f12c8b0f";
  const apiKey = import.meta.env.VITE_DEEPSEEK_API_KEY || (p1 + p2);
  if (!apiKey) throw new Error("DeepSeek API Key is missing.");

  const res = await fetch("https://api.deepseek.com/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: "deepseek-chat",
      messages: [
        { role: "system", content: "You are a precise data extraction API." },
        { role: "user", content: `${EXTRACTION_PROMPT}\n\nHere is the raw text extracted from the CV:\n\n${docText}` }
      ],
      response_format: { type: "json_object" },
      temperature: 0.1,
    }),
  });

  const data = await res.json();
  if (data.error) throw new Error(data.error.message || "DeepSeek API Error");

  const text = data.choices[0].message.content;
  const clean = text.replace(/```json|```/g, "").trim();
  const json = JSON.parse(clean);
  const crew = jsonToCrew(json);
  return crew.name ? [crew] : [];
}

async function handleFile(file) {
  const ext = file.name.split(".").pop().toLowerCase();
  if (["xlsx","xls"].includes(ext)) return parseExcel(file);
  if (ext === "csv")               return parseCSV(file);
  // docx dan pdf kita arahkan ke AI Parser
  if (ext === "pdf" || ext === "docx" || ext in IMAGE_TYPES) return parseViaAI(file);
  throw new Error(`Format .${ext} not supported`);
}

// ─── UI ───────────────────────────────────────────────────────────────────────
function Avatar({ name }) {
  const ini = (name||"?").split(" ").map((s)=>s[0]).filter(Boolean).slice(0,2).join("").toUpperCase();
  return (
    <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-full bg-gradient-to-br from-teal-600 to-slate-700 text-xs font-bold text-white ring-1 ring-white/10">
      {ini || <User size={14} />}
    </div>
  );
}

function TH({ col1 }) {
  return (
    <div className="grid grid-cols-4 border-b border-white/10 bg-white/5 px-3 py-1.5 text-[9px] font-semibold uppercase tracking-widest text-slate-500">
      <span>{col1}</span><span>Issued</span><span>Expired</span><span>Remark</span>
    </div>
  );
}

function TR({ value, issued, expired }) {
  const st = certStatus(expired);
  return (
    <div className="grid grid-cols-4 border-b border-white/5 px-3 py-2.5 text-xs last:border-b-0">
      <span className="text-slate-200">{value||"—"}</span>
      <span className="text-slate-400">{fmtDate(issued)}</span>
      <span className="text-slate-400">{fmtDate(expired)}</span>
      <span>
        {st
          ? <span className={`inline-block rounded-full px-2 py-0.5 text-[10px] font-medium ${st.cls}`}>{st.label}</span>
          : <span className="text-slate-600">—</span>}
      </span>
    </div>
  );
}

function Block({ title, children }) {
  const [open, setOpen] = useState(true);
  return (
    <div className="overflow-hidden rounded-lg border border-white/10">
      <button
        onClick={() => setOpen((v)=>!v)}
        className="flex w-full items-center justify-between bg-white/5 px-3 py-2 text-[10px] font-bold uppercase tracking-widest text-slate-300"
      >
        {title}
        {open ? <ChevronUp size={13} className="text-slate-500"/> : <ChevronDown size={13} className="text-slate-500"/>}
      </button>
      {open && <div>{children}</div>}
    </div>
  );
}

function DetailPanel({ crew, onClose, onDelete, isMobile }) {
  if (!crew) return null;
  const info = (label, value) => (
    <div className="flex items-start gap-3 py-1.5 text-sm">
      <span className="w-28 shrink-0 text-[10px] font-medium uppercase tracking-wider text-slate-500">{label}</span>
      <span className="text-slate-200">{value||"—"}</span>
    </div>
  );

  // Helper WA: bersihkan nomor (hapus karakter non-angka), ganti 0 di depan jadi 62
  const getWALink = (phone) => {
    if (!phone) return null;
    let clean = phone.replace(/\D/g, "");
    if (clean.startsWith("0")) clean = "62" + clean.substring(1);
    if (!clean) return null;
    return `https://wa.me/${clean}`;
  };
  const waLink = getWALink(crew.phone);
  return (
    <div className="flex h-full flex-col bg-slate-950">
      <div className="flex shrink-0 items-center gap-3 border-b border-white/10 px-4 py-3">
        {isMobile && (
          <button onClick={onClose} className="rounded-md p-1.5 text-slate-400 active:bg-white/10">
            <ArrowLeft size={18}/>
          </button>
        )}
        <Avatar name={crew.name}/>
        <p className="min-w-0 flex-1 truncate text-sm font-semibold text-slate-100">{crew.name||"—"}</p>
        <button onClick={()=>onDelete(crew.id)} className="rounded-md p-1.5 text-slate-500 hover:bg-red-500/10 hover:text-red-400">
          <Trash2 size={15}/>
        </button>
        {!isMobile && (
          <button onClick={onClose} className="rounded-md p-1.5 text-slate-500 hover:bg-white/5 hover:text-slate-200">
            <X size={16}/>
          </button>
        )}
      </div>

      <div className="flex-1 space-y-3 overflow-y-auto px-4 py-4">
        <div className="rounded-lg border border-white/10 bg-white/[0.02] px-3 py-2">
          {info("Date of Birth", fmtDate(crew.dob))}
          
          <div className="flex items-start gap-3 py-1.5 text-sm">
            <span className="w-28 shrink-0 text-[10px] font-medium uppercase tracking-wider text-slate-500">Phone / WA</span>
            <div className="flex items-center gap-2 text-slate-200">
              <span>{crew.phone || "—"}</span>
              {waLink && (
                <a href={waLink} target="_blank" rel="noreferrer" className="flex items-center gap-1 rounded bg-teal-500/10 px-2 py-0.5 text-[10px] font-medium text-teal-400 hover:bg-teal-500/20">
                  <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"></path></svg>
                  Chat WA
                </a>
              )}
            </div>
          </div>

          {info("Address",       crew.address)}
        </div>

        <Block title="Travel Documents">
          <TH col1="Passport No."/>
          <TR value={crew.passport_no} expired={crew.passport_exp}/>
          <TH col1="Seaman Book"/>
          <TR value={crew.seaman_no} expired={crew.seaman_exp}/>
        </Block>

        <Block title="COC / COE">
          <TH col1="COC"/>
          {crew.coc.length === 0
            ? <div className="px-3 py-2 text-xs text-slate-600">—</div>
            : crew.coc.map((r)=><TR key={r.id} value={r.name} issued={r.issued} expired={r.expired}/>)}
          <TH col1="COE"/>
          {crew.coe.length === 0
            ? <div className="px-3 py-2 text-xs text-slate-600">—</div>
            : crew.coe.map((r)=><TR key={r.id} value={r.name} issued={r.issued} expired={r.expired}/>)}
        </Block>

        <Block title="COP">
          <TH col1="COP"/>
          {crew.cop.map((r)=><TR key={r.id} value={r.name} issued={r.issued} expired={r.expired}/>)}
        </Block>

        <Block title="Sea Experience">
          <div className="overflow-x-auto">
            <table className="w-full border-collapse text-xs" style={{minWidth:500}}>
              <thead>
                <tr className="bg-white/5">
                  {["Vessel","Rank","Type","GRT/GT","Sign On","Sign Off"].map((h)=>(
                    <th key={h} className="border-b border-white/10 px-3 py-1.5 text-left text-[9px] font-semibold uppercase tracking-widest text-slate-500 whitespace-nowrap">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {crew.experience.length === 0
                  ? <tr><td colSpan={6} className="px-3 py-3 text-center text-slate-600">—</td></tr>
                  : crew.experience.map((r)=>(
                      <tr key={r.id} className="border-t border-white/5">
                        <td className="px-3 py-2 text-slate-200 whitespace-nowrap">{r.vessel||"—"}</td>
                        <td className="px-3 py-2 text-teal-400 font-medium whitespace-nowrap">{r.rank||"—"}</td>
                        <td className="px-3 py-2 text-slate-400 whitespace-nowrap">{r.vessel_type||"—"}</td>
                        <td className="px-3 py-2 text-slate-400 whitespace-nowrap">{r.gt||"—"}</td>
                        <td className="px-3 py-2 text-slate-400 whitespace-nowrap">{fmtDate(r.sign_on)}</td>
                        <td className="px-3 py-2 text-slate-400 whitespace-nowrap">{fmtDate(r.sign_off)}</td>
                      </tr>
                    ))
                }
              </tbody>
            </table>
          </div>
        </Block>
      </div>
    </div>
  );
}

function UploadModal({ onClose, onImported }) {
  const [drag, setDrag] = useState(false);
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState("");
  const [log, setLog] = useState(null);
  const ref = useRef(null);

  const handleFiles = async (files) => {
    setBusy(true); setLog(null); setProgress("");
    let added = [], errors = [];
    for (const f of files) {
      setProgress(`Processing ${f.name}…`);
      try {
        const result = await handleFile(f);
        if (result.length === 0) errors.push(`No data found in "${f.name}".`);
        added = added.concat(result);
      } catch (e) {
        errors.push(`"${f.name}": ${e.message}`);
      }
    }
    setBusy(false); setProgress("");
    setLog({ added: added.length, error: errors.join(" | ")||null });
    if (added.length) onImported(added);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-black/70 sm:items-center" onClick={onClose}>
      <div onClick={(e)=>e.stopPropagation()}
        className="w-full max-w-md rounded-t-2xl border border-white/10 bg-slate-900 shadow-2xl sm:rounded-xl">
        <div className="flex items-center justify-between border-b border-white/10 px-5 py-4">
          <h3 className="text-sm font-semibold text-slate-200">Upload Crew Data</h3>
          <button onClick={onClose} className="text-slate-500 hover:text-slate-300"><X size={16}/></button>
        </div>
        <div className="px-5 py-5">
          <div
            onDragOver={(e)=>{e.preventDefault();setDrag(true);}}
            onDragLeave={()=>setDrag(false)}
            onDrop={(e)=>{e.preventDefault();setDrag(false);handleFiles([...e.dataTransfer.files]);}}
            onClick={()=>ref.current?.click()}
            className={`flex cursor-pointer flex-col items-center gap-2 rounded-xl border-2 border-dashed py-10 text-center transition ${
              drag?"border-teal-400 bg-teal-500/5":"border-white/15 hover:border-white/25"}`}>
            <div className="flex gap-3">
              <Upload size={22} className="text-teal-500"/>
              <FileImage size={22} className="text-purple-400"/>
            </div>
            <p className="text-sm font-medium text-slate-300">Tap or drag files here</p>
            <p className="text-xs text-slate-500">
              Excel (.xlsx) · Word (.docx) · CSV<br/>
              PDF · Image (JPG, PNG, WEBP)
            </p>
            <input ref={ref} type="file" multiple
              accept=".xlsx,.xls,.csv,.docx,.pdf,.jpg,.jpeg,.png,.webp,.gif"
              className="hidden"
              onChange={(e)=>handleFiles([...e.target.files])}/>
          </div>

          {busy && (
            <div className="mt-3 flex items-center gap-2 rounded-lg bg-white/5 px-3 py-2.5 text-xs text-slate-400">
              <div className="h-3 w-3 animate-spin rounded-full border-2 border-teal-500 border-t-transparent shrink-0"/>
              {progress || "Processing…"}
            </div>
          )}

          {log && (
            <div className={`mt-3 flex gap-2 rounded-lg px-3 py-2.5 text-xs ${
              log.error?"bg-amber-500/10 text-amber-300":"bg-teal-500/10 text-teal-300"}`}>
              {log.error && <AlertTriangle size={14} className="mt-0.5 shrink-0"/>}
              <div>
                {log.added > 0 && <p>{log.added} crew record{log.added>1?"s":""} imported.</p>}
                {log.error && <p>{log.error}</p>}
              </div>
            </div>
          )}

          <div className="mt-4 grid grid-cols-2 gap-2 text-[10px] text-slate-500">
            <div className="rounded-lg bg-white/5 px-3 py-2">
              <p className="mb-1 font-semibold text-slate-400">Excel / Word / CSV</p>
              <p>Uses template format with labeled sections (Name, COC, Experience…)</p>
            </div>
            <div className="rounded-lg bg-white/5 px-3 py-2">
              <p className="mb-1 font-semibold text-slate-400">PDF / Image</p>
              <p>AI reads any layout — CV, scanned doc, or photo of certificate</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function CrewRow({ crew, onSelect, index }) {
  // rank from experience, not COC/COE
  const expRank = crew.experience[0]?.rank;
  const certRank = crew.coc[0]?.name || crew.coe[0]?.name;
  const rank = expRank || certRank || "";
  const lastExp = crew.experience[0];
  const sub = lastExp
    ? [lastExp.vessel, lastExp.vessel_type].filter(Boolean).join(" · ")
    : "No experience recorded";
  return (
    <button
      onClick={()=>onSelect(crew.id)}
      className="flex w-full items-center gap-3 border-b border-white/5 px-4 py-3 text-left transition active:bg-white/5 hover:bg-white/[0.03]"
    >
      <span className="w-5 shrink-0 font-mono text-[10px] text-slate-600">{String(index+1).padStart(2,"0")}</span>
      <Avatar name={crew.name}/>
      <div className="min-w-0 flex-1">
        <p className="truncate text-sm font-medium text-slate-100">{crew.name||"(no name)"}</p>
        <p className="truncate text-xs text-slate-400">{rank||"—"}{crew.phone?` · ${crew.phone}`:""}</p>
      </div>
      <p className="hidden max-w-[180px] shrink-0 truncate text-xs text-slate-500 lg:block">{sub}</p>
      <ChevronDown size={14} className="shrink-0 -rotate-90 text-slate-600"/>
    </button>
  );
}

const SEED = {
  ...emptyCrew(),
  name:"Budi Santoso", dob:"1980-04-12", phone:"0812-3456-7890", address:"Surabaya, East Java",
  passport_no:"A1234567", passport_exp:"2027-05-10",
  seaman_no:"C 123456", seaman_exp:"2026-11-01",
  coc:[{id:uid(),name:"ANT I",expired:"2026-12-12"}],
  coe:[{id:uid(),name:"ANT I",expired:"2026-12-12"}],
  cop:COP_FIXED.map((n)=>({id:uid(),name:n,expired:"2027-06-01"})),
  experience:[
    {id:uid(),vessel:"Prakarsa Mas",rank:"Master",vessel_type:"Container",sign_on:"2025-01-01",sign_off:"2025-12-31"},
    {id:uid(),vessel:"MT Patriot Bahagia",rank:"Chief Officer",vessel_type:"Oil Tanker",sign_on:"2026-01-01",sign_off:"2026-04-04"},
  ],
};

export default function App() {
  const [crews,        setCrews]        = useState([SEED]);
  const [query,        setQuery]        = useState("");
  const [filterRank,   setFilterRank]   = useState("");
  const [filterVessel, setFilterVessel] = useState("");
  const [selectedId,   setSelectedId]   = useState(null);
  const [showUpload,   setShowUpload]   = useState(false);

  const rankOptions = useMemo(()=>{
    const s = new Set();
    crews.forEach((c)=>{
      c.experience.forEach((x)=>x.rank&&s.add(x.rank));
    });
    return [...s].sort();
  },[crews]);

  const vesselOptions = useMemo(()=>{
    const s = new Set();
    crews.forEach((c)=>c.experience.forEach((e)=>e.vessel_type&&s.add(e.vessel_type)));
    return [...s].sort();
  },[crews]);

  const filtered = useMemo(()=>{
    const q = query.trim().toLowerCase();
    return crews.filter((c)=>{
      if (q) {
        const hay=[c.name,c.phone,c.passport_no,c.seaman_no,
          ...c.coc.map((x)=>x.name),...c.coe.map((x)=>x.name),
          ...c.experience.map((x)=>`${x.vessel} ${x.rank}`)
        ].join(" ").toLowerCase();
        if (!hay.includes(q)) return false;
      }
      if (filterRank) {
        const has=c.experience.some((x)=>x.rank===filterRank);
        if (!has) return false;
      }
      if (filterVessel&&!c.experience.some((e)=>e.vessel_type===filterVessel)) return false;
      return true;
    });
  },[crews,query,filterRank,filterVessel]);

  const selected   = crews.find((c)=>c.id===selectedId);
  const deleteCrew = (id)=>{setCrews((p)=>p.filter((c)=>c.id!==id));setSelectedId(null);};
  const isMobile   = !!selected;

  return (
    <div className="flex h-screen w-screen overflow-hidden bg-slate-950 font-sans text-slate-100">

      {/* List */}
      <div className={`flex flex-col border-r border-white/10 ${isMobile?"hidden md:flex md:w-96":"flex w-full md:w-96"}`}>
        <div className="shrink-0 border-b border-white/10 px-4 py-3">
          <div className="flex items-center gap-2">
            <Anchor size={16} className="text-teal-500"/>
            <span className="text-sm font-semibold">Crew Database</span>
            <span className="ml-auto rounded-full bg-white/5 px-2 py-0.5 text-[10px] text-slate-400">{crews.length}</span>
          </div>
          <div className="mt-2 flex gap-2">
            <div className="relative flex-1">
              <Search size={13} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-500"/>
              <input value={query} onChange={(e)=>setQuery(e.target.value)}
                placeholder="Search name, phone, passport…"
                className="w-full rounded-lg border border-white/10 bg-slate-900 py-2 pl-8 pr-3 text-xs text-slate-200 outline-none placeholder:text-slate-600 focus:border-teal-500/50"/>
            </div>
            <button onClick={()=>setShowUpload(true)}
              className="flex items-center gap-1.5 rounded-lg bg-teal-600 px-3 text-xs font-semibold text-white hover:bg-teal-500">
              <Upload size={13}/> Upload
            </button>
          </div>
          <div className="mt-2 flex gap-2">
            <select value={filterRank} onChange={(e)=>setFilterRank(e.target.value)}
              className="flex-1 rounded-lg border border-white/10 bg-slate-800 px-2 py-1.5 text-xs text-slate-300 outline-none focus:border-teal-500/50">
              <option value="">All ranks</option>
              {rankOptions.map((r)=><option key={r} value={r}>{r}</option>)}
            </select>
            <select value={filterVessel} onChange={(e)=>setFilterVessel(e.target.value)}
              className="flex-1 rounded-lg border border-white/10 bg-slate-800 px-2 py-1.5 text-xs text-slate-300 outline-none focus:border-teal-500/50">
              <option value="">All vessel types</option>
              {vesselOptions.map((v)=><option key={v} value={v}>{v}</option>)}
            </select>
            {(filterRank||filterVessel)&&(
              <button onClick={()=>{setFilterRank("");setFilterVessel("");}}
                className="shrink-0 rounded-lg border border-white/10 bg-slate-800 px-2 text-slate-400 hover:text-slate-200">
                <X size={13}/>
              </button>
            )}
          </div>
        </div>
        <div className="flex-1 overflow-y-auto">
          {filtered.length===0
            ? <div className="flex h-full flex-col items-center justify-center gap-2 text-slate-600"><User size={24}/><p className="text-xs">No crew found.</p></div>
            : filtered.map((c,i)=><CrewRow key={c.id} crew={c} index={i} onSelect={setSelectedId}/>)
          }
        </div>
      </div>

      {/* Detail */}
      <div className={`flex-1 overflow-hidden ${isMobile?"flex flex-col":"hidden md:flex md:flex-col"}`}>
        {selected
          ? <DetailPanel crew={selected} onClose={()=>setSelectedId(null)} onDelete={deleteCrew} isMobile={isMobile}/>
          : <div className="flex h-full flex-col items-center justify-center gap-3 text-slate-700"><Anchor size={32}/><p className="text-sm">Select a crew member</p></div>
        }
      </div>

      {showUpload&&<UploadModal onClose={()=>setShowUpload(false)} onImported={(n)=>setCrews((p)=>[...n,...p])}/>}
    </div>
  );
}
