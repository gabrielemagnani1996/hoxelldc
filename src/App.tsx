import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { Upload, FileSpreadsheet, Copy, Download, AlertTriangle, CheckCircle2 } from "lucide-react";
import { Card, CardContent } from "@/components/ui/card";
import { Button } from "@/components/ui/button";

const PIPE = "||";
const NL = String.fromCharCode(10);

const DEPARTMENT_MAP = {
  housekeeping: "HOU",
  housekeepine: "HOU",
  hou: "HOU",
  maintenance: "MAIN",
  main: "MAIN",
  "front office": "REC",
  "front ofice": "REC",
  reception: "REC",
  rec: "REC",
  "food and beverage": "FB",
  "food & beverage": "FB",
  "f&b": "FB",
  fb: "FB",
  management: "MAN",
  direzione: "MAN",
};

const SECTIONS = [
  {
    id: "roomTypes",
    label: "Room Type / Categorie Camere",
    sourceSheet: "Categorie Camere",
    header: "pms_room_type_code||name||cleaning_time_empty||cleaning_time_arrival_or_departure||cleaning_time_in_house",
    start: 4,
    cols: { code: 1, name: 2, empty: 3, departure: 4, inHouse: 5 },
    build: (r) => [r.code, r.name, r.empty, r.departure, r.inHouse],
  },
  {
    id: "rooms",
    label: "Rooms / Camere",
    sourceSheet: "Camere",
    header: "name||pms_room_name||display_order||cleaning_time_empty||cleaning_time_arrival_or_departure||cleaning_time_in_house",
    start: 3,
    cols: { name: 1, pmsType: 2, building: 3, floor: 4, empty: 5, departure: 6, inHouse: 7 },
    build: (r, i) => [r.name, r.name, r.name || i + 1, r.empty, r.departure, r.inHouse],
  },
  {
    id: "commonAreas",
    label: "Common Area / Aree Comuni",
    sourceSheet: "Aree Comuni",
    header: "name||pms_room_name||display_order||cleaning_time_empty||cleaning_time_arrival_or_departure||cleaning_time_in_house",
    start: 3,
    cols: { name: 1, type: 2, section: 3 },
    build: (r, i) => [r.name, r.name, 100100 + i * 10, "", "", ""],
  },
  {
    id: "roomSections",
    label: "Room Sections / Sezioni Camere",
    sourceSheet: "Camere",
    header: "name",
    start: 3,
    cols: { section: 4 },
    unique: true,
    build: (r) => [r.section],
  },
  {
    id: "faultCategories",
    label: "Maintenance Categories / Categorie Guasti",
    sourceSheet: "Guasti",
    header: "name",
    start: 4,
    cols: { category: 3 },
    unique: true,
    build: (r) => [r.category],
  },
  {
    id: "faults",
    label: "Maintenance Faults / Guasti per Categoria",
    sourceSheet: "Guasti",
    header: "category||name||severity||rooms||common_areas||zone_details",
    start: 4,
    cols: { name: 1, severity: 2, category: 3, rooms: 4, commons: 5, zone: 6 },
    build: (r) => [r.category, r.name, severity(r.severity), checked(r.rooms), checked(r.commons), r.zone],
  },
  {
    id: "users",
    label: "Users / Utenti",
    sourceSheet: "Utenti",
    header: "first_name||last_name||username||pms_id||job_title||email||language||department_id",
    start: 3,
    cols: { firstName: 1, lastName: 2, department: 3, jobTitle: 4 },
    build: (r) => [r.firstName, r.lastName, username(r.firstName, r.lastName), "", r.jobTitle, "", "", depId(r.department)],
  },
  {
    id: "buildingAutomation",
    label: "Building Automation",
    sourceSheet: "Building Automation",
    header: "supplier_name||type||contact_details",
    start: 3,
    cols: { supplier: 1, type: 2, contact: 3 },
    build: (r) => [r.supplier, r.type, r.contact],
  },
];

function clean(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim().split(" ").filter(Boolean).join(" ");
}

function lower(v) {
  return clean(v).toLowerCase();
}

function depId(v, allowAll = false) {
  const key = lower(v);
  if (allowAll && ["tutti", "all", "all departments"].includes(key)) return "ALL";
  return DEPARTMENT_MAP[key] || clean(v);
}

function username(firstName, lastName) {
  const first = clean(firstName);
  let last = clean(lastName).toLowerCase();
  last = last.normalize("NFD").split("").filter((c) => c >= "a" && c <= "z").join("");
  if (!first || !last) return "";
  return first.charAt(0).toUpperCase() + last;
}

function checked(v) {
  const x = lower(v);
  return ["x", "yes", "si", "sì", "true", "1"].includes(x) ? "1" : "0";
}

function severity(v) {
  const text = clean(v);
  for (const c of text) if (c >= "0" && c <= "9") return c;
  return text;
}

function excelTime(v) {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number" && v >= 0 && v < 1) {
    const total = Math.round(v * 24 * 60);
    const h = String(Math.floor(total / 60)).padStart(2, "0");
    const m = String(total % 60).padStart(2, "0");
    return h + ":" + m;
  }
  return clean(v);
}

function skipRow(raw, mapped) {
  const marker = lower(raw[0]);
  if (["esempi", "examples", "example"].includes(marker)) return true;
  return Object.values(mapped).map(clean).every((v) => !v);
}

function parseSection(workbook, config) {
  const sheet = workbook.Sheets[config.sourceSheet];
  if (!sheet) return { ...config, rows: [], lines: [], warnings: ["Foglio mancante: " + config.sourceSheet] };

  const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const rows = [];
  const lines = [];
  const warnings = [];
  const seen = new Set();

  for (let i = config.start; i < rawRows.length; i++) {
    const raw = rawRows[i];
    const mapped = {};
    Object.entries(config.cols).forEach(([key, col]) => mapped[key] = clean(raw[col]));
    if (skipRow(raw, mapped)) continue;

    const built = config.build(mapped, rows.length, raw).map(clean);
    if (config.unique) {
      const uniqueKey = built.join(PIPE).toLowerCase();
      if (!uniqueKey || seen.has(uniqueKey)) continue;
      seen.add(uniqueKey);
    }

    const line = built.join(PIPE);
    if (!built[0]) warnings.push("Riga " + (i + 1) + ": valore principale mancante");
    if (config.id === "users" && built[7] && !["HOU", "MAIN", "REC", "FB", "MAN"].includes(built[7])) warnings.push("Riga " + (i + 1) + ": department non mappato: " + mapped.department);
    rows.push({ sourceRow: i + 1, line });
    lines.push(line);
  }

  if (config.id === "faults") {
    rows.sort((a, b) => a.line.localeCompare(b.line));
    lines.splice(0, lines.length, ...rows.map((r) => r.line));
  }

  return { ...config, rows, lines, warnings };
}

function downloadText(filename, content) {
  const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export default function HoxellDataCollectImporter() {
  const [fileName, setFileName] = useState("");
  const [sheetNames, setSheetNames] = useState([]);
  const [sections, setSections] = useState([]);
  const [activeId, setActiveId] = useState(SECTIONS[0].id);
  const [copiedId, setCopiedId] = useState("");

  const active = useMemo(() => sections.find((s) => s.id === activeId) || sections[0], [sections, activeId]);

  async function handleFile(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    setCopiedId("");
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    setSheetNames(workbook.SheetNames);
    const parsed = SECTIONS.map((section) => parseSection(workbook, section));
    setSections(parsed);
    setActiveId(parsed[0].id);
  }

  function output(section) {
    if (!section) return "";
    return [section.header, ...section.lines].join(NL);
  }

  function allOutput() {
    return sections.map((s) => "### " + s.label + NL + output(s)).join(NL + NL);
  }

  async function copy(section) {
    await navigator.clipboard.writeText(output(section));
    setCopiedId(section.id);
    setTimeout(() => setCopiedId(""), 1600);
  }

  return (
    <div className="min-h-screen bg-slate-50 p-6 text-slate-900">
      <div className="mx-auto max-w-7xl space-y-6">
        <div className="rounded-3xl bg-white p-8 shadow-sm">
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div>
              <p className="text-sm font-medium uppercase tracking-wide text-slate-500">Hoxell utility</p>
              <h1 className="mt-2 text-3xl font-bold tracking-tight">Data Collect → Stringhe Import</h1>
              <p className="mt-3 max-w-3xl text-slate-600">Carica il Data Collect e l'app divide automaticamente i dati per foglio: room type, camere, aree comuni, minibar, lavanderia, extra, lost & found, guasti, utenti, turni e building automation.</p>
            </div>
            <div className="rounded-2xl bg-slate-100 p-4"><FileSpreadsheet className="h-10 w-10 text-slate-700" /></div>
          </div>
        </div>

        <Card className="rounded-3xl border-0 shadow-sm">
          <CardContent className="p-6">
            <label className="flex cursor-pointer flex-col items-center justify-center rounded-3xl border-2 border-dashed border-slate-300 bg-white p-10 text-center transition hover:bg-slate-50">
              <Upload className="mb-4 h-10 w-10 text-slate-500" />
              <span className="text-lg font-semibold">Carica Data Collect Excel</span>
              <span className="mt-1 text-sm text-slate-500">Formato supportato: .xlsx / .xls</span>
              <input type="file" accept=".xlsx,.xls" onChange={handleFile} className="hidden" />
            </label>
            {fileName && <div className="mt-4 rounded-2xl bg-slate-100 p-4 text-sm text-slate-700">File caricato: <span className="font-semibold">{fileName}</span></div>}
          </CardContent>
        </Card>

        {!!sheetNames.length && <Card className="rounded-3xl border-0 shadow-sm"><CardContent className="p-6"><h2 className="mb-3 text-lg font-semibold">Fogli trovati nel file</h2><div className="flex flex-wrap gap-2">{sheetNames.map((name) => <span key={name} className="rounded-full bg-slate-100 px-3 py-1 text-sm text-slate-700">{name}</span>)}</div></CardContent></Card>}

        {!!sections.length && <div className="grid gap-6 lg:grid-cols-[320px_1fr]">
          <Card className="rounded-3xl border-0 shadow-sm"><CardContent className="p-4">
            <div className="mb-3 flex items-center justify-between px-2"><h2 className="text-lg font-semibold">Sezioni</h2><Button variant="outline" className="rounded-2xl" onClick={() => downloadText("hoxell_all_import_strings.txt", allOutput())}><Download className="mr-2 h-4 w-4" /> Tutto</Button></div>
            <div className="space-y-2">{sections.map((section) => <button key={section.id} onClick={() => setActiveId(section.id)} className={"w-full rounded-2xl p-3 text-left transition " + (activeId === section.id ? "bg-slate-900 text-white" : "bg-slate-100 hover:bg-slate-200")}><div className="flex items-center justify-between gap-2"><span className="font-medium">{section.label}</span>{section.warnings.length ? <AlertTriangle className="h-4 w-4 text-amber-500" /> : <CheckCircle2 className="h-4 w-4 text-emerald-500" />}</div><div className={"mt-1 text-xs " + (activeId === section.id ? "text-slate-300" : "text-slate-500")}>{section.lines.length} righe generate</div></button>)}</div>
          </CardContent></Card>

          {active && <div className="space-y-6">
            {!!active.warnings.length && <Card className="rounded-3xl border-0 bg-amber-50 shadow-sm"><CardContent className="p-6"><div className="mb-3 flex items-center gap-2 font-semibold text-amber-900"><AlertTriangle className="h-5 w-5" /> Controlli da fare</div><div className="space-y-1 text-sm text-amber-900">{active.warnings.slice(0, 12).map((warning, index) => <p key={index}>{warning}</p>)}</div></CardContent></Card>}
            <Card className="rounded-3xl border-0 shadow-sm"><CardContent className="p-6">
              <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between"><div><h2 className="text-xl font-semibold">{active.label}</h2><p className="mt-1 text-sm text-slate-500">Foglio sorgente: {active.sourceSheet}</p></div><div className="flex gap-2"><Button onClick={() => copy(active)} variant="outline" className="rounded-2xl"><Copy className="mr-2 h-4 w-4" /> {copiedId === active.id ? "Copiato" : "Copia"}</Button><Button onClick={() => downloadText(active.id + "_import.txt", output(active))} className="rounded-2xl"><Download className="mr-2 h-4 w-4" /> TXT</Button></div></div>
              <textarea className="h-96 w-full rounded-2xl border border-slate-200 bg-slate-950 p-4 font-mono text-sm text-slate-100 outline-none" value={output(active)} readOnly />
            </CardContent></Card>
            <Card className="rounded-3xl border-0 shadow-sm"><CardContent className="p-6"><h3 className="mb-3 text-lg font-semibold">Anteprima dati letti</h3>{!active.rows.length ? <p className="text-sm text-slate-500">Nessuna riga utile trovata.</p> : <div className="max-h-80 overflow-auto rounded-2xl border border-slate-200"><table className="w-full text-left text-sm"><thead className="sticky top-0 bg-slate-100"><tr><th className="p-3">Riga</th><th className="p-3">Stringa generata</th></tr></thead><tbody>{active.rows.map((row, index) => <tr key={active.id + index} className="border-t border-slate-100"><td className="w-20 p-3 text-slate-500">{row.sourceRow}</td><td className="p-3 font-mono text-xs">{row.line}</td></tr>)}</tbody></table></div>}</CardContent></Card>
          </div>}
        </div>}
      </div>
    </div>
  );
}
