import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  Upload,
  FileSpreadsheet,
  Copy,
  Download,
  AlertTriangle,
  CheckCircle2,
} from "lucide-react";

const PIPE = "||";
const NL = "\n";

const DEPARTMENT_MAP: Record<string, string> = {
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
};

type RowMapped = Record<string, any>;

type SectionConfig = {
  id: string;
  label: string;
  sourceSheet: string;
  header: string;
  start: number;
  cols: Record<string, number>;
  unique?: boolean;
  build: (r: RowMapped, i: number, raw: any[]) => any[];
};

const SECTIONS: SectionConfig[] = [
  {
    id: "rooms",
    label: "Rooms / Camere",
    sourceSheet: "Camere",
    header:
      "name||pms_room_name||display_order||cleaning_time_empty||cleaning_time_arrival_or_departure||cleaning_time_in_house",
    start: 3,
    cols: {
      name: 1,
      empty: 5,
      departure: 6,
      inHouse: 7,
    },
    build: (r, i) => [
      r.name,
      r.name,
      r.name || i + 1,
      r.empty,
      r.departure,
      r.inHouse,
    ],
  },
  {
    id: "roomSections",
    label: "Room Sections / Sezioni Camere",
    sourceSheet: "Camere",
    header: "name",
    start: 3,
    cols: {
      section: 4,
    },
    unique: true,
    build: (r) => [r.section],
  },
  {
    id: "commonAreas",
    label: "Common Areas / Spazi Comuni",
    sourceSheet: "Aree Comuni",
    header:
      "name||pms_room_name||display_order||cleaning_time_empty||cleaning_time_arrival_or_departure||cleaning_time_in_house",
    start: 3,
    cols: {
      name: 1,
    },
    build: (r, i) => [
      r.name,
      r.name,
      100100 + i * 10,
      "",
      "",
      "",
    ],
  },
  {
    id: "extraItems",
    label: "Extra Items / Extra",
    sourceSheet: "Extra",
    header: "name||category||cost||price||room_type_code||display_order",
    start: 4,
    cols: {
      name: 1,
      category: 2,
      cost: 3,
      price: 4,
      roomType: 5,
    },
    build: (r, i) => [
      r.name,
      r.category,
      r.cost,
      r.price,
      r.roomType,
      (i + 1) * 10,
    ],
  },
  {
    id: "lostFound",
    label: "Lost & Found",
    sourceSheet: "Lost & Found",
    header: "storage_area",
    start: 3,
    cols: {
      name: 1,
    },
    build: (r) => [r.name],
  },
  {
    id: "faultCategories",
    label: "Maintenance Categories / Categorie Guasti",
    sourceSheet: "Guasti",
    header: "name",
    start: 4,
    cols: {
      category: 3,
    },
    unique: true,
    build: (r) => [r.category],
  },
  {
  id: "faults",
  label: "Maintenance Faults / Guasti per Categoria",
  sourceSheet: "Guasti",
  header: "category||severity",
  start: 4,
  cols: {
    severity: 2,
    category: 3,
    rooms: 4,
    commons: 5,
  },
  build: (r) => {
    const roomInfo = checked(r.rooms) === "1" ? "ROOMS" : "";
    const commonInfo =
      checked(r.commons) === "1" ? "COMMON AREAS" : "";

    console.log(
      `Categoria: ${r.category} | Rooms: ${roomInfo} | Common: ${commonInfo}`
    );

    return [
      r.category,
      severity(r.severity),
    ];
  },
}
  {
    id: "users",
    label: "Users / Utenti",
    sourceSheet: "Utenti",
    header:
      "first_name||last_name||username||pms_id||job_title||email||language||department_id",
    start: 3,
    cols: {
      firstName: 1,
      lastName: 2,
      department: 3,
      jobTitle: 4,
    },
    build: (r) => [
      r.firstName,
      r.lastName,
      username(r.firstName, r.lastName),
      "",
      r.jobTitle,
      "",
      "",
      depId(r.department),
    ],
  },
  {
    id: "buildingAutomation",
    label: "Building Automation",
    sourceSheet: "Building Automation",
    header: "supplier_name||type||contact_details",
    start: 3,
    cols: {
      supplier: 1,
      type: 2,
      contact: 3,
    },
    build: (r) => [r.supplier, r.type, r.contact],
  },
];

function clean(value: any): string {
  if (value === null || value === undefined) return "";
  return String(value).trim().replace(/\s+/g, " ");
}

function lower(value: any): string {
  return clean(value).toLowerCase();
}

function depId(value: any): string {
  const key = lower(value);
  return DEPARTMENT_MAP[key] || clean(value);
}

function username(firstName: any, lastName: any): string {
  const first = clean(firstName);

  let last = clean(lastName)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z]/g, "");

  if (!first || !last) return "";

  return first.charAt(0).toUpperCase() + last;
}

function checked(value: any): string {
  const key = lower(value);
  return ["x", "yes", "si", "sì", "true", "1"].includes(key)
    ? "1"
    : "0";
}

function severity(value: any): string {
  const text = clean(value);
  const match = text.match(/\d+/);
  return match ? match[0] : text;
}

function shouldSkipRow(raw: any[], mapped: RowMapped): boolean {
  const marker = lower(raw[0]);

  if (["esempi", "example", "examples"].includes(marker)) {
    return true;
  }

  return Object.values(mapped)
    .map(clean)
    .every((v) => !v);
}

function parseSection(workbook: XLSX.WorkBook, config: SectionConfig) {
  const sheet = workbook.Sheets[config.sourceSheet];

  if (!sheet) {
    return {
      ...config,
      rows: [],
      lines: [],
      warnings: [`Foglio mancante: ${config.sourceSheet}`],
    };
  }

  const rawRows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
  }) as any[][];

  const rows: any[] = [];
  const lines: string[] = [];
  const warnings: string[] = [];
  const seen = new Set<string>();

  for (let i = config.start; i < rawRows.length; i++) {
    const raw = rawRows[i];
    const mapped: RowMapped = {};

    Object.entries(config.cols).forEach(([key, colIndex]) => {
      mapped[key] = clean(raw[colIndex]);
    });

    if (shouldSkipRow(raw, mapped)) continue;

    const built = config.build(mapped, rows.length, raw).map(clean);

    if (config.unique) {
      const uniqueKey = built.join(PIPE).toLowerCase();

      if (!uniqueKey || seen.has(uniqueKey)) continue;

      seen.add(uniqueKey);
    }

    const line = built.join(PIPE);

    rows.push({
      sourceRow: i + 1,
      line,
    });

    lines.push(line);
  }

  if (config.id === "faults") {
    rows.sort((a, b) => a.line.localeCompare(b.line));
    lines.splice(0, lines.length, ...rows.map((r) => r.line));
  }

  return {
    ...config,
    rows,
    lines,
    warnings,
  };
}

function downloadText(filename: string, content: string) {
  const blob = new Blob([content], {
    type: "text/plain;charset=utf-8",
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");

  a.href = url;
  a.download = filename;

  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);

  URL.revokeObjectURL(url);
}

export default function App() {
  const [fileName, setFileName] = useState("");
  const [sections, setSections] = useState<any[]>([]);
  const [activeId, setActiveId] = useState(SECTIONS[0].id);

  const activeSection = useMemo(() => {
    return sections.find((s) => s.id === activeId) || sections[0];
  }, [sections, activeId]);

  async function handleFile(event: any) {
    const file = event.target.files?.[0];

    if (!file) return;

    setFileName(file.name);

    const buffer = await file.arrayBuffer();

    const workbook = XLSX.read(buffer, {
      type: "array",
    });

    const parsed = SECTIONS.map((section) =>
      parseSection(workbook, section)
    );

    setSections(parsed);
    setActiveId(parsed[0]?.id || SECTIONS[0].id);
  }

  function sectionOutput(section: any): string {
    if (!section) return "";

    return [section.header, ...section.lines].join(NL);
  }

  return (
    <div
      style={{
        minHeight: "100vh",
        background: "#f8fafc",
        padding: 24,
        fontFamily: "Arial",
      }}
    >
      <div style={{ maxWidth: 1400, margin: "0 auto" }}>
        <div
          style={{
            background: "white",
            borderRadius: 24,
            padding: 32,
            marginBottom: 24,
          }}
        >
          <h1>Hoxell Data Collect</h1>

          <p>Carica il Data Collect Excel.</p>

          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFile}
          />

          {fileName && (
            <p style={{ marginTop: 12 }}>
              File caricato: <b>{fileName}</b>
            </p>
          )}
        </div>

        {!!sections.length && (
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "300px 1fr",
              gap: 24,
            }}
          >
            <div
              style={{
                background: "white",
                borderRadius: 24,
                padding: 20,
              }}
            >
              <h2>Sezioni</h2>

              {sections.map((section) => (
                <button
                  key={section.id}
                  onClick={() => setActiveId(section.id)}
                  style={{
                    width: "100%",
                    padding: 12,
                    marginTop: 10,
                    borderRadius: 16,
                    border: "none",
                    cursor: "pointer",
                    background:
                      activeId === section.id
                        ? "#0f172a"
                        : "#e2e8f0",
                    color:
                      activeId === section.id
                        ? "white"
                        : "black",
                  }}
                >
                  {section.label}
                </button>
              ))}
            </div>

            {activeSection && (
              <div
                style={{
                  background: "white",
                  borderRadius: 24,
                  padding: 24,
                }}
              >
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    marginBottom: 20,
                  }}
                >
                  <h2>{activeSection.label}</h2>

                  <button
                    onClick={() =>
                      downloadText(
                        `${activeSection.id}.txt`,
                        sectionOutput(activeSection)
                      )
                    }
                    style={{
                      padding: "10px 14px",
                      borderRadius: 12,
                      border: "none",
                      background: "#0f172a",
                      color: "white",
                      cursor: "pointer",
                    }}
                  >
                    Download TXT
                  </button>
                </div>

                <textarea
                  value={sectionOutput(activeSection)}
                  readOnly
                  style={{
                    width: "100%",
                    height: 500,
                    borderRadius: 16,
                    padding: 16,
                    background: "#020617",
                    color: "white",
                    fontFamily: "monospace",
                    fontSize: 12,
                  }}
                />
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
