import * as XLSX from "xlsx";

function escapeCsvCell(value: unknown): string {
  const s = String(value ?? "");
  if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function normStr(v: unknown): string {
  if (v === null || v === undefined) return "";
  const s = String(v).trim();
  if (s.toLowerCase() === "nan") return "";
  return s;
}

export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cur = "";
  let inQuotes = false;
  const t = text.replace(/^\uFEFF/, "");
  for (let i = 0; i < t.length; i++) {
    const c = t[i];
    if (inQuotes) {
      if (c === '"') {
        if (t[i + 1] === '"') {
          cur += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        cur += c;
      }
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ",") {
      row.push(cur);
      cur = "";
    } else if (c === "\n" || c === "\r") {
      if (c === "\r" && t[i + 1] === "\n") i++;
      row.push(cur);
      cur = "";
      if (row.some((cell) => String(cell).trim() !== "")) rows.push(row);
      row = [];
    } else {
      cur += c;
    }
  }
  row.push(cur);
  if (row.some((cell) => String(cell).trim() !== "")) rows.push(row);
  return rows;
}

function parseCsvStreaming(
  text: string,
  onRow: (row: string[]) => void
): void {
  let row: string[] = [];
  let cur = "";
  let inQuotes = false;
  const t = text.replace(/^\uFEFF/, "");

  const pushRowIfNotEmpty = () => {
    if (row.some((cell) => String(cell).trim() !== "")) onRow(row);
    row = [];
  };

  for (let i = 0; i < t.length; i++) {
    const c = t[i];
    if (inQuotes) {
      if (c === '"') {
        if (t[i + 1] === '"') {
          cur += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        cur += c;
      }
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ",") {
      row.push(cur);
      cur = "";
    } else if (c === "\n" || c === "\r") {
      if (c === "\r" && t[i + 1] === "\n") i++;
      row.push(cur);
      cur = "";
      pushRowIfNotEmpty();
    } else {
      cur += c;
    }
  }

  row.push(cur);
  pushRowIfNotEmpty();
}

function getSpeciesSheet(wb: XLSX.WorkBook): XLSX.WorkSheet {
  const name = wb.SheetNames.includes("Species")
    ? "Species"
    : wb.SheetNames[0];
  if (!name) throw new Error("Workbook has no sheets.");
  const sh = wb.Sheets[name];
  if (!sh) throw new Error("Species sheet not found.");
  return sh;
}

type MasterEntry = {
  human_id: string;
  ai_id: string;
  incongruent_reason: string;
};

const REASON_COLUMNS = [
  "Blurry",
  "Low-light",
  "Body part",
  "Blends in",
  "Unidentifiable to taxonomic level by human ground-truth",
  "Other",
  "Similar species that does not occur in the area",
];

// CHANGED: Now async, receives getter functions to control memory pacing
export async function mergeMasterSheet(input: {
  getFile1: () => Promise<Buffer>;
  getFile2: () => Promise<Buffer>;
  getFile3: () => Promise<Buffer>;
  getJsonFile: () => Promise<Buffer>;
  getMetadataFiles: (() => Promise<Buffer>)[];
}): Promise<string> {
  const masterData = new Map<string, MasterEntry>();

  // ==========================================
  // PROCESS FILE 1 (Then instantly free memory)
  // ==========================================
  let buf1: Buffer | null = await input.getFile1();
  let wb1: XLSX.WorkBook | null = XLSX.read(buf1, { type: "buffer", cellStyles: false, dense: true });
  buf1 = null; // Free Buffer RAM
  
  let sh1: XLSX.WorkSheet | null = getSpeciesSheet(wb1);
  let rows1: Record<string, unknown>[] | null = XLSX.utils.sheet_to_json(sh1, { defval: "", raw: false, blankrows: false });
  wb1 = null; // Free Workbook RAM
  sh1 = null; // Free Sheet RAM

  for (const row of rows1) {
    const keys = Object.keys(row);
    const humanKey = keys.find((k) => k.trim().toLowerCase() === "human");
    const aiKey = keys.find((k) => k.trim().toLowerCase() === "ai");
    if (!humanKey || !aiKey) continue;

    const human_id = normStr(row[humanKey]);
    const ai_id = normStr(row[aiKey]);
    const fnameCols = keys.filter((k) => k.toLowerCase().includes("filename"));

    for (const col of fnameCols) {
      const filenamesStr = normStr(row[col]);
      if (!filenamesStr) continue;
      for (const fname of filenamesStr.split(",").map((f) => f.trim())) {
        if (!fname) continue;
        const lower = fname.toLowerCase();
        if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) {
          masterData.set(fname, { human_id, ai_id, incongruent_reason: "" });
        }
      }
    }
  }
  rows1 = null; // Free JSON Rows RAM

  if (masterData.size === 0) {
    throw new Error("File #1: no .jpg/.jpeg filenames found in Species sheet.");
  }

  // ==========================================
  // PROCESS FILE 2 (Then instantly free memory)
  // ==========================================
  let buf2: Buffer | null = await input.getFile2();
  let wb2: XLSX.WorkBook | null = XLSX.read(buf2, { type: "buffer", cellStyles: false, dense: true });
  buf2 = null; 
  
  let sh2: XLSX.WorkSheet | null = getSpeciesSheet(wb2);
  let rows2: Record<string, unknown>[] | null = XLSX.utils.sheet_to_json(sh2, { defval: "", raw: false });
  wb2 = null; 
  sh2 = null; 

  for (const row of rows2) {
    const keys = Object.keys(row);
    for (const reasonCol of REASON_COLUMNS) {
      const matchKey = keys.find((k) => k === reasonCol) || keys.find((k) => k.trim().toLowerCase() === reasonCol.toLowerCase());
      if (!matchKey) continue;
      const cellValue = normStr(row[matchKey]);
      if (!cellValue) continue;
      for (const fname of cellValue.split(",").map((f) => f.trim())) {
        const entry = masterData.get(fname);
        if (entry) entry.incongruent_reason = reasonCol;
      }
    }
  }
  rows2 = null; 

  // ==========================================
  // PROCESS FILE 3 (CSV)
  // ==========================================
  let buf3: Buffer | null = await input.getFile3();
  let csvText: string | null = buf3.toString("utf-8");
  buf3 = null;

  const finalIdMap = new Map<string, string>();
  let headerSeen = false;
  let fnIx = -1;
  let spIx = -1;
  let dataRows = 0;

  parseCsvStreaming(csvText, (r) => {
    if (!headerSeen) {
      headerSeen = true;
      const header3 = r.map((c) => c.trim().toLowerCase());
      fnIx = header3.findIndex((c) => c === "filename");
      spIx = header3.findIndex((c) => c === "species");
      if (fnIx < 0 || spIx < 0) throw new Error('File #3 must include "filename" and "species" columns.');
      return;
    }
    const fn = normStr(r[fnIx]);
    const sp = normStr(r[spIx]);
    if (fn) finalIdMap.set(fn, sp);
    dataRows++;
  });
  csvText = null;

  if (!headerSeen || dataRows < 1) throw new Error("File #3 must be a non-empty CSV with a header row.");

  // ==========================================
  // PROCESS JSON
  // ==========================================
  let jsonBuf: Buffer | null = await input.getJsonFile();
  let jsonText: string | null = jsonBuf.toString("utf-8");
  jsonBuf = null;

  let predictionsJson: { predictions?: unknown[] } | null;
  try {
    predictionsJson = JSON.parse(jsonText);
  } catch {
    throw new Error("Predictions file is not valid JSON.");
  }
  jsonText = null;

  const confidenceMap = new Map<string, { top: string; second: string }>();
  for (const pred of predictionsJson?.predictions || []) {
    if (!pred || typeof pred !== "object") continue;
    const p = pred as { filepath?: string; classifications?: { scores?: unknown[] } };
    const filepath = p.filepath ?? "";
    const base = filepath.replace(/\\/g, "/").split("/").pop()?.trim() ?? "";
    if (!base) continue;
    const scores = p.classifications?.scores ?? [];
    confidenceMap.set(base, {
      top: scores[0] !== undefined && scores[0] !== null ? String(scores[0]) : "",
      second: scores[1] !== undefined && scores[1] !== null ? String(scores[1]) : "",
    });
  }
  predictionsJson = null;

  // ==========================================
  // PROCESS METADATA CSVs
  // ==========================================
  const metadataLookup = new Map<string, { date: string; time: string }>();

  for (const getMetaFn of input.getMetadataFiles) {
    let metaBuf: Buffer | null = await getMetaFn();
    let metaText: string | null = metaBuf.toString("utf-8");
    metaBuf = null;

    try {
      let headerParsed = false;
      let fnameColIdx = -1;
      let tsColIdx = -1;

      parseCsvStreaming(metaText, (r) => {
        if (!headerParsed) {
          headerParsed = true;
          const h = r.map((c) => c.trim().toLowerCase());
          fnameColIdx = h.findIndex((c) => c.includes("filename") || c.includes("file"));
          tsColIdx = h.findIndex((c) => c.includes("timestamp") || c.includes("date") || c.includes("time"));
          if (fnameColIdx < 0 || tsColIdx < 0) throw new Error("METADATA_HEADER_MISMATCH");
          return;
        }

        const fname = normStr(r[fnameColIdx]);
        const ts_str = normStr(r[tsColIdx]);
        if (!ts_str) return;
        const baseName = fname.replace(/\\/g, "/").split("/").pop()?.replace(/\.[^.]+$/i, "").toLowerCase() ?? "";
        if (!baseName) return;
        const parts = ts_str.split(/\s+/).filter(Boolean);
        const date_val = parts[0] ?? "";
        const time_val = parts.slice(1).join(" ");
        metadataLookup.set(baseName, { date: date_val, time: time_val });
      });
    } catch {
      // Silently skip metadata file when it doesn't match expected columns.
    }
    metaText = null;
  }

  // ==========================================
  // BUILD FINAL CSV OUTPUT
  // ==========================================
  const cols = [
    "batchname", "filename", "date", "time", "ai_id", "human_id",
    "incongruent", "incongruent_reason", "confidence_score",
    "second_confidence", "final_id",
  ] as const;

  let csvOutput = cols.join(",") + "\n";

  for (const [fname, data] of masterData) {
    const parts = fname.split("_");
    const batch_name = parts.length >= 3 ? parts[1] : "Unknown";
    const fname_base = fname.replace(/\.[^.]+$/i, "").toLowerCase();
    const meta = metadataLookup.get(fname_base) ?? { date: "", time: "" };

    const h_id = data.human_id;
    const a_id = data.ai_id;
    const incongruent = h_id && a_id && h_id.toLowerCase() === a_id.toLowerCase() ? "no" : "yes";
    const conf = confidenceMap.get(fname) ?? { top: "", second: "" };
    
    let final_id_val = finalIdMap.get(fname) ?? "";
    if (final_id_val.toLowerCase() === "nan") final_id_val = "";

    const row: Record<(typeof cols)[number], string> = {
      batchname: batch_name,
      filename: fname,
      date: meta.date,
      time: meta.time,
      ai_id: a_id,
      human_id: h_id,
      incongruent,
      incongruent_reason: data.incongruent_reason,
      confidence_score: conf.top,
      second_confidence: conf.second,
      final_id: final_id_val,
    };

    csvOutput += cols.map((c) => escapeCsvCell(row[c])).join(",") + "\n";
  }

  return csvOutput;
}