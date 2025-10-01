import fs from "fs";
import path from "path";
import XLSX from "xlsx";
import moment from "moment-timezone";

// MASHREQ_ROUTING => 203320101
// ADIB_ROUTING => 405010101

// EMPLOYER_ID => 0000002571863  // total 13 chars 
// EMPLOYEE_ID => up to 14 chars 

// ================== HARD-CODED EMPLOYER DETAILS ==================
const EMPLOYER_ID = "0000002571863";   // from you
const EMPLOYER_ROUTING = "203320101";  
const CURRENCY = "AED";
const REFERENCE = "0000002571863";
// =================================================================

// CLI: --input / -i (required), --output / -o (optional)
function getArg(name, def = undefined) {
  const i = process.argv.findIndex(a => a === name || a.startsWith(name + "="));
  if (i === -1) return def;
  const a = process.argv[i];
  if (a.includes("=")) return a.split("=").slice(1).join("=");
  return process.argv[i + 1] ?? def;
}
function die(msg) { console.error("❌ " + msg); process.exit(1); }
function to2(n) { return (Number(n) || 0).toFixed(2); }
function normHeader(h) { return String(h ?? "").toLowerCase().replace(/\s+/g, "_").trim(); }

function parseDateISO(s) {
  // Expect YYYY-MM-DD
  if (!s) return null;
  const d = new Date(`${s}T00:00:00Z`);
  if (Number.isNaN(d.getTime())) return null;
  const Y = d.getUTCFullYear();
  const M = (d.getUTCMonth() + 1).toString().padStart(2, "0");
  const D = d.getUTCDate().toString().padStart(2, "0");
  return `${Y}-${M}-${D}`;
}
function daysInclusive(startISO, endISO) {
  const s = new Date(startISO + "T00:00:00Z");
  const e = new Date(endISO + "T00:00:00Z");
  return Math.floor((e - s) / (1000*60*60*24)) + 1;
}
function pick(row, map) {
  const out = {};
  for (const [want, aliases] of Object.entries(map)) {
    let val;
    for (const a of aliases) {
      if (row[a] !== undefined && row[a] !== null && String(row[a]).trim() !== "") {
        val = row[a];
        break;
      }
    }
    out[want] = typeof val === "string" ? val.trim() : val;
  }
  return out;
}
function mostCommon(arr) {
  const m = new Map();
  for (const v of arr) m.set(v, (m.get(v) || 0) + 1);
  let best = null, cnt = -1;
  for (const [k, v] of m) if (v > cnt) { best = k; cnt = v; }
  return best;
}
function nowDubai() {
  const m = moment.tz("Asia/Dubai");
  return {
    dateISO: m.format("YYYY-MM-DD"), // e.g., 2025-10-01
    hhmm:    m.format("HHmm"),       // e.g., 1742
    dateTime: m.format("YYMMDDHHmmss") // e.g., 251001174201
  };
}



// -------------------- MAIN --------------------
const inputPath = getArg("--input") || getArg("-i");
const outputPathArg = getArg("--output") || getArg("-o");
if (!inputPath) die(`Usage: node make_sif_hardcoded.mjs --input ./employees.xlsx [--output ./salary.sif]`);
if (!fs.existsSync(inputPath)) die(`Input not found: ${inputPath}`);

const wb = XLSX.readFile(inputPath);
const sheetName = wb.SheetNames[0];
if (!sheetName) die("No sheets found.");
const ws = wb.Sheets[sheetName];

// defval keeps blanks as "" (not undefined)
let rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
if (rows.length === 0) die("No rows in the first sheet.");

// Normalize headers
const headerMap = {};
Object.keys(rows[0]).forEach(k => headerMap[normHeader(k)] = k);
rows = rows.map(r => {
  const o = {};
  for (const [norm, orig] of Object.entries(headerMap)) o[norm] = r[orig];
  return o;
});

// Expected columns (case/space-insensitive)
const mapping = {
  employee_id: ["employee_id","emp_id","id","employeeid","employee id"],
  employee_routing: ["employee_routing","routing","bank_routing","routing_code"],
  employee_iban: ["employee_iban","iban","employee_iban_number","employee iban"],
  period_start: ["period_start","start_date","salary_start","from_date"],
  period_end: ["period_end","end_date","salary_end","to_date"],
  period_days: ["period_days","days","num_days"],
  fixed_amount: ["fixed_amount","fixed","basic","basic_amount"],
  variable_amount: ["variable_amount","variable","allowance","bonus","overtime"],
  unpaid_leave_days: ["unpaid_leave_days","unpaid_days","lwop_days"]
};

const edrLines = [];
let totalAmount = 0;
let count = 0;
const endMonths = []; // collect MMYYYY from period_end to infer salaryMonth

function excelSerialToISO(serial) {
  // Excel serial: days since 1899-12-30 (Excel's epoch, incl. 1900 leap-year bug)
  const n = Number(serial);
  if (!Number.isFinite(n)) return null;
  const ms = Math.round((n * 24 * 60 * 60) * 1000); // include time fraction
  const excelEpoch = Date.UTC(1899, 11, 30); // 1899-12-30
  const d = new Date(excelEpoch + ms);
  if (Number.isNaN(d.getTime())) return null;
  const Y = d.getUTCFullYear().toString().padStart(4, "0");
  const M = (d.getUTCMonth() + 1).toString().padStart(2, "0");
  const D = d.getUTCDate().toString().padStart(2, "0");
  return `${Y}-${M}-${D}`;
}

function normalizeYMD(y, m, d) {
  // pads & validates
  const Y = String(y).padStart(4, "0");
  const M = String(m).padStart(2, "0");
  const D = String(d).padStart(2, "0");
  const test = new Date(`${Y}-${M}-${D}T00:00:00Z`);
  if (Number.isNaN(test.getTime())) return null;
  return `${Y}-${M}-${D}`;
}

function parseDateAny(v) {
  if (v === undefined || v === null) return null;
  const s = String(v).trim();

  // 1) Excel serial number (integer or float)
  if (/^\d+(\.\d+)?$/.test(s)) {
    const iso = excelSerialToISO(s);
    if (iso) return iso;
  }

  // 2) ISO-like strings: YYYY-MM-DD or with time
  //    accept 2025-08-31, 2025/08/31, 2025-08-31T12:34:56, etc.
  let m;
  if ((m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/))) {
    const [, Y, M, D] = m.map(Number);
    const iso = normalizeYMD(Y, M, D);
    if (iso) return iso;
  }

  // 3) D/M/Y or M/D/Y (ambiguous). Heuristic: if first > 12 => D/M/Y else M/D/Y
  if ((m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})/))) {
    let a = Number(m[1]), b = Number(m[2]), c = Number(m[3]);
    if (c < 100) c += 2000; // 2-digit year -> 20xx
    const isDMY = a > 12; // if first part >12, treat as day
    const Y = c, M = isDMY ? b : a, D = isDMY ? a : b;
    const iso = normalizeYMD(Y, M, D);
    if (iso) return iso;
  }

  // 4) Fallback: let Date parse and re-normalize if valid
  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) {
    const Y = d.getUTCFullYear();
    const M = d.getUTCMonth() + 1;
    const D = d.getUTCDate();
    const iso = normalizeYMD(Y, M, D);
    if (iso) return iso;
  }

  return null; // not parseable
}


for (const [idx, raw] of rows.entries()) {
  const r = pick(raw, mapping);
  const rowNum = idx + 2;

  // Required fields
  const must = ["employee_id","employee_routing","employee_iban","period_start","period_end","fixed_amount","variable_amount","unpaid_leave_days"];
  for (const m of must) {
    if (r[m] === undefined || r[m] === null || String(r[m]).trim() === "") {
      die(`Row ${rowNum}: Missing required "${m}"`);
    }
  }

  const startISO = parseDateAny(String(r.period_start));
  const endISO = parseDateAny(String(r.period_end));
  if (!startISO) die(`Row ${rowNum}: Invalid period_start "${r.period_start}" (YYYY-MM-DD).`);
  if (!endISO) die(`Row ${rowNum}: Invalid period_end "${r.period_end}" (YYYY-MM-DD).`);

  let days = Number(r.period_days);
  if (!Number.isFinite(days) || days <= 0) {
    days = daysInclusive(startISO, endISO);
    if (!Number.isFinite(days) || days <= 0) {
      die(`Row ${rowNum}: period_days missing/invalid and cannot be derived.`);
    }
  }

  const fixed = Number(String(r.fixed_amount).replace(/,/g,""));
  const variable = Number(String(r.variable_amount).replace(/,/g,""));
  const unpaid = Number(r.unpaid_leave_days);

  if (!Number.isFinite(fixed)) die(`Row ${rowNum}: fixed_amount not a number.`);
  if (!Number.isFinite(variable)) die(`Row ${rowNum}: variable_amount not a number.`);
  if (!Number.isFinite(unpaid)) die(`Row ${rowNum}: unpaid_leave_days not a number.`);

  const employeeIban = String(r.employee_iban).replace(/\s+/g,"").toUpperCase();
  if (!/^AE[0-9A-Z]{21}$/i.test(employeeIban)) {
    die(`Row ${rowNum}: employee_iban must be UAE IBAN (AE + 21 chars). Got: ${r.employee_iban}`);
  }

  const employeeRouting = String(r.employee_routing).trim();
  if (!/^\d{9}$/.test(employeeRouting)) {
    console.warn(`⚠️  Row ${rowNum}: employee_routing is usually 9 digits. Got: "${employeeRouting}"`);
  }

  // Build EDR line (Mashreq comma-separated order)
  const edr = [
    "EDR",
    String(r.employee_id).trim(),
    employeeRouting,
    employeeIban,
    startISO,
    endISO,
    String(days),
    to2(fixed),
    to2(variable),
    String(unpaid)
  ].join(",");

  edrLines.push(edr);
  totalAmount += fixed + variable;
  count += 1;

  // derive MMYYYY from period_end
  const [Y, M] = endISO.split("-");
  endMonths.push(`${M}${Y}`);
}
if (count === 0) die("No valid employee rows found.");

// Auto: salaryMonth from most common period_end month; else today's month
const guessedSalaryMonth = mostCommon(endMonths);
const now = nowDubai();
const salaryMonth = guessedSalaryMonth || (() => {
  const d = now.dateISO; // YYYY-MM-DD
  const [Y, M] = d.split("-");
  return `${M}${Y}`;
})();

// Auto: creationDate / creationTime (Dubai)
const creationDate = now.dateISO; // YYYY-MM-DD
const creationTime = now.hhmm;    // HHMM

// Build SCR header/summary
const scr = [
  "SCR",
  EMPLOYER_ID,
  EMPLOYER_ROUTING, // keep as-is (you'll replace later)
  creationDate,
  creationTime,
  salaryMonth,
  String(count),
  to2(totalAmount),
  CURRENCY,
  REFERENCE
].join(",");

// Output
const defaultName = `${EMPLOYER_ID}${nowDubai().dateTime}.sif`;
const outPath =  path.join(process.cwd(), defaultName)
const all = [scr, ...edrLines].join("\n");
fs.writeFileSync(outPath, all, "utf8");

console.log("✅ SIF generated");
console.log("   File:", outPath);
console.log("   Employees:", count);
console.log("   Total amount:", to2(totalAmount));
console.log("   Salary month:", salaryMonth);
console.log("   Created (Dubai):", creationDate, creationTime);

// Friendly reminder if routing is a placeholder
if (!/^\d{9}$/.test(EMPLOYER_ROUTING)) {
  console.warn("⚠️  EMPLOYER_ROUTING is not 9 digits (placeholder). Update it before bank upload.");
}
