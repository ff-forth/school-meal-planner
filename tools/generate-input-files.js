#!/usr/bin/env node
/*
 * generate-input-files.js
 *
 * Turns the raw production files in data/ into the two clean input files the
 * planner imports:
 *
 *   1. data/skolmat-plan-vX-vY-All.xlsx
 *        From "Inskrivna & Matsedel.xlsx" (sheet "Matsedel skolor").
 *        One row per served day, full format:
 *          Datum | Veckodag | Enhet | Måltid | Vegetariskt
 *          | Protein | Kolhydrat | Grönäsar | Sås | Tillagning
 *          | Svinn/ätande (mätt) | Svinn/ätande (estimerat)
 *        Enhet = "All". Components are derived from the meal name. The two
 *        svinn columns are left empty (they come from file 2).
 *
 *   2. data/svinn-report-period.xlsx
 *        From "Portioner …xlsx" + "Svinn …xlsx" + "rapport-2.xlsx".
 *        One row per Enhet + Datum (date is the key), filtered to Lunch:
 *          Enhet | Datum | Antal ätande | Totalt svinn (kg)
 *          | Svinn/ätande (mätt) | Svinn/ätande (estimerat)
 *        Antal ätande from Portioner, Totalt svinn from Svinn; when both are
 *        present, Svinn/ätande (mätt) = svinn / ätande * 1000. Rapport rows
 *        (per-meal) are aggregated to the same Enhet+Datum key.
 *
 * Run:  node tools/generate-input-files.js
 * (uses the xlsx package already in node_modules)
 */
"use strict";
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const DATA = path.join(__dirname, "..", "data");
const OUT = path.join(DATA, "processed");
if (!fs.existsSync(OUT)) fs.mkdirSync(OUT, { recursive: true });

// ---- load the app's decompose + component helpers (they attach to `window`) ----
global.window = {};
for (const f of ["js/data.js", "js/decompose.js", "js/import-real.js"]) {
  eval(fs.readFileSync(path.join(__dirname, "..", f), "utf8"));
}
const D = global.window.MealData;
const Decompose = global.window.Decompose;

// ---- date helpers ----
function iso(d) {
  if (!(d instanceof Date) || isNaN(d)) return null;
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}
function isoWeek(d) {
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  t.setUTCDate(t.getUTCDate() + 4 - (t.getUTCDay() || 7));
  return Math.ceil(((t - new Date(Date.UTC(t.getUTCFullYear(), 0, 1))) / 864e5 + 1) / 7);
}
function norm(s) {
  return String(s == null ? "" : s).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").trim();
}
function cellDate(v) {
  return v instanceof Date && !isNaN(v) ? v : null;
}
function num(v) {
  if (typeof v === "number") return isFinite(v) ? v : null;
  if (v == null) return null;
  const n = parseFloat(String(v).replace(",", "."));
  return isFinite(n) ? n : null;
}
const LUNCH = (s) => norm(s) === "lunch";

// =========================================================================
// FILE 1 — menu (full format)
// =========================================================================
function buildMenu(byDate) {
  const wb = XLSX.readFile(path.join(DATA, "Inskrivna & Matsedel.xlsx"), { cellDates: true });
  const sheetName = wb.SheetNames.find((n) => norm(n).includes("matsedel") && norm(n).includes("skolor"))
    || "Matsedel skolor";
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null, raw: true });
  const header = rows.find((r) => r && r.some((c) => norm(c).includes("datum")) && r.some((c) => norm(c).includes("veckodag")));
  if (!header) throw new Error("menu header not found");
  const ci = {};
  header.forEach((c, i) => { const n = norm(c); if (n.includes("datum")) ci.date = i; if (n.includes("veckodag")) ci.day = i; if (n.includes("alternativ") && !n.includes("vegetar")) ci.meal = i; if (n.includes("vegetar")) ci.veg = i; });

  const out = [];
  let minDate = null, maxDate = null;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i]; if (!r) continue;
    const d = cellDate(r[ci.date]); if (!d) continue;
    const meal = String(r[ci.meal] == null ? "" : r[ci.meal]).trim();
    if (!meal || meal === "-") continue;
    const veg = String(r[ci.veg] == null ? "" : r[ci.veg]).trim();
    const comps = Decompose.decomposeDish(meal) || {};
    const pseudo = { components: comps };
    out.push({
      date: iso(d),
      day: String(r[ci.day] == null ? "" : r[ci.day]).trim(),
      meal,
      veg: veg && veg !== "-" ? veg : "",
      protein: comps.protein || "",
      starch: comps.starch || "",
      vegs: D.vegsOf(pseudo) || "",
      sauce: comps.sauce || "",
      prep: comps.prep || "",
    });
    if (!minDate || d < minDate) minDate = d;
    if (!maxDate || d > maxDate) maxDate = d;
  }
  out.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));

  const r1 = (x) => Math.round(x * 10) / 10;
  const aoa = [["Datum", "Veckodag", "Enhet", "Måltid", "Vegetariskt",
    "Protein", "Kolhydrat", "Grönäsar", "Sås", "Tillagning",
    "Svinn/ätande (mätt)"]];
  let withMatt = 0;
  for (const m of out) {
    const agg = byDate && byDate.get(m.date);
    const matt = agg && agg.eaters && agg.waste != null ? r1((agg.waste / agg.eaters) * 1000) : "";
    if (matt !== "") withMatt++;
    aoa.push([m.date, m.day, "All", m.meal, m.veg, m.protein, m.starch, m.vegs, m.sauce, m.prep, matt]);
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 11 }, { wch: 10 }, { wch: 6 }, { wch: 46 }, { wch: 46 },
    { wch: 16 }, { wch: 14 }, { wch: 22 }, { wch: 14 }, { wch: 14 }, { wch: 18 }];

  const wbOut = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbOut, ws, "Matsedel");
  const readme = [
    ["Skolmat plan — meny (full format)"],
    [""],
    ["Källa: 'Inskrivna & Matsedel.xlsx', blad 'Matsedel skolor'."],
    ["Enhet = 'All'. Komponenter (Protein/Kolhydrat/Grönäsar/Sås/Tillagning) är"],
    ["upplåtna ur måltidens namn — kontrollera och rätta vid behov."],
    ["'Svinn/ätande (mätt)' = alla skolors totala svinn / totala ätande × 1000 för dagen"],
    ["(aggregerat ur Portioner + Svinn)."],
    ["Period: " + iso(minDate) + " → " + iso(maxDate) + " (ISO-vecka " + minDate.getFullYear() + "-v" + isoWeek(minDate) + "–" + maxDate.getFullYear() + "-v" + isoWeek(maxDate) + ")."],
  ];
  XLSX.utils.book_append_sheet(wbOut, XLSX.utils.aoa_to_sheet(readme), "README");
  return { wb: wbOut, minDate, maxDate, rows: out.length, withMatt };
}

// =========================================================================
// FILE 2 — svinn report (keyed by Enhet + Datum, Lunch only)
// =========================================================================
// Scan a workbook: for each sheet, find header row(s) with the given date +
// value markers, then collect data rows. Returns [{enhet, date, servering, val}].
function scanSheet(wb, { dateMark, valMark, valMarkAlt, serveringMark }) {
  const dm = norm(dateMark), vm = norm(valMark), vmAlt = norm(valMarkAlt), sm = serveringMark ? norm(serveringMark) : null;
  const recs = [];
  for (const sn of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: null, raw: true });
    let cDate = -1, cVal = -1, cServ = -1;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i]; if (!r) continue;
      const cells = r.map(norm);
      const isHeader = cells.some((c) => c.includes(dm)) &&
        (cells.some((c) => c.includes(vm)) || cells.some((c) => c.includes(vmAlt)));
      if (isHeader) {
        cDate = cells.findIndex((c) => c.includes(dm));
        cVal = cells.findIndex((c) => c.includes(vm));
        if (cVal < 0) cVal = cells.findIndex((c) => c.includes(vmAlt));
        cServ = sm ? cells.findIndex((c) => c.includes(sm)) : -1;
        continue;
      }
      if (cDate < 0) continue;
      const d = cellDate(r[cDate]); if (!d) continue;
      const val = num(r[cVal]); if (val == null) continue;
      const serv = cServ >= 0 ? r[cServ] : null;
      recs.push({ enhet: sn, date: iso(d), servering: serv, val });
    }
  }
  return recs;
}

function buildSvinnReport() {
  const portionerFile = fs.readdirSync(DATA).find((f) => norm(f).includes("portion"));
  const svinnFile = fs.readdirSync(DATA).find((f) => norm(f).includes("svinn") && f.endsWith(".xlsx"));
  const rapportFile = fs.readdirSync(DATA).find((f) => norm(f).includes("rapport") && f.endsWith(".xlsx"));

  // Portioner: eaters per (enhet, date), Lunch only.
  const pwb = XLSX.readFile(path.join(DATA, portionerFile), { cellDates: true });
  const portions = scanSheet(pwb, { dateMark: "datum", valMark: "antal ätande", valMarkAlt: "mängd", serveringMark: "servering" })
    .filter((r) => r.servering == null || LUNCH(r.servering));

  // Svinn: waste kg per (enhet, date), Lunch only, sum all svinntyp.
  const swb = XLSX.readFile(path.join(DATA, svinnFile), { cellDates: true });
  const wastes = scanSheet(swb, { dateMark: "datum", valMark: "vikt", valMarkAlt: "vikt", serveringMark: "servering" })
    .filter((r) => r.servering == null || LUNCH(r.servering));

  // Rapport: per-meal rows → aggregate to (enhet, date).
  const rwb = XLSX.readFile(path.join(DATA, rapportFile), { cellDates: true });
  const rrows = XLSX.utils.sheet_to_json(rwb.Sheets[rwb.SheetNames[0]], { header: 1, defval: null, raw: true });
  const rheader = rrows.find((r) => r && r.some((c) => norm(c).includes("serveringsdatum")));
  const rc = {};
  if (rheader) {
    rheader.forEach((c, i) => { const n = norm(c);
      if (n.includes("serveringsdatum")) rc.date = i;
      if (n === "enhet") rc.enhet = i;
      if (n.includes("atande")) rc.eaters = i;
      if (n.includes("totalt svinn")) rc.waste = i;
    });
  }
  const rapport = [];
  if (rc.date != null) {
    for (let i = 1; i < rrows.length; i++) {
      const r = rrows[i]; if (!r) continue;
      const d = cellDate(r[rc.date]); if (!d) continue;
      const enhet = String(r[rc.enhet] == null ? "" : r[rc.enhet]).trim(); if (!enhet) continue;
      const eaters = num(r[rc.eaters]); const waste = num(r[rc.waste]);
      if (eaters == null && waste == null) continue;
      rapport.push({ enhet, date: iso(d), eaters, waste });
    }
  }

  // Merge by (enhet, date).
  const map = new Map();
  const key = (e, d) => e + "\u0000" + d;
  const get = (e, d) => {
    const k = key(e, d);
    if (!map.has(k)) map.set(k, { enhet: e, date: d, eaters: null, waste: null });
    return map.get(k);
  };
  for (const p of portions) { const o = get(p.enhet, p.date); o.eaters = (o.eaters || 0) + p.val; }
  for (const w of wastes) { const o = get(w.enhet, w.date); o.waste = (o.waste || 0) + w.val; }
  // Rapport: eaters is the day's count (same across meals) → take max; waste sums.
  for (const rp of rapport) {
    const o = get(rp.enhet, rp.date);
    if (rp.eaters != null) o.eaters = o.eaters == null ? rp.eaters : Math.max(o.eaters, rp.eaters);
    if (rp.waste != null) o.waste = (o.waste || 0) + rp.waste;
  }

  const rows = Array.from(map.values()).sort((a, b) =>
    a.enhet < b.enhet ? -1 : a.enhet > b.enhet ? 1 : (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));

  const r1 = (x) => Math.round(x * 10) / 10;
  const aoa = [["Enhet", "Datum", "Antal ätande", "Totalt svinn (kg)", "Svinn/ätande (mätt)"]];
  let withWaste = 0;
  for (const o of rows) {
    const g = o.eaters && o.waste != null ? r1((o.waste / o.eaters) * 1000) : "";
    if (g !== "") withWaste++;
    aoa.push([o.enhet, o.date, o.eaters != null ? o.eaters : "", o.waste != null ? r1(o.waste) : "", g]);
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 24 }, { wch: 11 }, { wch: 14 }, { wch: 16 }, { wch: 20 }];

  // All-schools aggregate per date (for the menu's Svinn/ätande (mätt)).
  const byDate = new Map();
  for (const o of rows) {
    if (!byDate.has(o.date)) byDate.set(o.date, { eaters: 0, waste: 0, hasWaste: false });
    const a = byDate.get(o.date);
    if (o.eaters != null) a.eaters += o.eaters;
    if (o.waste != null) { a.waste += o.waste; a.hasWaste = true; }
  }
  for (const a of byDate.values()) if (!a.hasWaste) a.waste = null;

  const dates = rows.map((o) => o.date).sort();
  const minD = dates[0], maxD = dates[dates.length - 1];
  const wbOut = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbOut, ws, "Svinn report");
  const readme = [
    ["Svinn report — per Enhet + Datum (Lunch)"],
    [""],
    ["Källor: " + portionerFile + " + " + svinnFile + " + " + rapportFile + "."],
    ["Antal ätande: från Portioner (Lunch, summerat per Enhet+Datum)."],
    ["Totalt svinn (kg): från Svinn (Lunch, alla svinntyper summerade per Enhet+Datum)."],
    ["Svinn/ätande (mätt) g = Totalt svinn / Antal ätande × 1000 (när båda finns)."],
    ["Rapport-rader (per rätt) är sammanlagda till samma Enhet+Datum-nyckel."],
    ["Period: " + minD + " → " + maxD + "."],
    [""],
    ["Rader totalt: " + rows.length + " · med mätt svinn/ätande: " + withWaste + "."],
  ];
  XLSX.utils.book_append_sheet(wbOut, XLSX.utils.aoa_to_sheet(readme), "README");
  return { wb: wbOut, rows: rows.length, withWaste, minD, maxD, byDate };
}

// =========================================================================
// FILE 3 — svinn report from a rapport file (per Enhet + Datum, Lunch only)
// =========================================================================
// Used for periods Portioner/Svinn don't cover (e.g. rapport-2, Dec 2025).
// rapport rows are per-meal; with Lunch only, Enhet+Datum is unique.
function buildRapportReport(rapportFile) {
  const wb = XLSX.readFile(path.join(DATA, rapportFile), { cellDates: true });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null, raw: true });
  const header = rows.find((r) => r && r.some((c) => norm(c).includes("serveringsdatum")));
  const c = {};
  if (header) header.forEach((x, i) => { const n = norm(x);
    if (n.includes("serveringsdatum")) c.date = i;
    if (n === "enhet") c.enhet = i;
    if (n.includes("enhetstyp")) c.etyp = i;
    if (n.includes("maltidstyp")) c.maltid = i;
    if (n.includes("atande")) c.eaters = i;
    if (n.includes("totalt svinn")) c.waste = i;
  });
  const map = new Map();
  const key = (e, d) => e + "\u0000" + d;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i]; if (!r) continue;
    const date = String(r[c.date] == null ? "" : r[c.date]).trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) continue;
    if (c.maltid != null && norm(r[c.maltid]) !== "lunch") continue;
    if (c.etyp != null && !["grundskola", "gymnasieskola"].includes(norm(r[c.etyp]))) continue;
    const enhet = String(r[c.enhet] == null ? "" : r[c.enhet]).trim(); if (!enhet) continue;
    const eaters = num(r[c.eaters]); const waste = num(r[c.waste]);
    if (eaters == null && waste == null) continue;
    const k = key(enhet, date);
    if (!map.has(k)) map.set(k, { enhet, date, eaters: null, waste: null });
    const o = map.get(k);
    if (eaters != null) o.eaters = o.eaters == null ? eaters : Math.max(o.eaters, eaters);
    if (waste != null) o.waste = (o.waste || 0) + waste;
  }
  const out = Array.from(map.values()).sort((a, b) =>
    a.enhet < b.enhet ? -1 : a.enhet > b.enhet ? 1 : (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
  const dates = out.map((o) => o.date).sort();
  const minD = dates[0], maxD = dates[dates.length - 1];

  const r1 = (x) => Math.round(x * 10) / 10;
  const aoa = [["Enhet", "Datum", "Antal ätande", "Totalt svinn (kg)", "Svinn/ätande (mätt)"]];
  let withWaste = 0;
  for (const o of out) {
    const g = o.eaters && o.waste != null ? r1((o.waste / o.eaters) * 1000) : "";
    if (g !== "") withWaste++;
    aoa.push([o.enhet, o.date, o.eaters != null ? o.eaters : "", o.waste != null ? r1(o.waste) : "", g]);
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 24 }, { wch: 11 }, { wch: 14 }, { wch: 16 }, { wch: 20 }];

  const wbOut = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbOut, ws, "Svinn report");
  const readme = [
    ["Svinn report — per Enhet + Datum (Lunch)"],
    [""],
    ["Källa: " + rapportFile + " (per rätt, Lunch, Grundskola + Gymnasieskola, sammanlagd per Enhet+Datum)."],
    ["Antal ätande: 'Antal ätande (st)' från rapporten."],
    ["Totalt svinn (kg): 'Totalt svinn (kg)' från rapporten."],
    ["Svinn/ätande (mätt) g = Totalt svinn / Antal ätande × 1000."],
    ["Period: " + minD + " → " + maxD + "."],
    [""],
    ["Rader totalt: " + out.length + " · med mätt svinn/ätande: " + withWaste + "."],
  ];
  XLSX.utils.book_append_sheet(wbOut, XLSX.utils.aoa_to_sheet(readme), "README");
  return { wb: wbOut, rows: out.length, withWaste, minD, maxD };
}

// ---- write files ----
const rep = buildSvinnReport();
const repName = `svinn-report-${rep.minD}-${rep.maxD}.xlsx`;
XLSX.writeFile(rep.wb, path.join(OUT, repName));

const menu = buildMenu(rep.byDate);
const wk = (d) => `v${d.getFullYear()}-v${String(isoWeek(d)).padStart(2, "0")}`;
const menuName = `skolmat-${wk(menu.minDate)}-${wk(menu.maxDate)}-All.xlsx`;
XLSX.writeFile(menu.wb, path.join(OUT, menuName));

// Second svinn report from rapport-2 (period Portioner/Svinn don't cover).
const rep2 = buildRapportReport("rapport-2.xlsx");
const rep2Name = `svinn-report-${rep2.minD}-${rep2.maxD}.xlsx`;
XLSX.writeFile(rep2.wb, path.join(OUT, rep2Name));

console.log("Wrote:");
console.log("  data/processed/" + menuName + "  (" + menu.rows + " days, " + wk(menu.minDate) + "–" + wk(menu.maxDate) + ", " + menu.withMatt + " with mätt svinn/ätande)");
console.log("  data/processed/" + repName + "  (" + rep.rows + " Enhet+Datum rows, " + rep.withWaste + " with measured svinn/ätande)");
console.log("  data/processed/" + rep2Name + "  (" + rep2.rows + " Enhet+Datum rows, " + rep2.withWaste + " with measured svinn/ätande)");
