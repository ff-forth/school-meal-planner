/* import-real.js — in-browser importer for the three real Karlskrona .xlsx
   files (menu, waste, portions). Classifies files by filename, reads them with
   SheetJS, and joins them on date to produce the day records consumed by
   MealData.buildMealPool.

   The three files:
     1. Menu (…Matsedel…)      sheet "Matsedel skolor":
                                Datum | Veckodag | Alternativ 1 | Alternativ - Vegetariskt
     2. Waste (…Svinn…)        many sheets, one per kitchen; each has a header
                                row containing "Svinntyp" and "Vikt"; sum Vikt per date.
     3. Portions (…Portioner…) many sheets; header row containing "Servering" and
                                "Antal ätande" (or "Mängd"); sum the count per date.

   Waste per eater = (total waste kg for the day) / (total eaters that day) * 1000
   = grams of waste per eater. Lower is better. */
(function (global) {
  "use strict";

  // ---- diacritic-insensitive text helpers (for header matching) ----
  function norm(s) {
    return String(s == null ? "" : s)
      .toLowerCase()
      .normalize("NFD")
      .replace(/[̀-ͯ]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  // Convert a cell value (Date object or string) to a YYYY-MM-DD key.
  function toISODate(v) {
    if (v == null) return null;
    if (v instanceof Date && !isNaN(v)) {
      const y = v.getFullYear();
      const m = String(v.getMonth() + 1).padStart(2, "0");
      const d = String(v.getDate()).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }
    const s = String(v).trim();
    if (!s) return null;
    // ISO-ish: 2024-03-01 …
    let m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (m) {
      return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
    }
    // dd/mm/yyyy
    m = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
    if (m) {
      return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
    }
    return null;
  }

  // Find the header row: the first row whose normalized cells contain all of
  // `need`. Returns {row, header} or null.
  function findHeaderRow(rows, need) {
    for (let i = 0; i < rows.length; i++) {
      const cells = (rows[i] || []).map(norm);
      if (need.every((n) => cells.some((c) => c.includes(n)))) {
        return { row: i, header: rows[i] };
      }
    }
    return null;
  }

  // Index of the first column whose normalized header contains `name`.
  function colIndex(header, name) {
    const cells = (header || []).map(norm);
    for (let j = 0; j < cells.length; j++) {
      if (cells[j].includes(name)) return j;
    }
    return -1;
  }

  // Index of the first column whose normalized header EQUALS `name` (used to
  // tell "Enhet" apart from "Enhetstyp", which colIndex would match both).
  function colIndexExact(header, name) {
    const cells = (header || []).map(norm);
    for (let j = 0; j < cells.length; j++) {
      if (cells[j] === name) return j;
    }
    return -1;
  }

  function sheetRows(wb, name) {
    const ws = wb.Sheets[name];
    if (!ws) return null;
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  }

  // ---- classifiers by filename ----
  function classify(filename) {
    const n = norm(filename);
    if (n.includes("rapport")) return "rapport";
    if (n.includes("svinn-report")) return "svinnreport";
    if (n.includes("svinn")) return "waste";
    if (n.includes("portion")) return "portions";
    if (n.includes("skolmat-plan")) return "plan";
    if (n.includes("skolmat-pool")) return "pool";
    if (n.includes("matsedel") || n.includes("inskrivna") || n.includes("skolmat")) return "menu";
    return null;
  }

  // ---- readers ----
  function readMenu(wb) {
    // locate the "Matsedel skolor" sheet (or any sheet with the right header)
    let rows = null;
    for (const name of Object.keys(wb.Sheets)) {
      const r = sheetRows(wb, name);
      if (r && findHeaderRow(r, ["datum", "veckodag"])) { rows = r; break; }
    }
    if (!rows) return [];
    const { row, header } = findHeaderRow(rows, ["datum", "veckodag"]);
    const cDate = colIndex(header, "datum");
    const cDay = colIndex(header, "veckodag");
    const cA1 = colIndex(header, "alternativ 1") >= 0
      ? colIndex(header, "alternativ 1")
      : colIndex(header, "maltid"); // full-format menu uses "Måltid"
    const cVeg = colIndex(header, "vegetar");
    const out = [];
    for (let i = row + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const date = toISODate(r[cDate]);
      if (!date) continue;
      const name = String(r[cA1] == null ? "" : r[cA1]).trim();
      if (!name || name === "-") continue;
      const veg = String(r[cVeg] == null ? "" : r[cVeg]).trim();
      out.push({
        date,
        weekday: String(r[cDay] == null ? "" : r[cDay]).trim(),
        name,
        vegName: veg && veg !== "-" ? veg : "",
      });
    }
    return out;
  }

  function readWaste(wb) {
    const byDate = {};
    for (const name of Object.keys(wb.Sheets)) {
      const rows = sheetRows(wb, name);
      if (!rows) continue;
      const found = findHeaderRow(rows, ["svinntyp", "vikt"]);
      if (!found) continue;
      const cDate = colIndex(found.header, "datum");
      const cVikt = colIndex(found.header, "vikt");
      if (cDate < 0 || cVikt < 0) continue;
      for (let i = found.row + 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r) continue;
        const date = toISODate(r[cDate]);
        if (!date) continue;
        const v = parseFloat(r[cVikt]);
        if (isNaN(v)) continue;
        byDate[date] = (byDate[date] || 0) + v;
      }
    }
    return byDate; // { "YYYY-MM-DD": totalKg }
  }

  function readPortions(wb) {
    const byDate = {};
    for (const name of Object.keys(wb.Sheets)) {
      const rows = sheetRows(wb, name);
      if (!rows) continue;
      let found = findHeaderRow(rows, ["servering", "atande"]);
      if (!found) found = findHeaderRow(rows, ["servering", "mangd"]);
      if (!found) continue;
      const cDate = colIndex(found.header, "datum");
      let cCount = colIndex(found.header, "atande");
      if (cCount < 0) cCount = colIndex(found.header, "mangd");
      if (cDate < 0 || cCount < 0) continue;
      for (let i = found.row + 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r) continue;
        const date = toISODate(r[cDate]);
        if (!date) continue;
        const v = parseFloat(r[cCount]);
        if (isNaN(v)) continue;
        byDate[date] = (byDate[date] || 0) + v;
      }
    }
    return byDate; // { "YYYY-MM-DD": totalEaters }
  }

  // ---- clean svinn report (…svinn-report-YYYY-MM-DD-YYYY-MM-DD.xlsx) ----
  // One sheet "Svinn report": Enhet | Datum | Antal ätande | Totalt svinn (kg)
  // | Svinn/ätande (mätt). Rows are per unit+date; for the menu join we
  // aggregate to per-date totals (sum eaters and waste across units).
  function readSvinnReport(wb) {
    const byDate = { eaters: {}, waste: {} };
    for (const name of Object.keys(wb.Sheets)) {
      const rows = sheetRows(wb, name);
      if (!rows) continue;
      const found = findHeaderRow(rows, ["enhet", "datum", "atande", "svinn"]);
      if (!found) continue;
      const cDate = colIndex(found.header, "datum");
      const cEaters = colIndex(found.header, "atande");
      const cWaste = colIndex(found.header, "totalt svinn");
      if (cDate < 0) continue;
      for (let i = found.row + 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r) continue;
        const date = toISODate(r[cDate]);
        if (!date) continue;
        const e = cEaters >= 0 ? parseFloat(r[cEaters]) : NaN;
        const w = cWaste >= 0 ? parseFloat(r[cWaste]) : NaN;
        if (isFinite(e)) byDate.eaters[date] = (byDate.eaters[date] || 0) + e;
        if (isFinite(w)) byDate.waste[date] = (byDate.waste[date] || 0) + w;
      }
    }
    return byDate; // { eaters: {date: n}, waste: {date: kg} }
  }

  // ---- saved 8-week plan (…skolmat-plan-….xlsx, sheet "8-veckorsplan") ----
  // Header: Vecka | Dag | Datum | Måltid | Vegetariskt | Svinn g/ätande.
  // Returns the raw rows (week, weekday, date, meal, veg, waste) so the UI can
  // rebuild the plan grid and match meals against the loaded pool.
  function readPlan(wb) {
    let rows = null;
    for (const name of Object.keys(wb.Sheets)) {
      const r = sheetRows(wb, name);
      if (r && findHeaderRow(r, ["vecka", "dag", "datum", "maltid"])) { rows = r; break; }
    }
    if (!rows) return [];
    const { row, header } = findHeaderRow(rows, ["vecka", "dag", "datum", "maltid"]);
    const cWeek = colIndex(header, "vecka");
    const cDay = colIndex(header, "dag");
    const cDate = colIndex(header, "datum");
    const cMeal = colIndex(header, "maltid");
    const cVeg = colIndex(header, "vegetar");
    const cWaste = colIndex(header, "svinn g/atande");
    const out = [];
    for (let i = row + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const meal = String(r[cMeal] == null ? "" : r[cMeal]).trim();
      if (!meal) continue;
      out.push({
        week: parseInt(r[cWeek], 10),
        weekday: String(r[cDay] == null ? "" : r[cDay]).trim(),
        date: toISODate(r[cDate]),
        meal,
        veg: String(r[cVeg] == null ? "" : r[cVeg]).trim(),
        waste: cWaste >= 0 ? parseFloat(r[cWaste]) : NaN,
      });
    }
    return out;
  }

  // ---- pool file (…skolmat-pool…) ----
  // One row per unique meal: Måltid | Vegetariskt | Protein | Kolhydrat |
  // Grönäsar | Sås | Tillagning | Svinn/ätande (mätt) | Svinn/ätande (estimerat).
  // Rebuilds the candidate pool so a saved plan can be resumed without the full
  // menu. Waste is measured when present, else the estimate column.
  function readPool(wb) {
    let rows = null;
    for (const name of Object.keys(wb.Sheets)) {
      const r = sheetRows(wb, name);
      if (r && findHeaderRow(r, ["maltid", "protein", "svinn"])) { rows = r; break; }
    }
    if (!rows) return [];
    const { row, header } = findHeaderRow(rows, ["maltid", "protein", "svinn"]);
    const cMeal = colIndex(header, "maltid");
    const cVeg = colIndex(header, "vegetar");
    const cMatt = colIndexExact(header, "svinn/atande (matt)");
    const cEst = colIndexExact(header, "svinn/atande (estimerat)");
    // Component columns the user may have edited in the pool file. When present
    // they override the components the app would otherwise re-derive from the
    // dish name, so tag/component edits carry through on import.
    const cProtein = colIndexExact(header, "protein");
    const cStarch = colIndexExact(header, "kolhydrat");
    const cVegs = colIndexExact(header, "gronasar");
    const cSauce = colIndexExact(header, "sas");
    const cPrep = colIndexExact(header, "tillagning");
    const cell = (r, c) => (c >= 0 ? String(r[c] == null ? "" : r[c]).trim() : "");
    const out = [];
    for (let i = row + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const meal = String(r[cMeal] == null ? "" : r[cMeal]).trim();
      if (!meal) continue;
      const m = cMatt >= 0 ? parseFloat(r[cMatt]) : NaN;
      const est = cEst >= 0 ? parseFloat(r[cEst]) : NaN;
      const waste = isFinite(m) ? m : (isFinite(est) ? est : null);
      const protein = cell(r, cProtein);
      const starch = cell(r, cStarch);
      const vegs = cell(r, cVegs);
      const sauce = cell(r, cSauce);
      const prep = cell(r, cPrep);
      const rec = {
        name: meal,
        vegName: cVeg >= 0 ? String(r[cVeg] == null ? "" : r[cVeg]).trim() : "",
        wasteGPerEater: waste,
      };
      // Only attach explicit components when the row actually has them, so a
      // pool file without the columns still falls back to name-derived ones.
      if (protein || starch || vegs || sauce || prep) {
        rec.components = {
          protein: protein || null,
          starch: starch || null,
          sauce: sauce || null,
          prep: prep || null,
          vegs: vegs || null,
        };
      }
      out.push(rec);
    }
    return out;
  }

  // ---- per-school waste record (…rapport… "Hantera livs" export) ----
  // One row per school-day: Serveringsdatum | Enhet (school) | Antal ätande |
  // Totalt svinn (kg) | Maträtter (dish name, prefixed "<Enhetstyp>/"). Each
  // row yields a measured waste-per-eater for that school's dish on that date.
  // The dish name is cleaned of its "<type>/" prefix and trailing dots, then
  // decomposed into components so waste can be looked up by combination later.
  function readRapport(wb) {
    const records = [];
    for (const name of Object.keys(wb.Sheets)) {
      const rows = sheetRows(wb, name);
      if (!rows) continue;
      const found = findHeaderRow(rows, ["serveringsdatum", "matratter"]);
      if (!found) continue;
      const cDate = colIndex(found.header, "serveringsdatum");
      const cSchool = colIndexExact(found.header, "enhet");
      const cEaters = colIndex(found.header, "atande");
      const cWaste = colIndex(found.header, "totalt svinn");
      const cMeal = colIndex(found.header, "matratter");
      if (cDate < 0 || cSchool < 0 || cEaters < 0 || cWaste < 0 || cMeal < 0) continue;
      for (let i = found.row + 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r) continue;
        const date = toISODate(r[cDate]);
        if (!date) continue;
        const raw = String(r[cMeal] == null ? "" : r[cMeal]).trim();
        if (!raw) continue;
        const name = raw.replace(/^[^/]+\/\s*/, "").replace(/\s*\.\s*$/, "").trim();
        if (!name) continue;
        const eaters = parseFloat(r[cEaters]);
        const waste = parseFloat(r[cWaste]);
        if (!isFinite(eaters) || eaters <= 0 || !isFinite(waste)) continue;
        records.push({
          date,
          school: String(r[cSchool] == null ? "" : r[cSchool]).trim(),
          name,
          wasteGPerEater: Math.round((waste / eaters) * 1000 * 10) / 10,
        });
      }
    }
    return records; // [{date, school, name, wasteGPerEater}]
  }

  // ---- join ----
  // menuDays: [{date, weekday, name, vegName}]
  // wasteKg / eaters: { date: number }
  function join(menuDays, wasteKg, eaters) {
    const records = [];
    let joined = 0;
    for (const d of menuDays) {
      const w = wasteKg[d.date];
      const e = eaters[d.date];
      let wasteGPerEater = null;
      if (w != null && e && e > 0) {
        wasteGPerEater = Math.round((w / e) * 1000 * 10) / 10;
        joined++;
      }
      records.push({ ...d, wasteGPerEater });
    }
    const vals = records.map((r) => r.wasteGPerEater).filter((v) => v != null);
    const stats = vals.length
      ? {
          min: Math.min.apply(null, vals),
          max: Math.max.apply(null, vals),
          mean: Math.round((vals.reduce((a, b) => a + b, 0) / vals.length) * 10) / 10,
        }
      : { min: null, max: null, mean: null };
    return { records, joined, total: menuDays.length, stats };
  }

  // ---- top-level: read an ArrayBuffer of one file, return partial state ----
  function readWorkbook(arrayBuffer) {
    return XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  }

  // ---- skolmat.info JSON import ----
  // Accepts either a single week object ({ok, facility, week:{year,weekNumber,days}})
  // or an array of several week objects (as produced by the bookmarklet).
  // Maps each open day's rows into day records; pickRows() decides which row is
  // the regular dish and which (if any) is the vegetarian alternative.
  const SKOLMAT_DAY_KEYS = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"];
  // Positive keywords that mark a dish as vegetarian. We only flag a row as veg
  // when there is explicit plant/veg evidence, so a meat dish is never wrongly
  // put in the "Vegetariskt" slot. When neither row qualifies, vegName stays empty.
  const VEG_KEYWORDS = ["balj", "falafel", "soja", "kikart", "vego", "veg", "groda",
    "grodnings", "gronsak", "rotfrukt", "linser", "champignon", "svamp"];
  function isVegDish(name) {
    const n = norm(name);
    return VEG_KEYWORDS.some((k) => n.includes(k));
  }
  // From a day's rows, pick the regular dish (first non-veg, else first) and the
  // vegetarian alternative (first veg row, if different from the regular one).
  function pickRows(rows) {
    const veg = rows.find(isVegDish);
    const name = rows.find((r) => !isVegDish(r)) || rows[0];
    const vegName = veg && veg !== name ? veg : "";
    return { name, vegName };
  }
  function mondayOfISO(y, w) {
    const j = new Date(Date.UTC(y, 0, 4));
    j.setUTCDate(j.getUTCDate() - ((j.getUTCDay() + 6) % 7));
    const m = new Date(j);
    m.setUTCDate(m.getUTCDate() + (w - 1) * 7);
    const p = (n) => String(n).padStart(2, "0");
    return `${m.getUTCFullYear()}-${p(m.getUTCMonth() + 1)}-${p(m.getUTCDate())}`;
  }
  function parseSkolmatJSON(text) {
    let data;
    try { data = JSON.parse(text); }
    catch (err) { return { records: [], error: "Ogiltig JSON: " + err.message }; }
    const weeks = Array.isArray(data) ? data : [data];
    const records = [];
    let facilityName = null;
    for (const wk of weeks) {
      if (!wk || !wk.week || !wk.week.days) continue;
      if (wk.facility && wk.facility.name) facilityName = wk.facility.name;
      const year = wk.week.year, weekNumber = wk.week.weekNumber;
      if (!year || !weekNumber) continue;
      const base = mondayOfISO(year, weekNumber);
      for (let d = 0; d < SKOLMAT_DAY_KEYS.length; d++) {
        const day = wk.week.days[SKOLMAT_DAY_KEYS[d]];
        if (!day || !day.isOpen) continue;
        const rows = (day.rows || []).map((r) => String(r && r.text ? r.text : "").trim()).filter(Boolean);
        if (!rows.length) continue;
        const dt = new Date(base);
        dt.setUTCDate(dt.getUTCDate() + d);
        const p = (n) => String(n).padStart(2, "0");
        const date = `${dt.getUTCFullYear()}-${p(dt.getUTCMonth() + 1)}-${p(dt.getUTCDate())}`;
        const picked = pickRows(rows);
        records.push({
          date,
          weekday: SKOLMAT_DAY_KEYS[d],
          name: picked.name,
          vegName: picked.vegName,
          wasteGPerEater: null,
        });
      }
    }
    if (!records.length) return { records: [], error: "Inga dagar hittades i filen." };
    return { records, facilityName };
  }

  // Accumulator for a set of uploaded files.
  function createLoader() {
    return { menuDays: [], wasteKg: {}, eaters: {}, rapport: [], stats: {} };
  }

  function addFile(loader, filename, arrayBuffer) {
    const kind = classify(filename);
    if (!kind) return { kind: null, count: 0, error: "unrecognised file name" };
    const wb = readWorkbook(arrayBuffer);
    let count = 0;
    if (kind === "rapport") {
      const recs = readRapport(wb);
      for (const r of recs) loader.rapport.push(r);
      count = recs.length;
      loader.stats.rapport = loader.rapport.length;
    } else if (kind === "menu") {
      // Accumulate across multiple matsedel files. Dedupe by date so two
      // overlapping menus don't double-count a day (later file wins); menus
      // for different weeks both survive.
      const days = readMenu(wb);
      const byDate = new Map((loader.menuDays || []).map((d) => [d.date, d]));
      for (const d of days) byDate.set(d.date, d);
      loader.menuDays = Array.from(byDate.values()).sort((a, b) =>
        a.date < b.date ? -1 : a.date > b.date ? 1 : 0
      );
      count = loader.menuDays.length;
      loader.stats.menu = count;
    } else if (kind === "svinnreport") {
      // Clean per-unit report: aggregate eaters + waste to per-date totals.
      const sr = readSvinnReport(wb);
      for (const k of Object.keys(sr.waste)) loader.wasteKg[k] = (loader.wasteKg[k] || 0) + sr.waste[k];
      for (const k of Object.keys(sr.eaters)) loader.eaters[k] = (loader.eaters[k] || 0) + sr.eaters[k];
      count = Math.max(Object.keys(sr.waste).length, Object.keys(sr.eaters).length);
      loader.stats.wasteDates = Object.keys(loader.wasteKg).length;
      loader.stats.eaterDates = Object.keys(loader.eaters).length;
    } else if (kind === "waste") {
      const w = readWaste(wb);
      for (const k of Object.keys(w)) loader.wasteKg[k] = (loader.wasteKg[k] || 0) + w[k];
      count = Object.keys(w).length;
      loader.stats.wasteDates = Object.keys(loader.wasteKg).length;
    } else if (kind === "portions") {
      const e = readPortions(wb);
      for (const k of Object.keys(e)) loader.eaters[k] = (loader.eaters[k] || 0) + e[k];
      count = Object.keys(e).length;
      loader.stats.eaterDates = Object.keys(loader.eaters).length;
    }
    if (count === 0) return { kind, count, error: "no usable rows found" };
    return { kind, count };
  }

  function finish(loader) {
    const out = join(loader.menuDays, loader.wasteKg, loader.eaters);
    out.rapport = loader.rapport || [];
    return out;
  }

  global.RealImport = {
    classify,
    toISODate,
    findHeaderRow,
    colIndex,
    readMenu,
    readWaste,
    readPortions,
    readSvinnReport,
    readPlan,
    readPool,
    readRapport,
    colIndexExact,
    join,
    readWorkbook,
    createLoader,
    addFile,
    finish,
    parseSkolmatJSON,
    mondayOfISO,
    isVegDish,
    pickRows,
  };
})(window);
