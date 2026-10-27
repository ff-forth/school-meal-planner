/* ui.js — wire up the real-data loader, criteria, grid, week cards, popover, status */
(function (global) {
  "use strict";
  const D = global.MealData;
  const IO = global.XlsxIO;
  const RI = global.RealImport;
  const M = global.Mining;
  const P = global.Planner;

  const app = {
    dayRecords: null,
    pool: null,
    mined: null,
    criteria: null,
    plan: null,       // plan[w][d] = {date, meal}
    state: null,      // {plan, pool, uses, lastUsed}
    eval: null,
    viewWeek: "all",  // "all" or a week index (0-based)
    compRules: null,  // structured component-rule rows
    wasteLookup: null, // {pooled, schools, schoolNames, count} from a rapport file
    targetSchool: "",  // school to prefer for waste lookup ("" = pooled)
  };

  const $ = (id) => document.getElementById(id);

  // ---------- Swedish labels ----------
  const DAY_SV = ["Måndag", "Tisdag", "Onsdag", "Torsdag", "Fredag"];
  const MONTH_SV = ["januari", "februari", "mars", "april", "maj", "juni",
    "juli", "augusti", "september", "oktober", "november", "december"];
  function formatDateSV(iso) {
    if (!iso) return "";
    const [y, m, d] = iso.split("-").map(Number);
    return `${d} ${MONTH_SV[m - 1]}`;
  }

  // ISO-8601 calendar week number (1-53), like vecka.nu.
  function isoWeek(iso) {
    if (!iso) return null;
    const [y, m, d] = iso.split("-").map(Number);
    const dt = new Date(Date.UTC(y, m - 1, d));
    // Shift to the Thursday of this week; ISO weeks start on Monday.
    dt.setUTCDate(dt.getUTCDate() + 4 - (dt.getUTCDay() || 7));
    const yearStart = Date.UTC(dt.getUTCFullYear(), 0, 1);
    return Math.ceil(((dt - yearStart) / 86400000 + 1) / 7);
  }

  // ---------- Category badge (protein → Fisk/Kyckling/Kött/Vegetariskt) ----------
  const ICONS = {
    salad: '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M7 21h10"/><path d="M12 21a9 9 0 0 0 9-9H3a9 9 0 0 0 9 9Z"/><path d="M11.38 12a2.4 2.4 0 0 1-.4-4.77 2.4 2.4 0 0 1 3.2-2.77 2.4 2.4 0 0 1 3.47-.63 2.4 2.4 0 0 1 3.37 3.37 2.4 2.4 0 0 1-1.1 3.7 2.51 2.51 0 0 1 .03 1.1"/><path d="m13 12 4-4"/><path d="M10.9 7.25A3.99 3.99 0 0 0 4 10c0 .73.2 1.41.54 2"/></svg>',
    fish: '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6.5 12c.94-3.46 4.94-6 8.5-6 3.56 0 6.06 2.54 7 6-.94 3.47-3.44 6-7 6s-7.56-2.53-8.5-6Z"/><path d="M18 12v.5"/><path d="M16 17.93a9.77 9.77 0 0 1 0-11.86"/><path d="M7 10.67C7 8 5.58 5.97 2.73 5.5c-1 1.5-1 5 .23 6.5-1.24 1.5-1.24 5-.23 6.5C5.58 18.03 7 16 7 13.33"/></svg>',
    drumstick: '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15.4 15.63a7.875 6 135 1 1 6.23-6.23 4.5 3.43 135 0 0-6.23 6.23"/><path d="m8.29 12.71-2.6 2.6a2.5 2.5 0 1 0-1.65 4.65A2.5 2.5 0 1 0 8.7 18.3l2.59-2.59"/></svg>',
    beef: '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3c5 0 9 3 9 7s-4 7-9 7-9-3-9-7 4-7 9-7Z"/><path d="M9 10c1.2-1.2 3.4-1.2 4.6 0"/><path d="M9 14c1.2-1.2 3.4-1.2 4.6 0"/></svg>',
    pork: '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="9" width="20" height="6" rx="3"/><path d="M8 9v6"/><path d="M16 9v6"/></svg>',
    leaf: '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 20A7 7 0 0 1 9.8 6.1C15.5 5 17 4.48 19 2c1 2 2 4.18 2 8 0 5.5-4.78 10-10 10Z"/><path d="M2 21c0-3 1.85-5.36 5.08-6C9.5 14.52 12 13 13 12"/></svg>',
    wheat: '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 22 16 8"/><path d="M3.47 12.53 5 11l1.5 1.5a1 1 0 0 1 0 1.4l-1.5 1.5a1 1 0 0 1-1.4 0L2.1 13.9a1 1 0 0 1 0-1.4Z"/><path d="M7.47 8.53 9 7l1.5 1.5a1 1 0 0 1 0 1.4L9 11.4a1 1 0 0 1-1.4 0L6.1 9.9a1 1 0 0 1 0-1.4Z"/><path d="M11.47 4.53 13 3l1.5 1.5a1 1 0 0 1 0 1.4L13 7.4a1 1 0 0 1-1.4 0L10.1 5.9a1 1 0 0 1 0-1.4Z"/><path d="m12.5 12.5 2.1 2.1a1 1 0 0 1 0 1.4l-2.1 2.1a1 1 0 0 1-1.4 0L9 16"/><path d="m14.5 6.5 2.1 2.1a1 1 0 0 1 0 1.4L14.5 12a1 1 0 0 1-1.4 0L11 9.9"/></svg>',
    soup: '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 21a9 9 0 0 0 9-9H3a9 9 0 0 0 9 9Z"/><path d="M7 21h10"/><path d="M7 3v2"/><path d="M12 2v3"/><path d="M17 3v2"/></svg>',
  };
  function categoryOf(meal) {
    const p = (meal.components && meal.components.protein ? meal.components.protein : "");
    if (/fisk|lax|tunfisk|skaldjur|krabba/.test(p)) return { label: "Fisk", icon: "fish", cls: "badge-info" };
    if (/kyckling|fågel|pullad|gyros/.test(p)) return { label: "Kyckling", icon: "drumstick", cls: "badge-neutral" };
    if (/soja|vego|balj|kikarts|linse|falafel|ost/.test(p)) return { label: "Vegetariskt", icon: "leaf", cls: "badge-success" };
    // pork (griskött) vs beef (nötkött) — these are distinct and must be shown separately
    if (/flask|falukorv|korv|wienerkorv|blodpudding/.test(p)) return { label: "Fläsk", icon: "pork", cls: "badge-warning" };
    if (/kott|farsbiff|kalops|biff/.test(p)) return { label: "Nötkött", icon: "beef", cls: "badge-neutral" };
    if (p) return { label: "Kött", icon: "beef", cls: "badge-neutral" };
    return null;
  }
  function chipHTML(icon, label, cls) {
    return `<span class="menu-editor-compact-allergen-chip ${cls || ""}">${ICONS[icon] || ""}${label}</span>`;
  }

  // ---------- Starch badge (ris/potatis/pasta/bröd) — for quick skimming ----------
  function starchOf(meal) {
    const s = (meal.components && meal.components.starch ? meal.components.starch : "").toLowerCase();
    if (!s) return null;
    if (/ris|risotto|nasi/.test(s)) return { label: "Ris", icon: "wheat", cls: "badge-info" };
    if (/potatis|mos|solskensmos|klyft/.test(s)) return { label: "Potatis", icon: "wheat", cls: "badge-warning" };
    if (/pasta|spagetti|nudlar|makaroner/.test(s)) return { label: "Pasta", icon: "wheat", cls: "badge-neutral" };
    if (/lasagne|lasangette|lasagnette/.test(s)) return { label: "Lasagne", icon: "wheat", cls: "badge-neutral" };
    if (/bulgur/.test(s)) return { label: "Bulgur", icon: "wheat", cls: "badge-neutral" };
    if (/brod|tortilla|paella|knacke/.test(s)) return { label: "Bröd", icon: "wheat", cls: "badge-neutral" };
    return { label: s, icon: "wheat", cls: "badge-neutral" };
  }

  // ---------- Soup badge (soppa) — for quick skimming ----------
  function soupOf(meal) {
    const prep = (meal.components && meal.components.prep ? meal.components.prep : "").toLowerCase();
    const name = (meal.name ? meal.name : "").toLowerCase();
    if (/soppa/.test(prep) || /soppa/.test(name)) return { label: "Soppa", icon: "soup", cls: "badge-info" };
    return null;
  }

  // ---------- Plan-wide component tally (for the "Alla" hover summary) ----------
  // Counts how often each category label (protein / starch / soup / theme)
  // appears across the whole plan, e.g. { "Fläsk": 5, "Fisk": 8, "Potatis": 12 }.
  function planCategoryCounts() {
    const counts = {};
    const bump = (label) => { if (label) counts[label] = (counts[label] || 0) + 1; };
    for (let w = 0; w < D.WEEKS; w++)
      for (let d = 0; d < D.DAYS.length; d++) {
        const m = app.plan[w][d].meal;
        if (!m) continue;
        const cat = categoryOf(m); if (cat) bump(cat.label);
        const st = starchOf(m); if (st) bump(st.label);
        const sp = soupOf(m); if (sp) bump(sp.label);
        if (m.components && m.components.theme) bump(m.components.theme);
      }
    return counts;
  }

  // The component labels a meal belongs to (same set as planCategoryCounts).
  function cellCompLabels(meal) {
    const labels = [];
    const cat = categoryOf(meal); if (cat) labels.push(cat.label);
    const st = starchOf(meal); if (st) labels.push(st.label);
    const sp = soupOf(meal); if (sp) labels.push(sp.label);
    if (meal.components && meal.components.theme) labels.push(meal.components.theme);
    return labels;
  }

  // "Fläsk 5×, Potatis 12×, Soppa 8×" for the components a meal belongs to.
  function mealSummary(meal, counts) {
    const parts = [];
    const cat = categoryOf(meal); if (cat) parts.push(`${cat.label} ${counts[cat.label] || 0}×`);
    const st = starchOf(meal); if (st) parts.push(`${st.label} ${counts[st.label] || 0}×`);
    const sp = soupOf(meal); if (sp) parts.push(`${sp.label} ${counts[sp.label] || 0}×`);
    if (meal.components && meal.components.theme) parts.push(`${meal.components.theme} ${counts[meal.components.theme] || 0}×`);
    return parts.join(", ");
  }

  // ---------- Criteria ----------
  // ---------- Component rules (fixed structured rows) ----------
  const COMP_RULE_DEFAULTS = [
    { name: "Fläsk", kind: "max", count: 2, period: "week", day: "" },
    { name: "Korv", kind: "max", count: 1, period: "14d", day: "" },
    { name: "Ris", kind: "max", count: 1, period: "14d", day: "" },
    { name: "Soppa", kind: "exact", count: 1, period: "week", day: "" },
    { name: "Fisk", kind: "min", count: 1, period: "week", day: "" },
    { name: "Kyckling", kind: "min", count: 1, period: "week", day: "" },
  ];

  // Day-of-week options ("" = any day). Used for the optional preferred day.
  function compRuleDayOptions(selected) {
    const opts = [`<option value=""${!selected ? " selected" : ""}>–</option>`];
    for (let d = 0; d < D.DAYS.length; d++) {
      opts.push(`<option value="${d}"${String(selected) === String(d) ? " selected" : ""}>${DAY_SV[d]}</option>`);
    }
    return opts.join("");
  }

  function compRuleRowHTML(r) {
    const esc = (s) => String(s || "").replace(/"/g, "&quot;");
    return `<div class="comp-rule">` +
      `<input type="text" class="inp comp-rule-name" value="${esc(r.name)}" placeholder="Komponent" />` +
      `<select class="inp comp-rule-kind">` +
        `<option value="max"${r.kind === "max" ? " selected" : ""}>max</option>` +
        `<option value="min"${r.kind === "min" ? " selected" : ""}>minst</option>` +
        `<option value="exact"${r.kind === "exact" ? " selected" : ""}>exakt</option>` +
      `</select>` +
      `<input type="number" class="inp comp-rule-count" min="1" step="1" value="${r.count || 1}" />` +
      `<select class="inp comp-rule-period">` +
        `<option value="week"${r.period === "week" ? " selected" : ""}>vecka</option>` +
        `<option value="14d"${r.period === "14d" ? " selected" : ""}>14 dagar</option>` +
      `</select>` +
      `<select class="inp comp-rule-day" title="Önskad dag (valfritt)">${compRuleDayOptions(r.day)}</select>` +
      `<button type="button" class="comp-rule-remove" title="Ta bort" aria-label="Ta bort">×</button>` +
      `</div>`;
  }

  // Capture the current DOM rows back into app.compRules (preserves in-progress
  // edits so a re-render doesn't wipe them).
  function syncCompRulesFromDOM() {
    const wrap = $("c-component-rules");
    if (!wrap || !wrap.children.length) return;
    app.compRules = Array.from(wrap.querySelectorAll(".comp-rule")).map((row) => ({
      name: (row.querySelector(".comp-rule-name").value || "").trim(),
      kind: row.querySelector(".comp-rule-kind").value,
      count: parseInt(row.querySelector(".comp-rule-count").value, 10) || 1,
      period: row.querySelector(".comp-rule-period").value,
      day: row.querySelector(".comp-rule-day").value,
    }));
  }

  function renderComponentRules() {
    const wrap = $("c-component-rules");
    if (!wrap) return;
    wrap.innerHTML = "";
    for (const r of app.compRules) {
      const div = document.createElement("div");
      div.innerHTML = compRuleRowHTML(r);
      wrap.appendChild(div.firstChild);
    }
    wrap.querySelectorAll(".comp-rule-remove").forEach((btn) => {
      btn.addEventListener("click", () => {
        const row = btn.closest(".comp-rule");
        const idx = Array.from(wrap.children).indexOf(row);
        syncCompRulesFromDOM(); // keep edits before rebuilding
        app.compRules.splice(idx, 1);
        renderComponentRules();
      });
    });
  }

  // Read the structured rows back into rule objects.
  function readComponentRules() {
    const wrap = $("c-component-rules");
    if (!wrap) return [];
    const rules = [];
    for (const row of wrap.querySelectorAll(".comp-rule")) {
      const name = (row.querySelector(".comp-rule-name").value || "").trim();
      const kind = row.querySelector(".comp-rule-kind").value;
      const count = parseInt(row.querySelector(".comp-rule-count").value, 10);
      const period = row.querySelector(".comp-rule-period").value;
      const dayRaw = row.querySelector(".comp-rule-day").value;
      const day = dayRaw === "" ? null : parseInt(dayRaw, 10);
      if (!name || !count) continue;
      rules.push({ name, kind, count, period, day });
    }
    return rules;
  }

  function readCriteria() {
    return P.parseCriteria({
      minGapWeeks: $("c-gap").value,
      maxUses: $("c-maxuse").value,
      preferLowWaste: $("c-prefer-waste").checked,
      maxAvgWaste: $("c-max-waste").value,
      excludedAllergens: "",
      required: "",
      componentRules: readComponentRules(),
    });
  }

  // ---------- Real-data loader (multi-file, auto-classified by name) ----------
  function fileStatusEl(kind) {
    return document.querySelector(`#real-file-status li[data-kind="${kind}"]`);
  }
  function markFileStatus(kind, ok, label) {
    const li = fileStatusEl(kind);
    if (!li) return;
    li.textContent = `${ok ? "✓" : "○"} ${label}`;
    li.style.color = ok ? "var(--primary)" : "";
  }

  // Rebuild an 8-week plan grid from a saved plan file's rows and show it.
  // Meals are matched against the loaded pool by name (so waste/components
  // carry over); unmatched names become minimal meals.
  function loadPlanFromRows(rows, label) {
    const status = $("status-real");
    if (!rows.length) {
      status.textContent = `Ingen plan hittades i ${label}.`;
      status.style.color = "var(--danger)";
      return;
    }
    const byName = new Map((app.pool || []).map((m) => [m.name, m]));
    const plan = [];
    let matched = 0;
    for (const r of rows) {
      const w = (r.week || 1) - 1;
      if (w < 0 || w >= D.WEEKS) continue;
      // Accept Swedish (app export) or English (D.DAYS) day names.
      const d = DAY_SV.indexOf(r.weekday) >= 0 ? DAY_SV.indexOf(r.weekday) : D.DAYS.indexOf(r.weekday);
      if (d < 0) continue;
      if (!plan[w]) plan[w] = [];
      const meal = byName.get(r.meal);
      if (meal) matched++;
      plan[w][d] = {
        date: r.date || "",
        meal: meal || { name: r.meal, vegName: r.veg || "", components: {}, wasteGPerEater: isFinite(r.waste) ? r.waste : null },
      };
    }
    if (!plan.some((wk) => wk && wk.length)) {
      status.textContent = `Ingen plan hittades i ${label}.`;
      status.style.color = "var(--danger)";
      return;
    }
    app.plan = plan;
    app.state = { plan, pool: app.pool, uses: new Map(), lastUsed: new Map() };
    app.viewWeek = "all";
    renderWeekTabs();
    renderPlan();
    showPlanPage();
    status.textContent = `✓ Plan laddad (${label}) · ${matched}/${rows.length} rätter matchade mot menyn.`;
    status.style.color = "var(--primary)";
  }

  // Rebuild the candidate pool from a saved pool file's records and show it.
  // Components are re-derived from each meal name; measured waste is kept, and
  // the component-based estimate fills in where it is missing.
  function loadPoolFromRows(rows, label) {
    const status = $("status-real");
    if (!rows.length) {
      status.textContent = `Ingen pool hittades i ${label}.`;
      status.style.color = "var(--danger)";
      return;
    }
    app.dayRecords = rows.map((r) => ({
      name: r.name, vegName: r.vegName || "", wasteGPerEater: r.wasteGPerEater != null ? r.wasteGPerEater : null,
    }));
    app.pool = D.buildMealPool(app.dayRecords);
    applyEstimates();
    $("drop-real").classList.add("ok");
    markFileStatus("pool", true, `${label} (${app.pool.length} måltider)`);
    status.textContent = `✓ Pool laddad (${label}) · ${app.pool.length} unika måltider`;
    status.style.color = "var(--primary)";
    checkReady();
  }

  async function handleRealFiles(files) {
    const status = $("status-real");
    const loader = RI.createLoader();
    // Split .json (skolmat.info) files from .xlsx files. JSON files are the
    // bookmarklet's output and are parsed separately; .xlsx files are
    // classified by name (menu / svinn / portioner / rapport).
    const jsonFiles = files.filter((f) => /\.json$/i.test(f.name));
    const xlsxFiles = files.filter((f) => !/\.json$/i.test(f.name));
    for (const f of jsonFiles) {
      try {
        const out = RI.parseSkolmatJSON(await f.text());
        if (!out.records.length) {
          markFileStatus("menu", false, `${f.name} — ${out.error || "inga dagar"}`);
        } else {
          // Merge by date (same rule as the xlsx menu path) so overlapping
          // files don't double-count a day.
          const byDate = new Map((loader.menuDays || []).map((d) => [d.date, d]));
          for (const r of out.records) byDate.set(r.date, r);
          loader.menuDays = Array.from(byDate.values()).sort((a, b) =>
            a.date < b.date ? -1 : a.date > b.date ? 1 : 0
          );
          loader.jsonFacility = out.facilityName || loader.jsonFacility;
          markFileStatus("menu", true, `${f.name} (${out.records.length} dagar)`);
        }
      } catch (err) {
        markFileStatus("menu", false, `${f.name} — ${err.message}`);
      }
    }
    for (const f of xlsxFiles) {
      const kind = RI.classify(f.name);
      try {
        const buf = await IO.readArrayBuffer(f);
        if (kind === "plan") {
          // A saved plan is loaded on its own (it isn't part of the menu data).
          loadPlanFromRows(RI.readPlan(IO.readWorkbook(buf)), f.name);
          continue;
        }
        if (kind === "pool") {
          // A saved pool rebuilds the candidate meals on its own.
          loadPoolFromRows(RI.readPool(IO.readWorkbook(buf)), f.name);
          continue;
        }
        const res = RI.addFile(loader, f.name, buf);
        markFileStatus(kind, !res.error, res.error ? `${f.name} — ${res.error}` : `${f.name} (${res.count})`);
      } catch (err) {
        markFileStatus(kind, false, `${f.name} — ${err.message}`);
      }
    }
    const out = RI.finish(loader);
    // Per-school waste record (rapport): build the lookup even if no menu was
    // dropped this time, so it can fill in waste once a menu exists.
    if (out.rapport && out.rapport.length) {
      app.wasteLookup = global.WasteLookup.build(out.rapport, app.pool || []);
      app.targetSchool = "";
      renderSchoolSelect();
    }
    if (!out.records.length) {
      if (app.wasteLookup && app.wasteLookup.count) {
        status.textContent = `✓ svinn-per-skola: ${app.wasteLookup.count} mätta måltider från ${app.wasteLookup.schoolNames.length} skolor laddade – lägg till en matsedel för att fylla i.`;
        status.style.color = "var(--primary)";
      } else {
        status.textContent = "Inga användbara menyrader. Kontrollera att menyn var med.";
        status.style.color = "var(--danger)";
      }
      app.dayRecords = null; app.pool = null;
      checkReady();
      return;
    }
    app.dayRecords = out.records;
    app.pool = D.buildMealPool(out.records);
    applyEstimates();
    const who = loader.jsonFacility ? ` · ${loader.jsonFacility}` : "";
    let line = `✓ ${out.total} dagar${who} · ${app.pool.length} unika måltider`;
    if (out.stats.min != null) {
      line += ` · svinn ${out.stats.min}–${out.stats.max} g/ätande (medelvärde ${out.stats.mean})`;
    }
    if (app.wasteFilled) line += ` · ${app.wasteFilled} fyllde ur svinn från rapport`;
    status.textContent = line;
    status.style.color = "var(--primary)";
    $("drop-real").classList.add("ok");
    checkReady();
  }

  function setupRealDropzone() {
    const drop = $("drop-real"), input = $("file-real");
    drop.addEventListener("click", () => input.click());
    ["dragenter", "dragover"].forEach((ev) =>
      drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("active"); })
    );
    ["dragleave", "drop"].forEach((ev) =>
      drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("active"); })
    );
    drop.addEventListener("drop", (e) => {
      const files = Array.from(e.dataTransfer.files || []);
      if (files.length) handleRealFiles(files);
    });
    input.addEventListener("change", () => {
      const files = Array.from(input.files || []);
      if (files.length) handleRealFiles(files);
    });
  }


  // ---------- Import a skolmat.info JSON file ----------
  async function handleSkolmatFile(file) {
    const status = $("status-real");
    try {
      const text = await file.text();
      const out = RI.parseSkolmatJSON(text);
      if (!out.records.length) {
        status.textContent = "⚠ " + (out.error || "Ingen matsedel hittades i filen.");
        status.style.color = "var(--danger)";
        return;
      }
      app.dayRecords = out.records;
      app.pool = D.buildMealPool(out.records);
      applyEstimates();
      const who = out.facilityName ? ` · ${out.facilityName}` : "";
      status.textContent = `✓ skolmat.info: ${out.records.length} dagar${who} · ${app.pool.length} unika måltider`;
      status.style.color = "var(--primary)";
      $("drop-real").classList.add("ok");
      markFileStatus("menu", true, `skolmat.info (${out.records.length} dagar)`);
      checkReady();
    } catch (err) {
      status.textContent = "⚠ Kunde inte läsa filen: " + err.message;
      status.style.color = "var(--danger)";
    }
  }

  function setupSkolmatImport() {
    const btn = $("btn-skolmat"), input = $("file-skolmat");
    if (!btn || !input) return;
    btn.addEventListener("click", () => input.click());
    input.addEventListener("change", () => {
      const f = input.files && input.files[0];
      if (f) handleSkolmatFile(f);
      input.value = "";
    });
  }

  function setupPlanImport() {
    const btn = $("btn-plan"), input = $("file-plan");
    if (!btn || !input) return;
    btn.addEventListener("click", () => input.click());
    input.addEventListener("change", async () => {
      const f = input.files && input.files[0];
      input.value = "";
      if (!f) return;
      try {
        const buf = await IO.readArrayBuffer(f);
        loadPlanFromRows(RI.readPlan(IO.readWorkbook(buf)), f.name);
      } catch (err) {
        const status = $("status-real");
        status.textContent = `Kunde inte läsa planen: ${err.message}`;
        status.style.color = "var(--danger)";
      }
    });
  }

  function setupPoolImport() {
    const btn = $("btn-pool"), input = $("file-pool");
    if (!btn || !input) return;
    btn.addEventListener("click", () => input.click());
    input.addEventListener("change", async () => {
      const f = input.files && input.files[0];
      input.value = "";
      if (!f) return;
      try {
        const buf = await IO.readArrayBuffer(f);
        loadPoolFromRows(RI.readPool(IO.readWorkbook(buf)), f.name);
      } catch (err) {
        const status = $("status-real");
        status.textContent = `Kunde inte läsa poolen: ${err.message}`;
        status.style.color = "var(--danger)";
      }
    });
  }

  // Attach waste to meals that lack a measured value. First fill from the
  // per-school lookup (real measured waste, matched by combination); any meal
  // still missing a value gets the component-combination estimate. Meals filled
  // by a previous lookup are cleared first so a school change re-fills them.
  function applyEstimates() {
    if (!app.pool) return;
    if (app.wasteLookup && app.wasteLookup.count) {
      // Clear meals a previous lookup filled so a school change re-fills them.
      for (const id of (app.wasteFilledIds || [])) {
        const m = app.pool.find((x) => x.id === id);
        if (m) m.wasteGPerEater = null;
      }
      app.wasteFilledIds = global.WasteLookup.fill(app.pool, app.wasteLookup, app.targetSchool);
      app.wasteFilled = app.wasteFilledIds.length;
    }
    app.baseline = D.applyWasteEstimates(app.pool);
  }

  // Build the "prefer this school's waste" selector from the loaded rapport.
  function renderSchoolSelect() {
    const sel = $("school-select");
    if (!sel) return;
    const names = (app.wasteLookup && app.wasteLookup.schoolNames) || [];
    sel.innerHTML = `<option value="">Alla skolor (genomsnitt)</option>` +
      names.map((n) => `<option value="${n.replace(/"/g, "&quot;")}"${n === app.targetSchool ? " selected" : ""}>${n}</option>`).join("");
    sel.classList.toggle("hidden", !names.length);
  }

  function checkReady() {
    const ready = app.dayRecords && app.pool && app.pool.length;
    $("btn-generate").disabled = !ready;
  }

  // ---------- Page navigation (setup <-> plan) ----------
  function showPlanPage() {
    $("page-setup").classList.add("hidden");
    $("page-plan").classList.remove("hidden");
    $("btn-back").classList.remove("hidden");
    $("btn-download-plan").classList.remove("hidden");
    const poolBtn = $("btn-download-pool");
    if (poolBtn) poolBtn.classList.remove("hidden");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }
  function showSetupPage() {
    $("page-plan").classList.add("hidden");
    $("page-setup").classList.remove("hidden");
    $("btn-back").classList.add("hidden");
    $("btn-download-plan").classList.add("hidden");
    const poolBtn = $("btn-download-pool");
    if (poolBtn) poolBtn.classList.add("hidden");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }
  $("btn-back").addEventListener("click", showSetupPage);

  // ---------- Generate ----------
  function setProgress(done, total) {
    const bar = $("gen-progress-bar");
    const label = $("gen-progress-label");
    if (!bar) return;
    const pct = total > 0 ? Math.round((done / total) * 100) : 0;
    bar.style.width = pct + "%";
    if (label) label.textContent = `Genererar plan… ${pct}%`;
  }
  function showProgress(on) {
    const el = $("gen-progress");
    if (!el) return;
    el.classList.toggle("hidden", !on);
    if (on) setProgress(0, 1);
  }

  $("btn-generate").addEventListener("click", async () => {
    const btn = $("btn-generate");
    btn.disabled = true;
    app.criteria = readCriteria();
    app.mined = M.mine(app.dayRecords);
    const startISO = $("c-start").value || null;
    showProgress(true);
    app.plan = await P.generate(app.pool, app.mined, app.criteria, startISO, setProgress);
    showProgress(false);
    app.state = { plan: app.plan, pool: app.pool, uses: new Map(), lastUsed: new Map() };
    rebuildState();
    app.eval = P.evaluatePlan(app.plan, app.criteria);
    renderWeekTabs();
    renderPlan();
    renderStatus();
    showPlanPage();
    btn.disabled = false;
  });

  // ---------- Week strip ----------
  function renderWeekTabs() {
    const wrap = $("week-tabs");
    if (!wrap) return;
    let html = `<button class='menu-week-button${app.viewWeek === "all" ? " is-active" : ""}' data-week='all'><span class='menu-week-button-number'>Alla</span></button>`;
    for (let w = 0; w < D.WEEKS; w++) {
      const iso = app.plan ? isoWeek(app.plan[w][0].date) : null;
      html += `<button class='menu-week-button${app.viewWeek === w ? " is-active" : ""}' data-week='${w}'><span class='menu-week-button-number'>${iso != null ? iso : w + 1}</span></button>`;
    }
    wrap.innerHTML = html;
    wrap.querySelectorAll(".menu-week-button").forEach((btn) => {
      btn.addEventListener("click", () => {
        const v = btn.dataset.week;
        app.viewWeek = v === "all" ? "all" : parseInt(v, 10);
        renderWeekTabs();
        renderPlan();
      });
    });
  }

  // ---------- Plan (all weeks OR single week, both as day cards) ----------
  function renderPlan() {
    const empty = $("grid-empty"), cards = $("week-cards");
    if (!app.plan) {
      empty.classList.remove("hidden"); cards.classList.add("hidden");
      return;
    }
    empty.classList.add("hidden");
    cards.classList.remove("hidden");
    // "all" renders the whole plan as a 5×8 table (tags only); a number
    // renders just that week (full: name + chips).
    if (app.viewWeek === "all") {
      renderPlanTable();
    } else {
      renderWeekCards([app.viewWeek], false);
    }
  }

  // ---------- Skolmat day cards (one or many weeks) ----------
  // Compact (Alla) cell: chips only + hover tooltip (name + component summary).
  function compactCell(w, d, counts) {
    const esc = (s) => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/"/g, "&quot;");
    const slot = app.plan[w][d];
    const meal = slot && slot.meal;
    if (!meal) return `<td class='plan-table-cell' data-w='${w}' data-d='${d}'><span class='day-empty'>—</span></td>`;
    const cat = categoryOf(meal);
    const starch = starchOf(meal);
    const soup = soupOf(meal);
    const chips = [];
    if (cat) chips.push(chipHTML(cat.icon, cat.label, cat.cls));
    if (starch) chips.push(chipHTML(starch.icon, starch.label, starch.cls));
    if (soup) chips.push(chipHTML(soup.icon, soup.label, soup.cls));
    if (meal.components.theme) chips.push(chipHTML(null, meal.components.theme, "badge-info"));
    const wv = meal.wasteGPerEater != null ? meal.wasteGPerEater : meal.wasteEstimate;
    if (wv != null) {
      const est = meal.wasteGPerEater == null;
      const cls = wv <= 30 ? "low" : wv <= 55 ? "mid" : "high";
      chips.push(`<span class="waste-chip ${cls}${est ? " est" : ""}">${wv} g${est ? "*" : ""}</span>`);
    }
    const comps = cellCompLabels(meal).map(esc).join("|");
    return `<td class='plan-table-cell' draggable='true' data-w='${w}' data-d='${d}' data-name='${esc(meal.name)}' data-comps='${comps}'>` +
      `<div class='plan-table-chips'>${chips.join("")}</div>` +
      `</td>`;
  }

  // Component totals across the whole 8-week plan, shown under the table.
  function componentStatsHTML(counts) {
    const esc = (s) => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/"/g, "&quot;");
    const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);
    if (!entries.length) return "";
    const items = entries.map(([label, n]) =>
      `<button type="button" class="plan-table-stat" data-comp="${esc(label)}"><span class="plan-table-stat-label">${esc(label)}</span>` +
      `<span class="plan-table-stat-count"><b>${n}</b> ggr</span></button>`
    ).join("");
    return `<div class="plan-table-stats">` +
      `<div class="plan-table-stats-title">Komponenter · 8 veckor</div>` +
      `<div class="plan-table-stats-row">${items}</div></div>`;
  }

  // "Alla" overview as a table: 5 weekday rows × 8 week columns, tags only.
  function renderPlanTable() {
    const counts = planCategoryCounts();
    let html = `<table class='plan-table'>` +
      `<thead><tr><th class='plan-table-corner'></th>`;
    for (let w = 0; w < D.WEEKS; w++) {
      const first = app.plan[w][0];
      const iso = first ? isoWeek(first.date) : null;
      html += `<th class='plan-table-week'>V${iso != null ? iso : w + 1}</th>`;
    }
    html += `</tr></thead><tbody>`;
    for (let d = 0; d < D.DAYS.length; d++) {
      html += `<tr><th class='plan-table-day'>${DAY_SV[d]}</th>`;
      for (let w = 0; w < D.WEEKS; w++) html += compactCell(w, d, counts);
      html += `</tr>`;
    }
    html += `</tbody></table>`;
    html += componentStatsHTML(counts);
    const wrap = $("week-cards");
    wrap.innerHTML = html;
    wrap.querySelectorAll(".plan-table-cell").forEach((el) => {
      el.addEventListener("click", (e) => { e.stopPropagation(); openPopover(el, e); });
    });
    // Drag a cell onto another to swap the two meals (Alla view).
    let dragFrom = null;
    wrap.querySelectorAll(".plan-table-cell").forEach((el) => {
      el.addEventListener("dragstart", (e) => {
        dragFrom = { w: +el.dataset.w, d: +el.dataset.d };
        el.classList.add("dragging");
        e.dataTransfer.effectAllowed = "move";
        e.dataTransfer.setData("text/plain", "");
      });
      el.addEventListener("dragend", () => {
        el.classList.remove("dragging");
        wrap.querySelectorAll(".plan-table-cell.drop-target").forEach((c) => c.classList.remove("drop-target"));
      });
      el.addEventListener("dragover", (e) => {
        if (!dragFrom) return;
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
        el.classList.add("drop-target");
      });
      el.addEventListener("dragleave", () => el.classList.remove("drop-target"));
      el.addEventListener("drop", (e) => {
        e.preventDefault();
        el.classList.remove("drop-target");
        if (!dragFrom) return;
        const from = dragFrom;
        dragFrom = null;
        const to = { w: +el.dataset.w, d: +el.dataset.d };
        if (from.w === to.w && from.d === to.d) return;
        swapMeals(from.w, from.d, to.w, to.d);
      });
    });
    // Click a component stat to highlight every cell that contains it.
    wrap.querySelectorAll(".plan-table-stat").forEach((el) => {
      el.addEventListener("click", () => {
        const comp = el.dataset.comp;
        const active = el.classList.contains("active");
        wrap.querySelectorAll(".plan-table-stat").forEach((s) => s.classList.remove("active"));
        wrap.querySelectorAll(".plan-table-cell").forEach((c) => c.classList.remove("hl"));
        if (active) return;
        el.classList.add("active");
        wrap.querySelectorAll(".plan-table-cell").forEach((c) => {
          const comps = (c.dataset.comps || "").split("|");
          if (comps.includes(comp)) c.classList.add("hl");
        });
      });
    });
    // Hover tooltip (styled, so the name + summary stack on separate lines).
    const tip = $("cell-tooltip");
    if (tip) {
      wrap.addEventListener("mouseover", (e) => {
        const cell = e.target.closest(".plan-table-cell");
        if (!cell || !cell.dataset.name) return;
        $("cell-tooltip-name").textContent = cell.dataset.name;
        $("cell-tooltip-summary").classList.add("hidden");
        tip.classList.remove("hidden");
        const r = cell.getBoundingClientRect();
        const tw = tip.offsetWidth, th = tip.offsetHeight;
        let left = r.left + r.width / 2 - tw / 2;
        left = Math.max(8, Math.min(left, window.innerWidth - tw - 8));
        let top = r.top - th - 8;
        if (top < 8) top = r.bottom + 8;
        tip.style.left = left + "px";
        tip.style.top = top + "px";
      });
      wrap.addEventListener("mouseout", (e) => {
        if (!e.relatedTarget || !e.relatedTarget.closest || !e.relatedTarget.closest(".plan-table-cell"))
          tip.classList.add("hidden");
      });
    }
  }

  // compact: the "Alla" overview — show only the tag chips (no meal name);
  // hovering a day reveals the name + a component-count summary.
  function renderWeekCards(weeks, compact) {
    const wrap = $("week-cards");
    const counts = compact ? planCategoryCounts() : null;
    const esc = (s) => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/"/g, "&quot;");
    let html = "";
    for (const w of weeks) {
      // Week header (only when showing multiple weeks)
      if (weeks.length > 1) {
        const first = app.plan[w][0];
        const firstDate = first ? formatDateSV(first.date) : "";
        const iso = first ? isoWeek(first.date) : null;
        html += `<div class='week-header'>` +
          `<h3 class='week-header-title'>Vecka ${iso != null ? iso : w + 1}</h3>` +
          (firstDate ? `<span class='week-header-date'>från ${firstDate}</span>` : "") +
          `</div>`;
      }
      html += `<div class='week-cards'>`;
      for (let d = 0; d < D.DAYS.length; d++) {
      const slot = app.plan[w][d];
      const meal = slot && slot.meal;
      const date = slot ? formatDateSV(slot.date) : "";
      let rows;
      if (meal) {
        const cat = categoryOf(meal);
        const starch = starchOf(meal);
        const soup = soupOf(meal);
        const extra = [];
        if (starch) extra.push(chipHTML(starch.icon, starch.label, starch.cls));
        if (soup) extra.push(chipHTML(soup.icon, soup.label, soup.cls));
        if (meal.components.theme) extra.push(chipHTML(null, meal.components.theme, "badge-info"));
        const wv = meal.wasteGPerEater != null ? meal.wasteGPerEater : meal.wasteEstimate;
        if (wv != null) {
          const est = meal.wasteGPerEater == null;
          const cls = wv <= 30 ? "low" : wv <= 55 ? "mid" : "high";
          const title = est ? "Uppskattat svinn per ätande (från komponenter)" : "Svinn per ätande";
          extra.push(`<span class="waste-chip ${cls}${est ? " est" : ""}" title="${title}">${wv} g${est ? "*" : ""}</span>`);
        }
        // Compact (Alla): chips only, with a hover tooltip of name + summary.
        if (compact) {
          const chips = [];
          if (cat) chips.push(chipHTML(cat.icon, cat.label, cat.cls));
          extra.forEach(c => chips.push(c));
          const summary = mealSummary(meal, counts);
          const tip = esc(meal.name) + (summary ? `<br>${esc(summary)}` : "");
          rows = `<div class='menu-editor-compact-row compact' data-w='${w}' data-d='${d}' title='${tip}'>` +
            `<div class='menu-editor-compact-row-main'>` +
            (chips.length ? `<div class='menu-editor-compact-allergen-list'>${chips.join("")}</div>` : `<span class='day-empty'>—</span>`) +
            `</div>` +
            `<div class='menu-editor-compact-row-actions'><button class='btn btn-ghost'>Byt</button></div>` +
            `</div>`;
        } else {
          // Full (single week): name + chips.
          const rowFor = (name, c2, extraChips) => {
            const chips = [];
            if (c2) chips.push(chipHTML(c2.icon, c2.label, c2.cls));
            (extraChips || []).forEach(c => chips.push(c));
            return `<div class='menu-editor-compact-row' data-w='${w}' data-d='${d}'>` +
              `<div class='menu-editor-compact-row-main'>` +
              `<p class='meal-text'>${name}</p>` +
              (chips.length ? `<div class='menu-editor-compact-allergen-list'>${chips.join("")}</div>` : "") +
              `</div>` +
              `<div class='menu-editor-compact-row-actions'><button class='btn btn-ghost'>Byt</button></div>` +
              `</div>`;
          };
          rows = rowFor(meal.name, cat, extra);
          if (meal.hasVeg && meal.vegName) {
            rows += rowFor(meal.vegName, { label: "Vegetariskt", icon: "leaf", cls: "badge-success" }, extra);
          }
        }
      } else {
        rows = `<div class='day-empty'>—</div>`;
      }
        html += `<div class='menu-editor-compact-day'>` +
          `<div class='menu-editor-compact-day-rail'>` +
          `<h3 class='day-name'>${DAY_SV[d]}</h3>` +
          (date ? `<p class='day-date'>${date}</p>` : "") +
          `</div>` +
          `<div class='menu-editor-compact-day-main'>${rows}</div>` +
          `</div>`;
      }
      html += `</div>`; // close .week-cards
    }
    wrap.innerHTML = html;
    wrap.querySelectorAll(".menu-editor-compact-row").forEach((el) => {
      el.addEventListener("click", (e) => { e.stopPropagation(); openPopover(el, e); });
    });
  }

  // ---------- Popover ----------
  const pop = $("popover");
  function openPopover(td, e) {
    const w = +td.dataset.w, d = +td.dataset.d;
    const cands = P.alternatives(w, d, app.state, app.criteria, app.mined);
    const current = app.plan[w][d].meal;
    const iso = isoWeek(app.plan[w][d].date);
    $("popover-title").textContent = `${DAY_SV[d]} · Vecka ${iso != null ? iso : w + 1}`;
    const list = $("popover-list");
    list.innerHTML = "";
    if (!cands.length) {
      list.innerHTML = "<p class='helper-text' style='padding:0.5rem'>Inga kandidater uppfyller de hårda kraven här.</p>";
    }
    cands.forEach((c) => {
      const row = document.createElement("div");
      const isCur = current && c.meal.id === current.id;
      row.className = "alt-row" + (isCur ? " current" : "");
      const wv = c.meal.wasteGPerEater != null ? c.meal.wasteGPerEater : c.meal.wasteEstimate;
      const waste = wv != null ? ` · ${wv} g${c.meal.wasteGPerEater == null ? "*" : ""}` : "";
      row.innerHTML =
        `<span class='name'>${c.meal.name}</span>` +
        `<span class='score'>${c.score.toFixed(2)}${waste}</span>`;
      if (c.reasons.length) {
        const r = document.createElement("div");
        r.className = "reason";
        r.textContent = c.reasons.slice(0, 2).join(" · ");
        row.appendChild(r);
      }
      row.addEventListener("click", () => {
        app.plan[w][d].meal = c.meal;
        rebuildState();
        app.eval = P.evaluatePlan(app.plan, app.criteria);
        renderPlan();
        renderStatus();
        openPopover(td, e); // refresh popover with new current
      });
      list.appendChild(row);
    });
    pop.classList.remove("hidden");
    const rect = td.getBoundingClientRect();
    const pw = 360, ph = pop.offsetHeight || 300;
    let left = rect.right + 8;
    if (left + pw > window.innerWidth - 8) left = rect.left - pw - 8;
    if (left < 8) left = 8;
    let top = rect.top;
    if (top + ph > window.innerHeight - 8) top = window.innerHeight - ph - 8;
    if (top < 8) top = 8;
    pop.style.left = left + "px";
    pop.style.top = top + "px";
  }
  $("popover-close").addEventListener("click", () => pop.classList.add("hidden"));
  document.addEventListener("click", (e) => {
    if (!pop.contains(e.target) && !e.target.closest(".cell-card") && !e.target.closest(".menu-editor-compact-row"))
      pop.classList.add("hidden");
  });

  // Swap the meals of two slots (drag-and-drop in the Alla view).
  function swapMeals(w1, d1, w2, d2) {
    const a = app.plan[w1][d1], b = app.plan[w2][d2];
    const tmp = a.meal;
    a.meal = b.meal;
    b.meal = tmp;
    rebuildState();
    app.eval = P.evaluatePlan(app.plan, app.criteria);
    renderPlan();
    renderStatus();
  }

  function rebuildState() {
    app.state.uses = new Map();
    app.state.lastUsed = new Map();
    for (let w = 0; w < D.WEEKS; w++)
      for (let d = 0; d < D.DAYS.length; d++) {
        const meal = app.plan[w][d].meal;
        if (!meal) continue;
        const di = P.dayIndexOf(w, d);
        app.state.uses.set(meal.id, (app.state.uses.get(meal.id) || 0) + 1);
        const last = app.state.lastUsed.get(meal.id);
        if (last === undefined || di > last) app.state.lastUsed.set(meal.id, di);
      }
  }

  // ---------- Status ----------
  function renderStatus() {
    const ev = app.eval, c = app.criteria;
    const chips = [];
    const chip = (cls, txt) => `<span class='chip ${cls}'>${txt}</span>`;
    chips.push(chip(ev.filled === ev.totalDays ? "ok" : "warn", `Fyllt ${ev.filled}/${ev.totalDays} dagar`));
    if (c.preferLowWaste && c.maxAvgWaste > 0) {
      const over = ev.avgWaste > c.maxAvgWaste;
      const avg = Math.round(ev.avgWaste * 10) / 10;
      chips.push(chip(over ? "warn" : "ok", `Genomsnittligt svinn: ${avg} g/ätande (max ${c.maxAvgWaste})`));
    }
    chips.push(chip(ev.varietyViolations ? "warn" : "ok", `Variation: ${ev.varietyViolations ? ev.varietyViolations + " över tak" : "inom tak"}`));
    for (const req of c.requiredTags) {
      const met = ev.issues.find((i) => i.startsWith(req.tag));
      chips.push(chip(met ? "warn" : "ok", `${req.tag}: ${met ? met.split(": ")[1] : "uppfyllt"}`));
    }
    for (const r of c.componentRules || []) {
      const issue = ev.issues.find((i) => i.startsWith(r.name + ":"));
      const per = r.period === "14d" ? "14d" : "v";
      const sym = r.kind === "max" ? `≤${r.count}/${per}` : r.kind === "exact" ? `=${r.count}/${per}` : `≥${r.count}/${per}`;
      const day = r.day != null && r.day !== "" && r.day >= 0 ? ` (${DAY_SV[r.day].slice(0, 3)})` : "";
      const label = r.name + " " + sym + day;
      chips.push(chip(issue ? "warn" : "ok", `${issue ? label + " ⚠" : label + " ✓"}`));
    }
    $("status-bar").innerHTML = chips.join(" ");
  }

  // ---------- Download plan ----------
  $("btn-download-plan").addEventListener("click", () => {
    if (!app.plan) { alert("Generera en plan först."); return; }
    const start = app.plan[0][0].date || $("c-start").value || "";
    const firstISO = isoWeek(app.plan[0][0].date);
    const lastISO = isoWeek(app.plan[app.plan.length - 1][0].date);
    const name = `skolmat-plan-v${firstISO ?? 1}-v${lastISO ?? app.plan.length}${start ? `-${start}` : ""}.xlsx`;
    IO.downloadWorkbook(IO.exportPlan(app.plan, app.criteria), name);
  });

  // ---------- Download pool (all meals) ----------
  $("btn-download-pool").addEventListener("click", () => {
    if (!app.pool) { alert("Importera data först."); return; }
    applyEstimates();
    IO.downloadWorkbook(IO.exportPool(app.pool), "skolmat-pool.xlsx");
  });

  // ---------- Init ----------
  setupRealDropzone();
  setupSkolmatImport();
  setupPlanImport();
  setupPoolImport();
  // Prefer a school's waste: re-fill the pool and re-render if a plan exists.
  $("school-select").addEventListener("change", (e) => {
    app.targetSchool = e.target.value;
    if (app.pool) {
      applyEstimates();
      if (app.plan) { app.eval = P.evaluatePlan(app.plan, app.criteria); renderPlan(); renderStatus(); }
    }
  });
  app.compRules = COMP_RULE_DEFAULTS.map((r) => ({ ...r }));
  renderComponentRules();
  $("btn-add-comp-rule").addEventListener("click", () => {
    syncCompRulesFromDOM(); // keep edits before appending
    app.compRules.push({ name: "", kind: "max", count: 1, period: "week" });
    renderComponentRules();
  });
  checkReady();
})(window);
