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

  // ---------- Category badge (protein → Fisk/Fågel/Kött/Vegetariskt) ----------
  const ICONS = {
    salad: '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M7 21h10"/><path d="M12 21a9 9 0 0 0 9-9H3a9 9 0 0 0 9 9Z"/><path d="M11.38 12a2.4 2.4 0 0 1-.4-4.77 2.4 2.4 0 0 1 3.2-2.77 2.4 2.4 0 0 1 3.47-.63 2.4 2.4 0 0 1 3.37 3.37 2.4 2.4 0 0 1-1.1 3.7 2.51 2.51 0 0 1 .03 1.1"/><path d="m13 12 4-4"/><path d="M10.9 7.25A3.99 3.99 0 0 0 4 10c0 .73.2 1.41.54 2"/></svg>',
    fish: '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6.5 12c.94-3.46 4.94-6 8.5-6 3.56 0 6.06 2.54 7 6-.94 3.47-3.44 6-7 6s-7.56-2.53-8.5-6Z"/><path d="M18 12v.5"/><path d="M16 17.93a9.77 9.77 0 0 1 0-11.86"/><path d="M7 10.67C7 8 5.58 5.97 2.73 5.5c-1 1.5-1 5 .23 6.5-1.24 1.5-1.24 5-.23 6.5C5.58 18.03 7 16 7 13.33"/></svg>',
    drumstick: '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15.4 15.63a7.875 6 135 1 1 6.23-6.23 4.5 3.43 135 0 0-6.23 6.23"/><path d="m8.29 12.71-2.6 2.6a2.5 2.5 0 1 0-1.65 4.65A2.5 2.5 0 1 0 8.7 18.3l2.59-2.59"/></svg>',
    beef: '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3c5 0 9 3 9 7s-4 7-9 7-9-3-9-7 4-7 9-7Z"/><path d="M9 10c1.2-1.2 3.4-1.2 4.6 0"/><path d="M9 14c1.2-1.2 3.4-1.2 4.6 0"/></svg>',
    pork: '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="9" width="20" height="6" rx="3"/><path d="M8 9v6"/><path d="M16 9v6"/></svg>',
    leaf: '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 20A7 7 0 0 1 9.8 6.1C15.5 5 17 4.48 19 2c1 2 2 4.18 2 8 0 5.5-4.78 10-10 10Z"/><path d="M2 21c0-3 1.85-5.36 5.08-6C9.5 14.52 12 13 13 12"/></svg>',
  };
  function categoryOf(meal) {
    const p = (meal.components && meal.components.protein ? meal.components.protein : "");
    if (/fisk|lax|tunfisk|skaldjur|krabba/.test(p)) return { label: "Fisk", icon: "fish", cls: "badge-info" };
    if (/kyckling|fågel|pullad|gyros/.test(p)) return { label: "Fågel", icon: "drumstick", cls: "badge-neutral" };
    if (/soja|vego|balj|kikarts|linse|falafel|ost/.test(p)) return { label: "Vegetariskt", icon: "leaf", cls: "badge-success" };
    // pork (griskött) vs beef (nötkött) — these are distinct and must be shown separately
    if (/flask|falukorv|korv|wienerkorv|blodpudding/.test(p)) return { label: "Griskött", icon: "pork", cls: "badge-warning" };
    if (/kott|farsbiff|kalops|biff/.test(p)) return { label: "Nötkött", icon: "beef", cls: "badge-neutral" };
    if (p) return { label: "Kött", icon: "beef", cls: "badge-neutral" };
    return null;
  }
  function chipHTML(icon, label, cls) {
    return `<span class="menu-editor-compact-allergen-chip ${cls || ""}">${ICONS[icon] || ""}${label}</span>`;
  }

  // ---------- Criteria ----------
  function readCriteria() {
    return P.parseCriteria({
      minGapWeeks: $("c-gap").value,
      maxUses: $("c-maxuse").value,
      preferLowWaste: $("c-prefer-waste").checked,
      maxAvgWaste: $("c-max-waste").value,
      excludedAllergens: $("c-allergens").value,
      required: $("c-required").value,
      componentRules: $("c-components").value,
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

  async function handleRealFiles(files) {
    const status = $("status-real");
    const loader = RI.createLoader();
    for (const f of files) {
      const kind = RI.classify(f.name);
      try {
        const buf = await IO.readArrayBuffer(f);
        const res = RI.addFile(loader, f.name, buf);
        markFileStatus(kind, !res.error, res.error ? `${f.name} — ${res.error}` : `${f.name} (${res.count})`);
      } catch (err) {
        markFileStatus(kind, false, `${f.name} — ${err.message}`);
      }
    }
    const out = RI.finish(loader);
    if (!out.records.length) {
      status.textContent = "Inga användbara menyrader. Kontrollera att menyn var med.";
      status.style.color = "var(--danger)";
      app.dayRecords = null; app.pool = null;
      checkReady();
      return;
    }
    app.dayRecords = out.records;
    app.pool = D.buildMealPool(out.records);
    applyEstimates();
    status.textContent = `✓ ${out.joined}/${out.total} dagar ihopslagna · ${app.pool.length} unika måltider · ` +
      `svinn ${out.stats.min}–${out.stats.max} g/ätande (medelvärde ${out.stats.mean})`;
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

  function loadSample() {
    app.dayRecords = D.loadSampleReal();
    app.pool = D.buildMealPool(app.dayRecords);
    applyEstimates();
    const status = $("status-real");
    status.textContent = `✓ exempel: ${app.dayRecords.length} dagar · ${app.pool.length} unika måltider`;
    status.style.color = "var(--primary)";
    $("drop-real").classList.add("ok");
    markFileStatus("menu", true, "exempelmatsedel");
    markFileStatus("waste", true, "exempelsvinn");
    markFileStatus("portions", true, "exempelportioner");
    checkReady();
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

  // Attach component-based waste estimates to meals that lack a measured value.
  function applyEstimates() {
    if (!app.pool) return;
    app.baseline = D.applyWasteEstimates(app.pool);
  }

  function checkReady() {
    const ready = app.dayRecords && app.pool && app.pool.length;
    $("btn-generate").disabled = !ready;
    const compBtn = $("btn-components");
    if (compBtn) compBtn.disabled = !ready;
  }

  // ---------- Page navigation (setup <-> plan) ----------
  function showPlanPage() {
    $("page-setup").classList.add("hidden");
    $("page-plan").classList.remove("hidden");
    $("btn-back").classList.remove("hidden");
    $("btn-download-plan").classList.remove("hidden");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }
  function showSetupPage() {
    $("page-plan").classList.add("hidden");
    $("page-setup").classList.remove("hidden");
    $("btn-back").classList.add("hidden");
    $("btn-download-plan").classList.add("hidden");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }
  $("btn-back").addEventListener("click", showSetupPage);

  // ---------- Generate ----------
  $("btn-generate").addEventListener("click", () => {
    app.criteria = readCriteria();
    app.mined = M.mine(app.dayRecords);
    const startISO = $("c-start").value || null;
    app.plan = P.generate(app.pool, app.mined, app.criteria, startISO);
    app.state = { plan: app.plan, pool: app.pool, uses: new Map(), lastUsed: new Map() };
    rebuildState();
    app.eval = P.evaluatePlan(app.plan, app.criteria);
    renderWeekTabs();
    renderPlan();
    renderStatus();
    showPlanPage();
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
    // "all" renders every week stacked with a header; a number renders just that week.
    const weeks = app.viewWeek === "all"
      ? Array.from({ length: D.WEEKS }, (_, i) => i)
      : [app.viewWeek];
    renderWeekCards(weeks);
  }

  // ---------- Skolmat day cards (one or many weeks) ----------
  function renderWeekCards(weeks) {
    const wrap = $("week-cards");
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
        // Render the regular and vegetarian meals as two separate rows.
        const rowFor = (name, cat, extraChips) => {
          const chips = [];
          if (cat) chips.push(chipHTML(cat.icon, cat.label, cat.cls));
          (extraChips || []).forEach(c => chips.push(c));
          return `<div class='menu-editor-compact-row' data-w='${w}' data-d='${d}'>` +
            `<div class='menu-editor-compact-row-main'>` +
            `<p class='meal-text'>${name}</p>` +
            (chips.length ? `<div class='menu-editor-compact-allergen-list'>${chips.join("")}</div>` : "") +
            `</div>` +
            `<div class='menu-editor-compact-row-actions'><button class='btn btn-ghost'>Byt</button></div>` +
            `</div>`;
        };
        const cat = categoryOf(meal);
        const extra = [];
        if (meal.components.theme) extra.push(chipHTML(null, meal.components.theme, "badge-info"));
        const wv = meal.wasteGPerEater != null ? meal.wasteGPerEater : meal.wasteEstimate;
        if (wv != null) {
          const est = meal.wasteGPerEater == null;
          const cls = wv <= 30 ? "low" : wv <= 55 ? "mid" : "high";
          const title = est ? "Uppskattat svinn per ätande (från komponenter)" : "Svinn per ätande";
          extra.push(`<span class="waste-chip ${cls}${est ? " est" : ""}" title="${title}">${wv} g${est ? "*" : ""}</span>`);
        }
        rows = rowFor(meal.name, cat, extra);
        if (meal.hasVeg && meal.vegName) {
          rows += rowFor(meal.vegName, { label: "Vegetariskt", icon: "leaf", cls: "badge-success" }, []);
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
      const label = r.name + (r.kind === "max" ? ` ≤${r.count}/${r.period === "14d" ? "14d" : "v"}` : ` ≥${r.count}/v`);
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

  // ---------- Download data templates ----------
  $("btn-templates").addEventListener("click", () => {
    IO.downloadWorkbook(IO.templates(), "skolmat-mallar.xlsx");
  });

  // ---------- Download editable component sheet ----------
  $("btn-components").addEventListener("click", () => {
    if (!app.pool) { alert("Importera data först."); return; }
    applyEstimates();
    IO.downloadWorkbook(IO.exportComponents(app.pool), "skolmat-komponenter.xlsx");
  });

  // ---------- Init ----------
  setupRealDropzone();
  setupSkolmatImport();
  $("btn-sample").addEventListener("click", loadSample);
  checkReady();
})(window);
