/* data.js — Meal data model, normalization, and pool building.
   A "Meal" is one complete day's serving, shown as a regular dish plus its
   vegetarian alternative. For analysis each dish is decomposed into
   components (see decompose.js) so pattern mining works on the ingredients. */
(function (global) {
  "use strict";

  const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
  const WEEKS = 8;

  // ---- slug / id ----
  function slugify(name) {
    return String(name || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[̀-ͯ]/g, "")
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 60);
  }

  // Placeholder menu days that are not real meals (holidays, choice days, …).
  function isPlaceholder(name) {
    const s = String(name || "").toLowerCase();
    return (
      !s.trim() ||
      /lov|k[öo]ket v[äe]ljer|klimatvecka|[öo]nskevecka|kl[äe]mdag|tom|---/.test(s)
    );
  }

  // ---- allergen / tag derivation from components ----
  const ALLERGEN_RULES = [
    // [component type, keyword substrings, allergen]
    ["protein", ["fisk", "lax", "tunfisk", "skaldjur", "krabba"], "fish"],
    ["protein", ["skaldjur", "krabba"], "shellfish"],
    ["starch", ["pasta", "spagetti", "nudlar", "brod", "tortilla", "lasagne", "lasangette", "lasagnette", "knackebrod", "risotto"], "gluten"],
    ["sauce", ["graddsas", "dillsas", "remouladsas", "bearnaisesas", "aioli", "myntayoghurt", "kramig"], "dairy"],
    ["extra", ["ost"], "dairy"],
  ];

  function deriveAllergens(components) {
    const set = new Set();
    for (const [type, kws, allergen] of ALLERGEN_RULES) {
      const val =
        type === "extras"
          ? (components.extras || []).join(" ")
          : String(components[type] || "");
      if (kws.some((k) => val.includes(k))) set.add(allergen);
    }
    return Array.from(set);
  }

  // Build a normalized Meal from a raw day record.
  // raw: { name, vegName, wasteGPerEater, date, weekday }
  function normalizeMeal(raw) {
    const name = String(raw.name || "").trim();
    const vegName = String(raw.vegName || "").trim();
    const components =
      global.Decompose && global.Decompose.decomposeDish
        ? global.Decompose.decomposeDish(name)
        : { theme: null, protein: null, starch: null, sauce: null, prep: null, extras: [] };
    return {
      id: slugify(name),
      name,
      vegName,
      hasVeg: !!vegName,
      components,
      allergens: deriveAllergens(components),
      tags: [components.theme].filter(Boolean),
      wasteGPerEater:
        raw.wasteGPerEater != null && !isNaN(raw.wasteGPerEater)
          ? Number(raw.wasteGPerEater)
          : null,
      timesServed: 0,
      lastSeenDate: raw.date || null,
    };
  }

  // Build the candidate pool from a list of day records.
  // Dedupes by regular-dish name; aggregates usage stats and keeps the most
  // recent date / a veg alternative if one appears.
  function buildMealPool(dayRecords) {
    const map = new Map();
    for (const raw of dayRecords) {
      if (isPlaceholder(raw.name)) continue;
      const m = normalizeMeal(raw);
      const prev = map.get(m.id);
      if (!prev) {
        m.timesServed = 1;
        map.set(m.id, m);
      } else {
        prev.timesServed += 1;
        if (m.vegName && !prev.vegName) {
          prev.vegName = m.vegName;
          prev.hasVeg = true;
        }
        if (m.wasteGPerEater != null && prev.wasteGPerEater == null) {
          prev.wasteGPerEater = m.wasteGPerEater;
        }
        if (raw.date && (!prev.lastSeenDate || raw.date > prev.lastSeenDate)) {
          prev.lastSeenDate = raw.date;
        }
      }
    }
    return Array.from(map.values());
  }

  // Load the bundled offline sample (window.SampleReal) into a pool.
  function loadSampleReal() {
    const recs = (global.SampleReal && global.SampleReal.dayRecords) || [];
    return buildMealPool(recs);
  }

  // ---- Component-based waste baseline ----
  // When a meal has no measured waste (e.g. imported from a menu without svinn),
  // estimate it from the average waste of meals that share each of its
  // components: protein, starch (carbs), sauce, and vegs. Each component value
  // only counts if enough meals back it (BASELINE_MIN_N); otherwise the global
  // mean is used. The estimate is the plain average of the available
  // component-value means, so a meal with no recognised components falls back
  // to the global mean.
  const BASELINE_DIMS = ["protein", "starch", "sauce", "vegs"];
  const BASELINE_MIN_N = 3;

  // Default component waste values (g/eater), derived from measured data.
  // Used to seed the baseline so estimates work even when a meal — or the whole
  // import — has no measured waste. Measured values always override these.
  // Entries are pre-vetted, so they are usable without the BASELINE_MIN_N guard.
  const DEFAULT_BASELINE = {
    global: 37.3,
    dims: {
      protein: {
        kottbullar: 33.2, flaskkarre: 38.9, korv: 39.7, kott: 37.8,
        flaskkott: 46.0, fisk: 38.5, kyckling: 37.7, farsbiff: 36.6,
        wienerkorv: 32.7, falukorv: 36.8,
      },
      starch: {
        potatis: 38.2, brod: 37.5, ris: 38.4, klyftpotatis: 36.6,
        pasta: 37.3, nudlar: 43.6, potatismos: 35.4, tortillabrod: 33.5,
        makaroner: 31.1, lasagne: 32.7,
      },
      sauce: {
        brunsas: 38.1, bbq: 38.5, dillsas: 43.1, graddsas: 34.9,
        aioli: 38.9, ketchup: 33.5, curry: 38.3, bearnaisesas: 38.0,
        vitlokssas: 39.6,
      },
      vegs: {
        rodbetor: 48.2, spenat: 41.2, tillbehor: 33.0, sallad: 39.1,
        artor: 39.1, purjolok: 30.8, gronsaker: 40.1, gurka: 32.7,
        svamp: 38.0,
      },
    },
  };

  // Vegetable-type extras (normalised keywords from decompose.js) that form the
  // "vegs" dimension. Berries/dessert/cheese extras are deliberately excluded.
  const VEG_EXTRA_KW = new Set([
    "rodbetor", "rodbets", "spenat", "gronsaker", "rotfrukter", "rotfrukt",
    "morot", "morotsbollar", "coleslaw", "sallad", "gurka", "paprika",
    "tomat", "broccoli", "artor", "purjolok", "svamp", "tillbehor",
  ]);

  // The "vegs" value for a meal: its vegetable extras joined, or null if none.
  function vegsOf(meal) {
    const ex = (meal.components && meal.components.extras) || [];
    const vegs = ex.filter((e) => VEG_EXTRA_KW.has(e));
    return vegs.length ? vegs.join(",") : null;
  }

  // value of a baseline dimension for a meal
  function dimValue(meal, dim) {
    return dim === "vegs" ? vegsOf(meal) : (meal.components && meal.components[dim]) || null;
  }

  // Build per-dimension means from the pool, seeded with DEFAULT_BASELINE so
  // estimates work even with no measured waste. Measured values override the
  // defaults. Returns { global, measured, dims: { [dim]: { [value]: {mean,n,def} } } }.
  function buildWasteBaseline(pool) {
    const measured = pool.filter((m) => m.wasteGPerEater != null);
    // Start from the defaults (pre-vetted, usable without the min-n guard).
    const dims = {};
    for (const dim of BASELINE_DIMS) {
      dims[dim] = {};
      for (const [k, v] of Object.entries(DEFAULT_BASELINE.dims[dim] || {}))
        dims[dim][k] = { mean: v, n: 0, def: true };
    }
    // Layer measured data on top; a measured value replaces its default.
    for (const dim of BASELINE_DIMS) {
      const groups = {};
      for (const m of measured) {
        const key = dimValue(m, dim);
        if (!key) continue;
        (groups[key] = groups[key] || []).push(m.wasteGPerEater);
      }
      for (const [k, arr] of Object.entries(groups))
        dims[dim][k] = { mean: arr.reduce((a, b) => a + b, 0) / arr.length, n: arr.length, def: false };
    }
    // Global mean: measured mean if available, else the default.
    const global = measured.length
      ? measured.reduce((s, m) => s + m.wasteGPerEater, 0) / measured.length
      : DEFAULT_BASELINE.global;
    return { global, measured: measured.length, dims };
  }

  // Estimate a meal's waste (g/eater) from its component profile.
  // Averages the means of the component values that are usable (measured with
  // enough data, or a pre-vetted default); falls back to the global mean.
  // Returns null if the meal has no recognised components at all (e.g. a
  // holiday/placeholder day) — a non-meal should not be given a fake figure.
  function estimateWaste(meal, baseline) {
    if (!baseline || !baseline.global) return null;
    const vals = [];
    let hasComponent = false;
    for (const dim of BASELINE_DIMS) {
      const key = dimValue(meal, dim);
      if (!key) continue;
      hasComponent = true;
      const g = baseline.dims[dim] && baseline.dims[dim][key];
      if (g && (g.def || g.n >= BASELINE_MIN_N)) vals.push(g.mean);
    }
    if (!hasComponent) return null; // no components -> not a real meal to estimate
    const est = vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : baseline.global;
    return Math.round(est * 10) / 10;
  }

  // Attach a `wasteEstimate` to every meal that lacks a measured value.
  function applyWasteEstimates(pool) {
    const baseline = buildWasteBaseline(pool);
    for (const m of pool) {
      m.wasteEstimate = m.wasteGPerEater != null ? null : estimateWaste(m, baseline);
    }
    return baseline;
  }

  global.MealData = {
    DAYS,
    WEEKS,
    slugify,
    isPlaceholder,
    normalizeMeal,
    buildMealPool,
    loadSampleReal,
    buildWasteBaseline,
    estimateWaste,
    applyWasteEstimates,
    vegsOf,
    BASELINE_DIMS,
  };
})(window);
