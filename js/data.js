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
      /lov|k[öo]ket v[äe]ljer|klimatvecka|[öo]nskevecka|kl[äe]mdag|\btom\b|---/.test(s)
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
    // A record may carry explicit components (a pool file the user edited the
    // component columns in). Those override the name-derived ones for the
    // fields they specify, so the user's tag/component edits survive import.
    if (raw.components) {
      const c = raw.components;
      if (c.protein != null) components.protein = c.protein;
      if (c.starch != null) components.starch = c.starch;
      if (c.sauce != null) components.sauce = c.sauce;
      if (c.prep != null) components.prep = c.prep;
      if (c.vegs) {
        const parts = String(c.vegs).split(",").map((s) => s.trim()).filter(Boolean);
        components.extras = Array.from(new Set([...(components.extras || []), ...parts]));
      }
    }
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

  // ---- Combination-based waste baseline ----
  // A meal is a COMBINATION of components (protein + starch + sauce + …), and
  // the interaction matters: "potato + korv" wastes a different amount than
  // "potato + fish". So we estimate a meal's waste from the average waste of
  // meals that share its COMBINATION, not from independent per-ingredient
  // averages. When a specific combination has too few measured meals behind it
  // (COMBO_MIN_N), we back off to a coarser combination (dropping the least
  // specific part) until we find one with enough data; if none does, we use the
  // global mean. We never fall back to single components.
  const COMBO_DIMS = ["protein", "starch", "sauce", "prep", "theme"];
  const COMBO_MIN_N = 3;

  // Fallback global mean (g/eater) used only when the import has no measured
  // waste at all. Derived from the measured Karlskrona data.
  const DEFAULT_GLOBAL_MEAN = 37.3;

  // Vegetable-type extras (normalised keywords from decompose.js) that form the
  // "vegs" value. Berries/dessert/cheese extras are deliberately excluded.
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

  // The value of a baseline dimension for a meal (vegs is derived from extras).
  function dimValue(meal, dim) {
    return dim === "vegs" ? vegsOf(meal) : (meal.components && meal.components[dim]) || null;
  }

  // The combination key for a meal at a given set of dimensions: the present
  // values joined in a fixed order, or null if none of the dims are present.
  function comboKey(meal, dims) {
    const parts = dims.map((d) => dimValue(meal, d)).filter(Boolean);
    return parts.length ? parts.join(" + ") : null;
  }

  // All combination levels for a meal, most specific first: the full set of its
  // present dimensions, then progressively coarser (dropping the least
  // specific part each step). Single-component levels are intentionally NOT
  // included — a component never serves alone.
  function comboLevels(meal) {
    const dims = COMBO_DIMS.filter((d) => dimValue(meal, d));
    const levels = [];
    let cur = dims;
    while (cur.length > 1) {
      levels.push(cur.slice());
      cur = cur.slice(0, cur.length - 1);
    }
    return levels;
  }

  // Build combination-level means from the pool. Returns
  // { global, measured, combos: { [comboKey]: { mean, n } } } where combos is
  // keyed at every combination level (full and coarser) for each measured meal.
  function buildWasteBaseline(pool) {
    const measured = pool.filter((m) => m.wasteGPerEater != null);
    const groups = {};
    for (const m of measured) {
      for (const dims of comboLevels(m)) {
        const key = comboKey(m, dims);
        if (!key) continue;
        (groups[key] = groups[key] || []).push(m.wasteGPerEater);
      }
    }
    const combos = {};
    for (const [k, arr] of Object.entries(groups))
      combos[k] = { mean: arr.reduce((a, b) => a + b, 0) / arr.length, n: arr.length };
    const global = measured.length
      ? measured.reduce((s, m) => s + m.wasteGPerEater, 0) / measured.length
      : DEFAULT_GLOBAL_MEAN;
    return { global, measured: measured.length, combos };
  }

  // Estimate a meal's waste (g/eater) from its combination profile. Walks the
  // combination levels most-specific-first and returns the mean of the first
  // level backed by enough measured meals (COMBO_MIN_N); otherwise the global
  // mean. Returns null if the meal has no recognised components at all (e.g. a
  // holiday/placeholder day) — a non-meal should not be given a fake figure.
  function estimateWaste(meal, baseline) {
    if (!baseline || !baseline.global) return null;
    const present = COMBO_DIMS.filter((d) => dimValue(meal, d));
    if (!present.length) return null; // no components -> not a real meal to estimate
    const levels = comboLevels(meal);
    for (const dims of levels) {
      const key = comboKey(meal, dims);
      const g = key && baseline.combos[key];
      if (g && g.n >= COMBO_MIN_N) return Math.round(g.mean * 10) / 10;
    }
    // Single-component meal (or no combination with enough data): use the global
    // mean so every real candidate meal gets a re-estimable figure.
    return Math.round(baseline.global * 10) / 10;
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
    comboKey,
    comboLevels,
  };
})(window);
