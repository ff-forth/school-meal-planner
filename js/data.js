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

  global.MealData = {
    DAYS,
    WEEKS,
    slugify,
    isPlaceholder,
    normalizeMeal,
    buildMealPool,
    loadSampleReal,
  };
})(window);
