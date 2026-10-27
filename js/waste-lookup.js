/* waste-lookup.js — look up a meal's measured waste-per-eater from a per-school
   historical record (the "Hantera livs" rapport), matched by the meal's
   component COMBINATION (wording-independent). Per-school values are preferred;
   when the target school has no data for a combination, the value pooled across
   all schools is used. The lookup walks each meal's combination levels
   most-specific-first (full combo → coarser) so sparse data still resolves. */
(function (global) {
  "use strict";
  const D = global.MealData;

  const round1 = (x) => Math.round(x * 10) / 10;

  // Build the lookup from rapport records [{school, name, wasteGPerEater}] and
  // the current meal pool (meals with .components). Each record's dish is
  // decomposed and indexed at every combination level (like the waste baseline),
  // so a later meal matches on the most specific shared combination. Returns
  // { pooled: {comboKey:{sum,n}}, schools: {schoolSlug:{comboKey:{sum,n}}},
  //   schoolNames: [..], count }.
  function build(rapport, pool) {
    const pooled = {};
    const schools = {};
    const schoolNames = [];
    let count = 0;
    const add = (store, key, v) => {
      const g = (store[key] = store[key] || { sum: 0, n: 0 });
      g.sum += v;
      g.n++;
    };
    for (const r of rapport || []) {
      if (r.wasteGPerEater == null) continue;
      const comps = global.Decompose && global.Decompose.decomposeDish
        ? global.Decompose.decomposeDish(r.name)
        : null;
      if (!comps) continue;
      const meal = { components: comps };
      const levels = D.comboLevels(meal);
      if (!levels.length) continue;
      const slug = D.slugify(r.school);
      if (slug && !schools[slug]) { schools[slug] = {}; schoolNames.push(r.school); }
      for (const dims of levels) {
        const key = D.comboKey(meal, dims);
        if (!key) continue;
        add(pooled, key, r.wasteGPerEater);
        if (slug) add(schools[slug], key, r.wasteGPerEater);
      }
      count++;
    }
    return { pooled, schools, schoolNames, count };
  }

  // Look up a meal's waste. `meal` must have .components. `school` (optional)
  // is the target school name; its per-school values are preferred, falling back
  // to the pooled value. Returns { value, source: "school"|"pooled", n } or null.
  function lookup(meal, lookupData, school) {
    if (!lookupData || !meal || !meal.components) return null;
    const levels = D.comboLevels(meal);
    if (!levels.length) return null;
    const slug = school ? D.slugify(school) : null;
    const schoolStore = slug && lookupData.schools[slug] || null;
    // Pass 1: per-school (most specific combo first).
    if (schoolStore) {
      for (const dims of levels) {
        const key = D.comboKey(meal, dims);
        const g = key && schoolStore[key];
        if (g) return { value: round1(g.sum / g.n), source: "school", n: g.n };
      }
    }
    // Pass 2: pooled across all schools.
    for (const dims of levels) {
      const key = D.comboKey(meal, dims);
      const g = key && lookupData.pooled[key];
      if (g) return { value: round1(g.sum / g.n), source: "pooled", n: g.n };
    }
    return null;
  }

  // Fill in wasteGPerEater on pool meals that lack a measured value, using the
  // lookup (per-school if `school` given, else pooled). Returns the ids of the
  // meals it filled (so a later call can clear them before re-filling). Existing
  // measured values are never overridden.
  function fill(pool, lookupData, school) {
    const filledIds = [];
    for (const m of pool || []) {
      if (m.wasteGPerEater != null) continue;
      const res = lookup(m, lookupData, school);
      if (res) { m.wasteGPerEater = res.value; filledIds.push(m.id); }
    }
    return filledIds;
  }

  global.WasteLookup = { build, lookup, fill };
})(window);
