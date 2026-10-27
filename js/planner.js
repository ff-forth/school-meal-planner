/* planner.js — constraint-aware 8-week plan generation + per-slot alternatives.
   Each slot is ONE complete meal (a regular dish plus its vegetarian
   alternative). Scoring works on the meal's components (from Mining). */
(function (global) {
  "use strict";
  const D = global.MealData;

  // Scoring weights (soft)
  const W = {
    cooc: 1.0,   // component co-occurrence affinity with the previous day
    seq: 0.6,    // sequence fit with the previous day's components
    freq: 0.5,   // how often this exact meal appeared in history
    waste: 0.8,  // lower waste-per-eater is better
    required: 1.2, // bonus for satisfying an unmet weekly required tag
    variety: 0.3,  // small bonus for meals not used yet
  };

  // Parse criteria inputs into a structured object.
  function parseCriteria(input) {
    const required = String(input.required || "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean)
      .map((s) => {
        const [tag, min] = s.split(":");
        return { tag: tag.trim().toLowerCase(), min: parseInt(min, 10) || 1 };
      });
    return {
      minGapWeeks: parseInt(input.minGapWeeks, 10) || 0,
      maxUses: parseInt(input.maxUses, 10) || 99,
      preferLowWaste: !!input.preferLowWaste,
      maxAvgWaste: parseFloat(input.maxAvgWaste) || 0, // 0 = no cap
      excludedAllergens: String(input.excludedAllergens || "")
        .split(",")
        .map((s) => s.trim().toLowerCase())
        .filter(Boolean),
      requiredTags: required,
    };
  }

  function dayIndexOf(w, d) { return w * D.DAYS.length + d; }

  // Hard constraints: allergens, variety gap, max uses.
  function hardOK(meal, dayIdx, state, criteria, relaxGap) {
    for (const a of meal.allergens) {
      if (criteria.excludedAllergens.includes(a)) return false;
    }
    const uses = state.uses.get(meal.id) || 0;
    if (uses >= criteria.maxUses) return false;
    if (!relaxGap) {
      const last = state.lastUsed.get(meal.id);
      const gapDays = criteria.minGapWeeks * 7;
      if (last !== undefined && gapDays > 0 && dayIdx - last < gapDays) return false;
    }
    return true;
  }

  // Score a candidate meal for a slot. Returns {score, reasons:[]}.
  function scoreMeal(meal, ctx, criteria, mined) {
    const { prevDay, weekIdx } = ctx;
    let score = 0;
    const reasons = [];

    // Historical frequency of this exact meal
    const sup = mined.freq.byType && mined.dayCount
      ? (meal.timesServed || 0) / mined.dayCount
      : 0;
    score += sup * W.freq;
    if (sup > 0.05) reasons.push(`served ${meal.timesServed}× in history`);

    // Component co-occurrence / sequence with the previous day's meal.
    // We look at the previous day's components and reward a meal whose
    // components are frequent, and penalise repeating the same protein/starch.
    if (prevDay && prevDay.meal) {
      const pc = prevDay.meal.components;
      const c = meal.components;
      // sequence fit: for each component type, does this value tend to follow
      // the previous day's value?
      for (const t of global.Mining.COMPONENT_TYPES) {
        if (!c[t] || !pc[t]) continue;
        const nexts = mined.seq.nextAfter(t, pc[t]);
        const hit = nexts.find((r) => r.value === c[t]);
        if (hit) {
          score += hit.prob * W.seq;
          if (hit.prob >= 0.4) reasons.push(`follows yesterday's ${t}`);
        }
      }
      // variety within components: penalise repeating yesterday's protein/starch
      if (c.protein && c.protein === pc.protein) {
        score -= 0.5;
        reasons.push(`repeats yesterday's ${c.protein}`);
      }
      if (c.starch && c.starch === pc.starch) {
        score -= 0.3;
      }
    }

    // Waste: lower is better (only when the meal has a waste figure).
    if (criteria.preferLowWaste && meal.wasteGPerEater != null) {
      // normalise: 0 at the fleet mean (~36), 1 at very low (~15)
      const w = meal.wasteGPerEater;
      const wasteScore = Math.max(0, Math.min(1, (40 - w) / 25));
      score += wasteScore * W.waste;
      if (w < 30) reasons.push(`low waste (${w} g/eater)`);
    }

    // Required weekly tags: bonus if this week still needs the tag and meal has it
    for (const req of criteria.requiredTags) {
      if (!meal.tags.includes(req.tag)) continue;
      if (weekTagCount(ctx, req.tag) < req.min) {
        score += W.required;
        reasons.push(`satisfies "${req.tag}" this week`);
      }
    }

    // Variety: small bonus if never used in this plan yet
    if ((ctx.state.uses.get(meal.id) || 0) === 0) score += W.variety;

    return { score, reasons };
  }

  // Count how many meals tagged `tag` are already placed in the current week.
  function weekTagCount(ctx, tag) {
    let count = 0;
    for (let d = 0; d < D.DAYS.length; d++) {
      const meal = ctx.plan[ctx.weekIdx][d].meal;
      if (meal && meal.tags.includes(tag)) count++;
    }
    return count;
  }

  // Get scored, hard-filtered candidates for a slot (excluding an optional meal).
  function getCandidates(w, d, state, criteria, mined, excludeId, relaxGap) {
    const dayIdx = dayIndexOf(w, d);
    const prevDayIdx = dayIdx - 1;
    const prevDay = prevDayIdx >= 0
      ? state.plan[Math.floor(prevDayIdx / D.DAYS.length)][prevDayIdx % D.DAYS.length]
      : null;
    const ctx = { prevDay, weekIdx: w, plan: state.plan, state };

    const out = [];
    for (const meal of state.pool) {
      if (excludeId && meal.id === excludeId) continue;
      if (!hardOK(meal, dayIdx, state, criteria, relaxGap)) continue;
      const { score, reasons } = scoreMeal(meal, ctx, criteria, mined);
      out.push({ meal, score, reasons });
    }
    out.sort((a, b) => b.score - a.score);
    return out;
  }

  // Generate the full plan (greedy + repair + polish). Returns plan[w][d] = meal.
  function generate(pool, mined, criteria, startISO) {
    const plan = [];
    for (let w = 0; w < D.WEEKS; w++) {
      plan.push([]);
      for (let d = 0; d < D.DAYS.length; d++) {
        plan[w].push({ date: dateFor(w, d, startISO) });
      }
    }
    const state = { plan, pool, uses: new Map(), lastUsed: new Map() };

    // Greedy pass
    for (let w = 0; w < D.WEEKS; w++) {
      for (let d = 0; d < D.DAYS.length; d++) {
        const cands = getCandidates(w, d, state, criteria, mined);
        if (cands.length) place(state, w, d, cands[0].meal);
      }
    }

    // Repair pass: fill any empty slot by relaxing the variety gap.
    for (let w = 0; w < D.WEEKS; w++) {
      for (let d = 0; d < D.DAYS.length; d++) {
        if (plan[w][d].meal) continue;
        const relaxed = relaxedCandidate(w, d, state, criteria, mined);
        if (relaxed) place(state, w, d, relaxed);
      }
    }

    // Polish: local search minimising the combined penalty.
    polish(state, criteria, mined);
    return plan;
  }

  // Total penalty for the current plan. Lower is better; 0 = fully compliant.
  function penalty(state, criteria) {
    const gapDays = criteria.minGapWeeks * 7;
    let pen = 0;
    const weekTag = {};
    for (const req of criteria.requiredTags) weekTag[req.tag] = new Array(D.WEEKS).fill(0);

    let wasteSum = 0, wasteN = 0;
    for (let w = 0; w < D.WEEKS; w++) {
      for (let d = 0; d < D.DAYS.length; d++) {
        const meal = state.plan[w][d].meal;
        if (!meal) { pen += 5; continue; }
        if (meal.wasteGPerEater != null) { wasteSum += meal.wasteGPerEater; wasteN++; }
        for (const req of criteria.requiredTags) {
          if (meal.tags.includes(req.tag)) weekTag[req.tag][w]++;
        }
      }
    }
    for (const req of criteria.requiredTags) {
      for (let w = 0; w < D.WEEKS; w++) {
        if (weekTag[req.tag][w] < req.min) pen += (req.min - weekTag[req.tag][w]) * 4;
      }
    }
    if (criteria.maxAvgWaste > 0 && wasteN > 0) {
      const avg = wasteSum / wasteN;
      if (avg > criteria.maxAvgWaste) pen += (avg - criteria.maxAvgWaste) * 2;
    }
    // variety: uses over cap + gap violations
    const uses = new Map();
    const byMeal = new Map();
    for (let w = 0; w < D.WEEKS; w++)
      for (let d = 0; d < D.DAYS.length; d++) {
        const meal = state.plan[w][d].meal;
        if (!meal) continue;
        uses.set(meal.id, (uses.get(meal.id) || 0) + 1);
        if (!byMeal.has(meal.id)) byMeal.set(meal.id, []);
        byMeal.get(meal.id).push(dayIndexOf(w, d));
      }
    for (const [id, n] of uses) if (n > criteria.maxUses) pen += (n - criteria.maxUses) * 2;
    if (gapDays > 0) {
      for (const days of byMeal.values()) {
        days.sort((a, b) => a - b);
        for (let i = 1; i < days.length; i++) if (days[i] - days[i - 1] < gapDays) pen += 1.5;
      }
    }
    return pen;
  }

  function polish(state, criteria, mined) {
    let cur = penalty(state, criteria);
    const maxIter = 3000;
    for (let iter = 0; iter < maxIter && cur > 0; iter++) {
      const w = Math.floor(Math.random() * D.WEEKS);
      const d = Math.floor(Math.random() * D.DAYS.length);
      const curMeal = state.plan[w][d].meal;
      if (!curMeal) continue;
      const cands = getCandidates(w, d, state, criteria, mined, curMeal.id, true);
      if (!cands.length) continue;
      let best = null, bestPen = cur;
      for (const c of cands.slice(0, 15)) {
        if (c.meal.id === curMeal.id) continue;
        swap(state, w, d, c.meal);
        const p = penalty(state, criteria);
        if (p < bestPen) { bestPen = p; best = c.meal; }
        swap(state, w, d, curMeal); // undo
      }
      if (best) { swap(state, w, d, best); cur = bestPen; }
    }
  }

  // Swap a slot to `meal`, updating uses/lastUsed bookkeeping.
  function swap(state, w, d, meal) {
    const di = dayIndexOf(w, d);
    const prev = state.plan[w][d].meal;
    if (prev) state.uses.set(prev.id, Math.max(0, (state.uses.get(prev.id) || 1) - 1));
    state.plan[w][d].meal = meal;
    state.uses.set(meal.id, (state.uses.get(meal.id) || 0) + 1);
    const last = state.lastUsed.get(meal.id);
    if (last === undefined || di > last) state.lastUsed.set(meal.id, di);
  }

  function relaxedCandidate(w, d, state, criteria, mined) {
    // relax the variety gap, keep allergens + max uses
    const dayIdx = dayIndexOf(w, d);
    const prevDayIdx = dayIdx - 1;
    const prevDay = prevDayIdx >= 0
      ? state.plan[Math.floor(prevDayIdx / D.DAYS.length)][prevDayIdx % D.DAYS.length]
      : null;
    const ctx = { prevDay, weekIdx: w, plan: state.plan, state };
    let best = null, bestScore = -Infinity;
    for (const meal of state.pool) {
      const uses = state.uses.get(meal.id) || 0;
      if (uses >= criteria.maxUses) continue;
      for (const a of meal.allergens) if (criteria.excludedAllergens.includes(a)) continue;
      const { score } = scoreMeal(meal, ctx, criteria, mined);
      if (score > bestScore) { bestScore = score; best = meal; }
    }
    return best;
  }

  function place(state, w, d, meal) {
    state.plan[w][d].meal = meal;
    const dayIdx = dayIndexOf(w, d);
    state.uses.set(meal.id, (state.uses.get(meal.id) || 0) + 1);
    const last = state.lastUsed.get(meal.id);
    if (last === undefined || dayIdx > last) state.lastUsed.set(meal.id, dayIdx);
  }

  // Per-slot ranked alternatives (for the cell popover).
  function alternatives(w, d, state, criteria, mined) {
    return getCandidates(w, d, state, criteria, mined).slice(0, 12);
  }

  // A date label for a slot, offset from the plan's start date (a Monday).
  function dateFor(w, d, startISO) {
    let start;
    if (startISO) {
      const [y, m, day] = startISO.split("-").map(Number);
      start = new Date(y, m - 1, day);
    } else {
      start = new Date(2026, 8, 14); // default: a Monday (14 Sep 2026)
    }
    const dt = new Date(start);
    dt.setDate(start.getDate() + w * 7 + d);
    const iso = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}-${String(dt.getDate()).padStart(2, "0")}`;
    return iso;
  }

  // Evaluate the finished plan against criteria -> status summary.
  function evaluatePlan(plan, criteria) {
    const issues = [];
    let filled = 0, wasteSum = 0, wasteN = 0, varietyViolations = 0;
    const tagCounts = {};
    for (const req of criteria.requiredTags) tagCounts[req.tag] = new Array(D.WEEKS).fill(0);
    const uses = new Map();

    for (let w = 0; w < D.WEEKS; w++) {
      for (let d = 0; d < D.DAYS.length; d++) {
        const meal = plan[w][d].meal;
        if (!meal) continue;
        filled++;
        uses.set(meal.id, (uses.get(meal.id) || 0) + 1);
        if (meal.wasteGPerEater != null) { wasteSum += meal.wasteGPerEater; wasteN++; }
        for (const req of criteria.requiredTags) {
          if (meal.tags.includes(req.tag)) tagCounts[req.tag][w]++;
        }
      }
    }
    for (const req of criteria.requiredTags) {
      const weeksMet = tagCounts[req.tag].filter((c) => c >= req.min).length;
      if (weeksMet < D.WEEKS) issues.push(`${req.tag}: ${weeksMet}/${D.WEEKS} weeks meet ≥${req.min}`);
    }
    for (const [id, n] of uses) if (n > criteria.maxUses) varietyViolations++;
    if (varietyViolations) issues.push(`${varietyViolations} meal(s) exceed the max-uses cap`);

    const avgWaste = wasteN ? wasteSum / wasteN : null;
    if (criteria.maxAvgWaste > 0 && avgWaste != null && avgWaste > criteria.maxAvgWaste) {
      issues.push(`average waste ${avgWaste.toFixed(1)} g/eater exceeds ${criteria.maxAvgWaste}`);
    }

    const totalDays = D.WEEKS * D.DAYS.length;
    return {
      totalDays,
      filled,
      avgWaste,
      varietyViolations,
      issues,
      ok: issues.length === 0 && filled === totalDays,
    };
  }

  global.Planner = {
    parseCriteria,
    generate,
    alternatives,
    evaluatePlan,
    dayIndexOf,
  };
})(window);
