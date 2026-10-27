/* mining.js — pattern mining over history, working on meal COMPONENTS
   (theme / protein / starch / sauce / prep / extras) rather than whole dish
   names. This gives much better coverage: "potatis" is recognised whether it
   appears as "kokt potatis", "potatismos" or "klyftpotatis".

   Input: dayRecords — an array of day objects, each with a `.components`
   object (from Decompose) and a `.date`. The array is sorted by date so that
   day-to-day sequences are meaningful. */
(function (global) {
  "use strict";

  const COMPONENT_TYPES = ["theme", "protein", "starch", "sauce", "prep"];

  // Collect the component "tokens" for one day: "type:value" strings, plus
  // each extra as "extra:value".
  function dayTokens(components) {
    const toks = [];
    if (!components) return toks;
    for (const t of COMPONENT_TYPES) {
      if (components[t]) toks.push(t + ":" + components[t]);
    }
    for (const e of components.extras || []) toks.push("extra:" + e);
    return toks;
  }

  // ---- Frequency: how often each component value appears ----
  // Returns { byType: { protein: Map(value->count), … }, total: n }
  function componentFrequency(dayRecords) {
    const byType = {};
    for (const t of COMPONENT_TYPES) byType[t] = new Map();
    const extra = new Map();
    let total = 0;
    for (const d of dayRecords) {
      const c = d.components;
      if (!c) continue;
      total++;
      for (const t of COMPONENT_TYPES) {
        if (c[t]) byType[t].set(c[t], (byType[t].get(c[t]) || 0) + 1);
      }
      for (const e of c.extras || []) extra.set(e, (extra.get(e) || 0) + 1);
    }
    byType.extra = extra;
    return { byType, total };
  }

  // ---- Co-occurrence: which component values appear together in a day ----
  // For a "seed" token (e.g. "protein:kottbullar"), list the other tokens that
  // co-occur with it, with confidence = P(other | seed).
  function componentCooc(dayRecords) {
    const n = Math.max(1, dayRecords.length);
    const tokenDays = new Map(); // token -> days containing it
    const pairCount = new Map(); // "a||b" (sorted) -> count
    for (const d of dayRecords) {
      const toks = Array.from(new Set(dayTokens(d.components)));
      for (const t of toks) tokenDays.set(t, (tokenDays.get(t) || 0) + 1);
      for (let i = 0; i < toks.length; i++) {
        for (let j = i + 1; j < toks.length; j++) {
          const a = toks[i], b = toks[j];
          const key = a < b ? a + "||" + b : b + "||" + a;
          pairCount.set(key, (pairCount.get(key) || 0) + 1);
        }
      }
    }
    function coocWith(seed) {
      const cs = tokenDays.get(seed) || 0;
      if (!cs) return [];
      const out = [];
      for (const [other, days] of tokenDays) {
        if (other === seed) continue;
        const key = seed < other ? seed + "||" + other : other + "||" + seed;
        const both = pairCount.get(key) || 0;
        if (!both) continue;
        out.push({ token: other, confidence: both / cs, support: both / n });
      }
      out.sort((a, b) => b.confidence - a.confidence || b.support - a.support);
      return out;
    }
    return { coocWith, tokenDays, n };
  }

  // ---- Sequences: for a component type, which value tends to follow which ----
  // Over consecutive days, track transitions of a single component type.
  function sequencesByComponent(dayRecords) {
    const trans = new Map(); // "type|from|to" -> count
    const fromCount = new Map(); // "type|from" -> count
    for (let i = 0; i < dayRecords.length - 1; i++) {
      const a = dayRecords[i].components;
      const b = dayRecords[i + 1].components;
      if (!a || !b) continue;
      for (const t of COMPONENT_TYPES) {
        const fa = a[t], fb = b[t];
        if (!fa || !fb || fa === fb) continue;
        const fk = t + "|" + fa;
        fromCount.set(fk, (fromCount.get(fk) || 0) + 1);
        const tk = t + "|" + fa + "|" + fb;
        trans.set(tk, (trans.get(tk) || 0) + 1);
      }
    }
    function nextAfter(type, prevValue) {
      const fk = type + "|" + prevValue;
      const total = fromCount.get(fk) || 0;
      if (!total) return [];
      const out = [];
      for (const [tk, c] of trans) {
        if (tk.startsWith(fk + "|")) {
          const to = tk.split("|")[2];
          out.push({ value: to, prob: c / total });
        }
      }
      out.sort((a, b) => b.prob - a.prob);
      return out;
    }
    return { nextAfter };
  }

  // ---- Combined miner ----
  function mine(dayRecords) {
    const sorted = Array.from(dayRecords).sort((a, b) =>
      (a.date || "") < (b.date || "") ? -1 : (a.date || "") > (b.date || "") ? 1 : 0
    );
    return {
      dayRecords: sorted,
      freq: componentFrequency(sorted),
      cooc: componentCooc(sorted),
      seq: sequencesByComponent(sorted),
      dayCount: sorted.length,
    };
  }

  global.Mining = {
    COMPONENT_TYPES,
    dayTokens,
    componentFrequency,
    componentCooc,
    sequencesByComponent,
    mine,
  };
})(window);
