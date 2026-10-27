/* Component decomposer for Swedish meal descriptions.
   Breaks a dish string into parts (theme/protein/starch/sauce/prep/extras)
   so pattern mining works on the ingredients, not the whole string.
   All matching is done on a diacritic-normalized, lowercased copy. */
(function () {
  "use strict";

  // lowercase, strip all diacritics (å→a, ä→a, ö→o, é→e, …), collapse whitespace
  function normalize(s) {
    return String(s || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[̀-ͯ]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  const THEMES = [
    "marocko", "kina", "usa", "spanien", "grekland", "italia", "indien",
    "thailand", "japan", "frankrike", "tyskland", "sverige", "norge", "portugal",
  ];

  // normalized keyword -> component type. Keys are NFD-normalized (diacritics
  // stripped: färs→fars, fläsk→flask, krämig→kramig, vitlökssås→vitlokssas, …).
  const LEXICON = {
    // ---- protein ----
    "kottbullar": "protein", "kottfarsbiff": "protein", "kottfarsrulle": "protein",
    "farsbiff": "protein", "kyckling": "protein", "flask": "protein",
    "flaskkarre": "protein", "flaskfile": "protein", "flaskstek": "protein",
    "flaskpanna": "protein", "flaskkott": "protein", "korv": "protein",
    "korvstubb": "protein", "korvring": "protein", "falukorv": "protein",
    "wienerkorv": "protein", "fisk": "protein", "lax": "protein",
    "laxfile": "protein", "tunfisk": "protein", "skaldjur": "protein",
    "krabba": "protein", "soja": "protein", "sojagratang": "protein",
    "sojafars": "protein", "sojastrimlor": "protein", "balj": "protein",
    "baljfars": "protein", "baljfarsbiff": "protein", "baljfarsruta": "protein",
    "kikartsbiff": "protein", "kikarter": "protein", "linser": "protein",
    "falafel": "protein", "vego": "protein", "ostbiff": "protein",
    "ostgratang": "protein", "ostgraterad": "protein", "ostgratin": "protein",
    "blodpudding": "protein", "kalops": "protein", "gyros": "protein",
    "pullad": "protein", "kott": "protein", "kottgryta": "protein",
    // ---- starch ----
    "potatis": "starch", "potatismos": "starch", "solskensmos": "starch",
    "klyftpotatis": "starch", "ris": "starch", "risotto": "starch",
    "nasi": "starch", "pasta": "starch", "spagetti": "starch",
    "nudlar": "starch", "bulgur": "starch", "makaroner": "starch",
    "brod": "starch", "tortilla": "starch", "tortillabrod": "starch",
    "paella": "starch", "lasagne": "starch", "lasangette": "starch",
    "lasagnette": "starch", "knackebrod": "starch", "potatisbullar": "starch",
    "grynskott": "starch",
    // ---- sauce ----
    "graddsas": "sauce", "dillsas": "sauce", "tomatsas": "sauce",
    "citronsas": "sauce", "remouladsas": "sauce", "brunsas": "sauce",
    "bbq": "sauce", "ketchup": "sauce", "senap": "sauce", "aioli": "sauce",
    "tzatziki": "sauce", "ortssas": "sauce", "curry": "sauce", "curryratt": "sauce",
    "salsa": "sauce", "pesto": "sauce", "bearnaisesas": "sauce",
    "myntayoghurt": "sauce", "vitlokssas": "sauce", "baljfarssas": "sauce",
    "kottfarssas": "sauce", "kycklingfarsas": "sauce", "lingonsas": "sauce",
    // ---- prep ----
    "stekad": "prep", "stekta": "prep", "panerad": "prep", "gratinad": "prep",
    "gratang": "prep", "ugns": "prep", "ugnsbakad": "prep", "ugnsstekt": "prep",
    "bakad": "prep", "gryta": "prep", "soppa": "prep", "soppas": "prep",
    "wok": "prep", "rostad": "prep", "kramig": "prep", "stuvad": "prep",
    "stuvade": "prep", "glaserad": "prep", "glacerad": "prep", "marinerad": "prep",
    "stroganoff": "prep", "schnitzel": "prep", "biff": "prep",
    // ---- extras ----
    "lingon": "extra", "rodbetor": "extra", "rodbets": "extra", "spenat": "extra",
    "gronsaker": "extra", "rotfrukter": "extra", "rotfrukt": "extra", "morot": "extra",
    "morotsbollar": "extra", "coleslaw": "extra", "sallad": "extra", "gurka": "extra",
    "paprika": "extra", "tomat": "extra", "tillbehor": "extra", "ost": "extra",
    "kall": "extra", "palleg": "extra", "svamp": "extra", "kalpudding": "extra",
    "purjolok": "extra", "broccoli": "extra", "artor": "extra", "apple": "extra",
    "banan": "extra", "melon": "extra", "potatissallad": "extra",
  };

  // index keywords by type, longest-first for matching
  const BY_TYPE = {};
  for (const [kw, type] of Object.entries(LEXICON)) {
    (BY_TYPE[type] = BY_TYPE[type] || []).push(kw);
  }
  for (const t of Object.keys(BY_TYPE)) BY_TYPE[t].sort((a, b) => b.length - a.length);

  function extractTheme(name) {
    const m = name.match(/^([a-z]+)\s*:\s*(.+)$/i);
    if (m) {
      const t = normalize(m[1]);
      if (THEMES.includes(t)) return { theme: t, rest: m[2] };
    }
    return { theme: null, rest: name };
  }

  // find first (longest) keyword of `type` present as a whole token in the text
  function matchType(normText, tokens, type) {
    for (const kw of BY_TYPE[type] || []) {
      for (const tok of tokens) {
        if (tok === kw || tok.startsWith(kw) || (kw.length >= 4 && tok.includes(kw))) {
          return kw;
        }
      }
    }
    return null;
  }

  function decomposeDish(rawName) {
    const name = String(rawName || "").trim();
    const { theme, rest } = extractTheme(name);
    const norm = normalize(rest);
    // strip punctuation from each token so "farsbiffar," still matches "farsbiff"
    const tokens = norm.split(/[\s\-/]+/).map((t) => t.replace(/[^\w]/g, "")).filter(Boolean);

    const protein = matchType(norm, tokens, "protein");
    const starch = matchType(norm, tokens, "starch");
    const sauce = matchType(norm, tokens, "sauce");
    const prep = matchType(norm, tokens, "prep");

    // extras: collect all matches, dedupe, first-seen order
    const extras = [];
    for (const kw of BY_TYPE.extra || []) {
      for (const tok of tokens) {
        if (tok === kw || (kw.length >= 4 && tok.includes(kw))) {
          if (!extras.includes(kw)) extras.push(kw);
          break;
        }
      }
    }

    return { theme, protein, starch, sauce, prep, extras };
  }

  window.Decompose = { decomposeDish, extractTheme, normalize, LEXICON, THEMES };
})();
