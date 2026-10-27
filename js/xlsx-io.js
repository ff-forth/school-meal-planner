/* xlsx-io.js — SheetJS helpers: read uploaded files to ArrayBuffer, and
   export the finished 8-week plan. Real-file parsing lives in import-real.js. */
(function (global) {
  "use strict";
  const D = global.MealData;

  // Read an .xlsx File -> ArrayBuffer (for RealImport.addFile).
  function readArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  function downloadWorkbook(wb, filename) {
    XLSX.writeFile(wb, filename);
  }

  // Swedish day names so an exported plan re-imports cleanly (the UI is Swedish).
  const DAY_SV = ["Måndag", "Tisdag", "Onsdag", "Torsdag", "Fredag"];

  // Export the finished 8-week plan.
  // plan: { [weekIdx]: { [dayIdx]: { date, meal } } } where meal is a Meal or null.
  function exportPlan(plan, criteria) {
    const header = ["Vecka", "Dag", "Datum", "Måltid", "Vegetariskt", "Svinn g/ätande"];
    const aoa = [header];
    for (let w = 0; w < D.WEEKS; w++) {
      for (let d = 0; d < D.DAYS.length; d++) {
        const day = plan[w] && plan[w][d];
        if (!day) continue;
        const meal = day.meal;
        aoa.push([
          w + 1, DAY_SV[d], day.date || "",
          meal ? meal.name : "",
          meal && meal.vegName ? meal.vegName : "",
          meal && meal.wasteGPerEater != null ? meal.wasteGPerEater : "",
        ]);
      }
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 6 }, { wch: 12 }, { wch: 12 }, { wch: 44 }, { wch: 44 }, { wch: 14 }];

    // Summary sheet
    const sumAoa = [
      ["Kriterier", "Värde"],
      ["Min. veckor innan upprepning", criteria.minGapWeeks],
      ["Max. gånger per rätt (8 v)", criteria.maxUses],
      ["Föredra lågt svinn", criteria.preferLowWaste ? "ja" : "nej"],
      ["Max. genomsnittligt svinn (g/ätande)", criteria.maxAvgWaste || "(ingen)"],
      ["Uteslutna allergener", criteria.excludedAllergens.join(", ") || "(inga)"],
      ["Önskade teman", criteria.requiredTags.map((r) => `${r.tag}:${r.min}/v`).join(", ") || "(inga)"],
      ["Komponentregler", (criteria.componentRules || []).map((r) => {
        const days = ["Mån", "Tis", "Ons", "Tor", "Fre"];
        const day = r.day != null && r.day >= 0 ? ` (${days[r.day]})` : "";
        return `${r.name} ${r.kind} ${r.count} g/${r.period === "14d" ? "14d" : "v"}${day}`;
      }).join("; ") || "(inga)"],
    ];
    const ws2 = XLSX.utils.aoa_to_sheet(sumAoa);
    ws2["!cols"] = [{ wch: 34 }, { wch: 34 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "8-veckorsplan");
    XLSX.utils.book_append_sheet(wb, ws2, "Sammanfattning");
    return wb;
  }

  // Downloadable templates so the client can build their own data files.
  // Each sheet's header must match what the importer looks for.
  function templates() {
    const wb = XLSX.utils.book_new();

    // Menu — importer needs a header row containing "datum" + "veckodag",
    // plus columns "alternativ 1" and "vegetar".
    const menu = [
      ["Datum", "Veckodag", "Alternativ 1", "Alternativ - Vegetariskt"],
      ["2026-09-14", "Måndag", "Köttbullar med kokt potatis och lingon", "Baljfärsruta med kokt potatis och lingon"],
      ["2026-09-15", "Tisdag", "Stekt fisk med kokt potatis och dillsås", "Vegetarisk schnitzel med kokt potatis och dillsås"],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(menu), "Matsedel");

    // Waste — importer sums "vikt" per "datum" (any sheet with those headers).
    const waste = [
      ["Kund", "Svinntyp", "Datum", "Servering", "Komponent", "Vikt"],
      ["Skola", "Matsvinn", "2026-09-14", "Lunch", "Måltid", 12.5],
      ["Skola", "Matsvinn", "2026-09-15", "Lunch", "Måltid", 9.8],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(waste), "Svinn");

    // Portions — importer sums "antal ätande" (or "mängd") per "datum".
    const portions = [
      ["Leverantör", "Datum", "Servering", "Antal ätande"],
      ["Mataktör", "2026-09-14", "Lunch", 320],
      ["Mataktör", "2026-09-15", "Lunch", 305],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(portions), "Portioner");

    // README with the expected columns.
    const readme = [
      ["Skolmat Planner — datamallar"],
      [""],
      ["MENY (filnamn ska innehålla 'Matsedel' eller 'Inskrivna')"],
      ["  Kolumner: Datum (ÅÅÅÅ-MM-DD) | Veckodag | Alternativ 1 (vanlig rätt) | Alternativ - Vegetariskt"],
      ["  En rad per serveringsdag. Datum används för att slå ihop med svinn & portioner."],
      [""],
      ["SVINN (filnamn ska innehålla 'Svinn')"],
      ["  Kolumner: ... | Datum | ... | Vikt (kg)"],
      ["  Importören summerar Vikt per Datum över alla blad."],
      [""],
      ["PORTIONER (filnamn ska innehålla 'Portioner')"],
      ["  Kolumner: ... | Datum | ... | Antal ätande (eller 'Mängd')"],
      ["  Importören summerar antalet ätande per Datum över alla blad."],
      [""],
      ["Svinn per ätande = totalt svinn (kg) / antal ätande × 1000 (g)."],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(readme), "README");

    return wb;
  }

  // Export the meal pool as an editable component sheet.
  // The customer reviews/corrects the component breakdown (protein, carbs,
  // vegs, sauce) and can add measured waste. "Svinn uppskattat" is the
  // component-based estimate the app uses when no measured value exists.
  function exportComponents(pool) {
    const D2 = global.MealData;
    const header = [
      "Måltid", "Protein", "Kolhydrat", "Svinn (mätt)",
      "Svinn uppskattat", "Grönäsar", "Sås", "Tillagning",
    ];
    const aoa = [header];
    const sorted = (pool || []).slice().sort((a, b) => a.name.localeCompare(b.name, "sv"));
    for (const m of sorted) {
      const c = m.components || {};
      aoa.push([
        m.name,
        c.protein || "",
        c.starch || "",
        m.wasteGPerEater != null ? m.wasteGPerEater : "",
        m.wasteEstimate != null ? m.wasteEstimate : "",
        D2.vegsOf(m) || "",
        c.sauce || "",
        c.prep || "",
      ]);
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [
      { wch: 46 }, { wch: 16 }, { wch: 14 }, { wch: 14 },
      { wch: 18 }, { wch: 22 }, { wch: 14 }, { wch: 14 },
    ];

    const readme = [
      ["Komponenter — redigerbar"],
      [""],
      ["Denna fil listar varje unika måltid med dess upplåtna komponenter."],
      ["Kontrollera och rätta kolumnerna Protein / Kolhydrat / Grönäsar / Sås."],
      ["Kolumnen 'Svinn (mätt)' kan fyllas i med svinn per ätande (g) om du har det."],
      ["'Svinn uppskattat' är appens uppskattning baserad på komponenterna."],
      [""],
      ["Komponenterna styr både svinnuppskattningen och mönsterminingen i appen."],
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Komponenter");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(readme), "README");
    return wb;
  }

  // Export the full candidate pool (every unique meal) as a resumable pool file.
  // Columns mirror the processed menu so the app can re-import it: one row per
  // meal with its components and waste. "Svinn/ätande (estimerat)" is the
  // component-based estimate the app uses when no measured value exists.
  function exportPool(pool) {
    const D2 = global.MealData;
    const header = [
      "Måltid", "Vegetariskt", "Protein", "Kolhydrat", "Grönäsar",
      "Sås", "Tillagning", "Svinn/ätande (mätt)", "Svinn/ätande (estimerat)",
    ];
    const aoa = [header];
    const sorted = (pool || []).slice().sort((a, b) => a.name.localeCompare(b.name, "sv"));
    for (const m of sorted) {
      const c = m.components || {};
      aoa.push([
        m.name,
        m.vegName || "",
        c.protein || "",
        c.starch || "",
        D2.vegsOf(m) || "",
        c.sauce || "",
        c.prep || "",
        m.wasteGPerEater != null ? m.wasteGPerEater : "",
        m.wasteEstimate != null ? m.wasteEstimate : "",
      ]);
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [
      { wch: 46 }, { wch: 46 }, { wch: 16 }, { wch: 14 }, { wch: 22 },
      { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 20 },
    ];

    const readme = [
      ["Skolmat plan — pool (alla måltider)"],
      [""],
      ["En rad per unik måltid med komponenter (Protein/Kolhydrat/Grönäsar/Sås/Tillagning)"],
      ["och svinn per ätande (g). 'Svinn/ätande (estimerat)' fylls i av appen när inget"],
      ["mätt värde finns. Importera denna fil (…skolmat-pool…) för att återställa poolen"],
      ["utan att ladda hela menyn igen — t.ex. för att återuppta en sparad plan."],
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pool");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(readme), "README");
    return wb;
  }

  global.XlsxIO = {
    readArrayBuffer,
    downloadWorkbook,
    exportPlan,
    exportPool,
    templates,
    exportComponents,
  };
})(window);
