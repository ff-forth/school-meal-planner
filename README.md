# Skolmat Planner

A standalone web app for planning an **8-week Swedish school meal menu** ("Skolmat"
style). It mines your historical meal data for patterns, then fills an 8-week plan
that respects your criteria — with a pattern-mining suggestion engine.

The app is **fully client-side**: no server, no build step, no account. It runs
straight from a local file or any static host. All data I/O is `.xlsx` (and
`.json` for the Skolmat import), handled in the browser with [SheetJS](https://sheetjs.com/).

The interface is in **Swedish**.

---

## Run it

There is nothing to install or build.

- **From a file:** open `index.html` in a browser (double-click it, or
  `file://…/school-meal-planner/index.html`).
- **From a static host:** copy the folder anywhere that serves static files
  (GitHub Pages, Netlify, a simple `python3 -m http.server`, …) and open
  `index.html`.

The only external dependency is the SheetJS library, loaded from the jsDelivr
CDN, so the page needs internet access for `.xlsx` import/export. Everything
else is local.

> **Note:** opening from `file://` works for the app itself. If your browser
> blocks file access for the CDN script, serve the folder over `http` instead.

---

## What it does

1. **Import your data** (historical meals, waste, portions, and/or a saved plan).
2. **Set criteria** — repetition limits, waste budget, and per-component rules.
3. **Generate an 8-week plan** that follows the criteria, using mined patterns
   from your history to suggest good meals.
4. **Refine and save** — swap/replace individual meals, then export the plan
   and/or the meal pool as `.xlsx`.

---

## Data you need

You already have what you need — the app reads the files you already produce.
There are **no generator or management scripts** required; just drop the files
in.

### File types (auto-detected by filename)

| File (filename contains) | What it provides |
|---|---|
| `…Matsedel…` / `…Inskrivna…` | **Menu** — the meals served, per day (regular + vegetarian). |
| `…Svinn…` (a svinn report) | **Waste** — total waste (kg) per date, summed across kitchens. |
| `…Portioner…` | **Portions** — number of eaters per date. |
| `…svinn-report-…` (clean report) | **Waste report** — per unit+date waste used for the waste lookup. |
| `…skolmat-plan-…` | **Saved plan** — re-import a previously generated 8-week plan. |
| `…skolmat-pool…` | **Pool** — a saved meal pool (all meals + waste + components). |

**Waste per eater** is computed as `total waste (kg) / eaters × 1000` (g per
eater per day), joined on date.

### Skolmat (`.json`) — getting the files

You can also import `.json` files exported from **skolmat.info** — pick several
at once and they are merged per day. The browser can't call the API directly
from the planner (CORS), so a small **bookmarklet** runs *on* the
[skolmat.info](https://www.skolmat.info) page, fetches the chosen weeks, and
saves a `.json` file you then import. The full write-up lives in
[`tools/skolmat-bookmarklet.md`](tools/skolmat-bookmarklet.md) (source in
[`tools/skolmat-bookmarklet.js`](tools/skolmat-bookmarklet.js)); the setup is
below.

**1 · Create the bookmark**

1. Right-click your bookmarks bar (or "Bookmarks") → **Add bookmark** / **New bookmark**.
2. Name: `Skolmat → JSON`
3. In the **URL** field, paste the whole line below as a single line (it's the
   same line in `tools/skolmat-bookmarklet.js` that starts with `javascript:`).

```
javascript:(function(){var m=location.search.match(/facilityId=([A-Za-z0-9-]+)/);var id=m?m[1]:null;if(!id){alert('Öppna en matsedelsida med facilityId i adressen.');return;}function isoW(d){var t=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()));t.setUTCDate(t.getUTCDate()+4-(t.getUTCDay()||7));return{y:t.getUTCFullYear(),w:Math.ceil(((t-new Date(Date.UTC(t.getUTCFullYear(),0,1)))/864e5+1)/7)};}var now=isoW(new Date());var sw=prompt('Startvecka (ISO, t.ex. 38):',now.w);if(sw===null)return;var sy=prompt('År:',now.y);if(sy===null)return;var n=prompt('Antal veckor (standard 8):','8');if(n===null)return;sw=parseInt(sw,10);sy=parseInt(sy,10);n=parseInt(n,10);if(!sw||!sy||!n){alert('Ogiltigt värde.');return;}var base=document.createElement('div');base.style.cssText='position:fixed;top:0;left:0;right:0;z-index:99999;background:#fff;color:#000;font:13px sans-serif;padding:8px;box-shadow:0 2px 6px rgba(0,0,0,.3)';base.textContent='Hämtar '+n+' veckor...';document.body.appendChild(base);var out=[];(async function(){for(var i=0;i<n;i++){var url='/api/public/matsedlar/'+encodeURIComponent(id)+'?year='+sy+'&week='+(sw+i);try{var r=await fetch(url,{headers:{'Accept':'application/json'}});var j=await r.json();if(j&&j.ok)out.push(j);}catch(e){}}base.remove();var nm=(out.length&&out[0].facility&&out[0].facility.name)||'';var slug=nm.replace(/[^a-z0-9]+/gi,'-').replace(/^-|-$/g,'').toLowerCase();var blob=new Blob([JSON.stringify(out,null,2)],{type:'application/json'});var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='skolmat-'+sy+'-v'+sw+'-v'+(sw+n-1)+(slug?'-'+slug:'')+'.json';document.body.appendChild(a);a.click();a.remove();setTimeout(function(){URL.revokeObjectURL(a.href);},1000);alert('Klart! Hamtade '+out.length+' av '+n+' veckor. Spara filen och importera den i planeraren.');})();})();
```

**2 · Fetch the meals**

1. Open the menu on [skolmat.info](https://www.skolmat.info) — the address must
   contain `facilityId=...` (e.g.
   `https://www.skolmat.info/matsedlar?countyCode=10&municipalityId=1081&facilityId=1C59B6EA-...`).
2. Click the **Skolmat → JSON** bookmark.
3. Answer the prompts:
   - **Startvecka (ISO)** — the week the plan should start in (default = current week).
   - **År** — e.g. `2026`.
   - **Antal veckor** — default `8`, change freely.
4. A file `skolmat-<year>-v<start>-v<end>-<school>.json` downloads automatically
   (the school name is taken from the page, e.g.
   `skolmat-2026-v47-v54-Aspöskolan.json`).

**3 · Import into the planner**

In **step 1 (Data)** click **📥 Importera matsedel (JSON)** and choose the file.
The planner builds the meal pool from the current meals (regular dish +
vegetarian alternative per day) and you can generate the 8-week plan as usual.

**Details**

- Fetches `/api/public/matsedlar/{facilityId}?year=YYYY&week=NN` per week
  (public API, no login).
- Weeks are **ISO weeks** (same as vecka.nu) — the same numbering used in the
  planner.
- The file holds one object per week:
  `{ ok, facility, week: { year, weekNumber, days: { MONDAY: { isOpen, rows: [{ text, allergens }] } } } }`.
- A missing week (e.g. a future week not yet published) is skipped — the
  bookmarklet reports how many of the weeks were fetched.

---

## The workflow

### 1 · Import data

The **Data** panel (visible on both the setup and plan pages) has three sections:

- **Historisk data** — drop the menu + svinn files here.
- **Plandata** — re-import a saved plan and/or a saved pool.
- **Skolmat (.json)** — import Skolmat `.json` files.

Load the **pool before the plan** so plan meals match by name and pick up
components and waste figures.

### 2 · Set criteria

On the setup page:

- **Min. veckor innan rätt upprepas** — minimum weeks before a meal repeats.
- **Max. gånger en rätt på 8 v** — max times a meal appears in the 8 weeks.
- **Max. genomsnittligt svinn (g / ätande / dag)** — average-waste budget.
- **Föredra svinn från skola** — pick a specific school's waste figures, or use
  the all-school average.
- **Komponentregler** — per-component rules. Each row sets a component and a
  rule: **max / minst / exakt** occurrences over a **vecka / 14 dagar** period,
  optionally on a preferred day of week.
- **Startvecka** — the starting week of the plan.

### 3 · Generate

Click **✨ Fyll i 8-veckorsplan**. The planner fills the 8 weeks, honoring the
criteria and using mined patterns to pick good meals. A progress bar shows the
work; a status bar summarizes the result.

### 4 · Refine & save

On the plan page:

- Navigate weeks with the **week tabs** (or the "Alla" view).
- **Swap / replace** any meal — the "Byt" popover offers ranked candidates from
  the pool that satisfy your criteria.
- **📂 Ladda om pool** — reload the pool without going back to setup.
- **← Ändra inställningar** — go back to the criteria.

Saving:

- **💾 Spara** — a dialog to save the **pool**, the **8-week plan**, or both as
  `.xlsx`.
- **⬇ Ladda ner pool** / **⬇ Ladda ner plan** — direct single-file downloads.

Nothing is saved automatically — the plan and pool live only in memory. When
there is **unsaved work** (a plan or pool that hasn't been downloaded yet), the
browser warns before you close or reload the tab. Once you've saved, the warning
stops; editing the plan again brings it back.

---

## Project structure

```
index.html                  App shell + UI
style.css                   Styling
js/
  data.js                   Data model, meal pool, sample loader
  decompose.js              Meal → component decomposition
  import-real.js            .xlsx/.json readers + filename classifier
  xlsx-io.js                SheetJS import/export (pool & plan workbooks)
  mining.js                 Pattern mining over historical meals
  planner.js                8-week plan generation + evaluation
  waste-lookup.js           Per-school waste figures
  ui.js                     All UI wiring
tools/
  skolmat-bookmarklet.md    How to use the Skolmat bookmarklet
  skolmat-bookmarklet.js    The bookmarklet source
data/                       (gitignored) your working data files
```

