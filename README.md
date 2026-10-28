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

### Skolmat (`.json`)

You can also import `.json` files exported from **skolmat.info** — pick several
at once and they are merged per day. The browser can't call the API directly
from the planner (CORS), so a small **bookmarklet** runs *on* the
[skolmat.info](https://www.skolmat.info) page, fetches the chosen weeks, and
saves a `.json` file you then import.

**How to set up the bookmarklet and fetch the files:** see
[`tools/skolmat-bookmarklet.md`](tools/skolmat-bookmarklet.md) (source in
[`tools/skolmat-bookmarklet.js`](tools/skolmat-bookmarklet.js)).

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

