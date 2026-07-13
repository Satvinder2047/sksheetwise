# Excel & HR Learning Hub

Your original multi-page hub design (navbar, hero, stats, feature cards, sidebar +
breadcrumb pattern, progress bar, back-to-top button), now with the Excel Function
Library section powered by a search-and-detail engine instead of ~200 hand-written
pages.

## Structure

```
index.html          Homepage
excel.html           Excel Function Library — search, sidebar browse, full detail view
hr.html               Stub page (build out later)
templates.html        Stub page
tutorials.html         Stub page
about.html             Stub page
contact.html            Stub page
css/style.css        All styling
js/script.js          Site utilities + the formula search/detail engine
data/formulas.json    66 Excel functions: syntax, arguments, examples, HR uses, mistakes
```

## How the Excel Function Library works

`excel.html` has three states, all handled by `js/script.js` without ever loading a
new page:

1. **Intro** — popular formulas + how-to-use blurb (default state)
2. **Results** — a filtered list, shown as you type in the search box
3. **Detail** — the full breakdown for one formula (syntax, arguments table, examples,
   HR uses, common mistakes, related functions), styled like your original `if.html`

Clicking a formula in the sidebar, a related-function pill, or a popular-formula card
all call the same `showFormulaDetail(name)` function. The URL updates with a hash
(`excel.html#VLOOKUP`) so links and the back button work.

## This is now a fully static site

No server, no API key, no build step — just HTML, CSS, and JS fetching a local JSON
file. That means it deploys directly to **GitHub Pages**:

1. Push all these files to your repo (root level — no need for a `public` folder this time)
2. Repo → **Settings** → **Pages** → under "Build and deployment", set **Source** to
   "Deploy from a branch", branch `main`, folder `/ (root)` → **Save**
3. GitHub gives you a URL like `https://yourusername.github.io/repo-name/` within a
   minute or two

No Render, no environment variables, no billing.

## Adding more formulas

Open `data/formulas.json` and add an entry in the same shape as the existing ones
(see the IF or VLOOKUP entries for the richest examples). It's searchable and
appears in the sidebar immediately — no code changes needed.

## Filling in the stub pages

`hr.html`, `templates.html`, `tutorials.html`, `about.html`, and `contact.html` are
placeholders with the same navbar/footer so the site doesn't have broken links. Build
these out the same way — copy the structure from `excel.html`'s header/footer and
swap in your own content.
