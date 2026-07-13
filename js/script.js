// =======================================
// EXCEL & HR LEARNING HUB
// =======================================

console.log("Excel & HR Learning Hub loaded.");

// =======================================
// Footer year
// =======================================
document.addEventListener("DOMContentLoaded", () => {
  const year = document.getElementById("year");
  if (year) year.textContent = new Date().getFullYear();
});

// =======================================
// Reading progress bar
// =======================================
window.addEventListener("scroll", () => {
  const bar = document.getElementById("progressBar");
  if (!bar) return;
  const scrollTop = document.documentElement.scrollTop || document.body.scrollTop;
  const height = document.documentElement.scrollHeight - document.documentElement.clientHeight;
  bar.style.width = (height > 0 ? (scrollTop / height) * 100 : 0) + "%";
});

// =======================================
// Back to top button
// =======================================
window.addEventListener("scroll", () => {
  const btn = document.getElementById("topBtn");
  if (!btn) return;
  btn.style.display = window.scrollY > 400 ? "block" : "none";
});

document.addEventListener("DOMContentLoaded", () => {
  const btn = document.getElementById("topBtn");
  if (btn) {
    btn.addEventListener("click", () => window.scrollTo({ top: 0, behavior: "smooth" }));
  }
});

// =======================================
// Smooth scroll for on-page anchors
// =======================================
document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll('a[href^="#"]:not([href="#"])').forEach((anchor) => {
    anchor.addEventListener("click", function (e) {
      const target = document.querySelector(this.getAttribute("href"));
      if (target) {
        e.preventDefault();
        target.scrollIntoView({ behavior: "smooth" });
      }
    });
  });
});

// =======================================
// Site-wide search box (hero on any page)
// Redirects to excel.html with the query, where
// the Formula Finder engine takes over.
// =======================================
function searchSite() {
  const box = document.getElementById("searchBox");
  if (!box) return;
  const query = box.value.trim();
  if (!query) {
    alert("Please enter an Excel function or topic.");
    return;
  }
  const onExcelPage = !!document.getElementById("formulaApp");
  if (onExcelPage) {
    runFormulaSearch(query);
  } else {
    window.location.href = "excel.html?q=" + encodeURIComponent(query);
  }
}

document.addEventListener("DOMContentLoaded", () => {
  const box = document.getElementById("searchBox");
  if (box) {
    box.addEventListener("keypress", (e) => {
      if (e.key === "Enter") searchSite();
    });
  }
});

// =======================================
// Formula Finder engine
// Runs only on excel.html (guarded by #formulaApp)
// =======================================

let FORMULAS = [];
const POPULAR_FORMULAS = ["IF", "VLOOKUP", "XLOOKUP", "SUMIFS", "INDEX", "MATCH", "COUNTIFS", "IFERROR"];

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

function scoreFormula(entry, query) {
  const q = query.toLowerCase();
  const name = entry.formula.toLowerCase();
  if (name === q) return 100;
  if (name.startsWith(q)) return 80;
  if (name.includes(q)) return 60;
  if (entry.description.toLowerCase().includes(q)) return 40;
  if (entry.category.toLowerCase().includes(q)) return 20;
  if ((entry.relatedFormulas || []).some((r) => r.toLowerCase().includes(q))) return 10;
  return 0;
}

function levelClass(level) {
  return "level-" + (level || "intermediate").toLowerCase();
}

function buildFormulaSidebar() {
  const sidebar = document.getElementById("formulaSidebar");
  if (!sidebar) return;

  const groups = { Beginner: [], Intermediate: [], Advanced: [] };
  FORMULAS.forEach((f) => { if (groups[f.level]) groups[f.level].push(f); });
  const icons = { Beginner: "\u{1F7E2}", Intermediate: "\u{1F535}", Advanced: "\u{1F7E3}" };

  sidebar.innerHTML = `<h4 class="mb-3">📚 Excel Functions</h4>`;
  Object.entries(groups).forEach(([level, items]) => {
    if (!items.length) return;
    const h6 = document.createElement("h6");
    h6.textContent = `${icons[level]} ${level}`;
    sidebar.appendChild(h6);
    const ul = document.createElement("ul");
    items.sort((a, b) => a.formula.localeCompare(b.formula)).forEach((f) => {
      const li = document.createElement("li");
      const a = document.createElement("a");
      a.href = "#";
      a.textContent = f.formula;
      a.dataset.formula = f.formula;
      a.addEventListener("click", (e) => {
        e.preventDefault();
        showFormulaDetail(f.formula);
      });
      li.appendChild(a);
      ul.appendChild(li);
    });
    sidebar.appendChild(ul);
  });
}

function markActiveSidebar(formulaName) {
  document.querySelectorAll("#formulaSidebar a").forEach((a) => {
    a.classList.toggle("active", a.dataset.formula === formulaName);
  });
}

function renderPopularGrid() {
  const grid = document.getElementById("popularGrid");
  if (!grid) return;
  grid.innerHTML = "";
  POPULAR_FORMULAS.forEach((name) => {
    const entry = FORMULAS.find((f) => f.formula === name);
    if (!entry) return;
    const col = document.createElement("div");
    col.className = "col-md-3 col-6";
    const btn = document.createElement("button");
    btn.className = "formula-card";
    btn.innerHTML = `<h4>${entry.formula}</h4><p>${escapeHtml(entry.category)}</p>`;
    btn.addEventListener("click", () => showFormulaDetail(entry.formula));
    col.appendChild(btn);
    grid.appendChild(col);
  });
}

function showFormulaHome() {
  const intro = document.getElementById("formulaIntro");
  const results = document.getElementById("formulaResults");
  const detail = document.getElementById("formulaDetail");
  const searchInput = document.getElementById("formulaSearchInput");
  if (intro) intro.hidden = false;
  if (results) results.hidden = true;
  if (detail) detail.hidden = true;
  if (searchInput) searchInput.value = "";
  markActiveSidebar(null);
  window.scrollTo({ top: 0, behavior: "smooth" });
  history.replaceState(null, "", "excel.html");
}

function runFormulaSearch(query) {
  const intro = document.getElementById("formulaIntro");
  const results = document.getElementById("formulaResults");
  const detail = document.getElementById("formulaDetail");
  if (!results) return;

  if (!query.trim()) {
    showFormulaHome();
    return;
  }

  const scored = FORMULAS
    .map((entry) => ({ entry, score: scoreFormula(entry, query) }))
    .filter((s) => s.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, 12);

  if (intro) intro.hidden = true;
  if (detail) detail.hidden = true;
  results.hidden = false;

  if (!scored.length) {
    results.innerHTML = `<div class="no-match-box">No formula matching "${escapeHtml(query)}" yet — try a different term, or browse the sidebar.</div>`;
    return;
  }

  const list = document.createElement("div");
  list.className = "match-list";
  scored.forEach(({ entry }) => {
    const btn = document.createElement("button");
    btn.innerHTML = `<span class="m-formula">=${entry.formula}</span><span class="m-description">${escapeHtml(entry.description)}</span>`;
    btn.addEventListener("click", () => showFormulaDetail(entry.formula));
    list.appendChild(btn);
  });
  results.innerHTML = "";
  results.appendChild(list);
}

function showFormulaDetail(formulaName) {
  const entry = FORMULAS.find((f) => f.formula === formulaName);
  const intro = document.getElementById("formulaIntro");
  const results = document.getElementById("formulaResults");
  const detail = document.getElementById("formulaDetail");
  const searchInput = document.getElementById("formulaSearchInput");
  if (!entry || !detail) return;

  if (intro) intro.hidden = true;
  if (results) results.hidden = true;
  if (searchInput) searchInput.value = "";
  markActiveSidebar(formulaName);

  const args = entry.arguments || [];
  const argTable = args.length
    ? `<table class="table table-bordered arg-table">
        <thead><tr><th>Argument</th><th>Description</th></tr></thead>
        <tbody>${args.map((a) => `<tr><td>${escapeHtml(a.name)}${a.optional ? " <em>(optional)</em>" : ""}</td><td>${escapeHtml(a.description)}</td></tr>`).join("")}</tbody>
      </table>`
    : "";

  const examplesHtml = (entry.examples || [])
    .map((ex) => `<div class="example-block"><code>${escapeHtml(ex.formula)}</code><p class="mb-0">${escapeHtml(ex.explanation)}</p></div>`)
    .join("");

  const hrUsesHtml = (entry.hrUses || []).length
    ? `<h2 class="mt-5">Common Uses</h2><ul>${entry.hrUses.map((u) => `<li>${escapeHtml(u)}</li>`).join("")}</ul>`
    : "";

  const mistakesHtml = (entry.commonMistakes || []).length
    ? `<h2 class="mt-5">⚠️ Watch Out For</h2><div class="mistake-block"><ul class="mb-0">${entry.commonMistakes.map((m) => `<li>${escapeHtml(m)}</li>`).join("")}</ul></div>`
    : "";

  const related = (entry.relatedFormulas || []).filter((r) => FORMULAS.some((f) => f.formula === r));
  const relatedHtml = related.length
    ? `<h2 class="mt-5">Related Functions</h2><div>${related.map((r) => `<span class="related-pill" data-formula="${r}">=${r}</span>`).join("")}</div>`
    : "";

  detail.innerHTML = `
    <nav aria-label="breadcrumb">
      <ol class="breadcrumb">
        <li class="breadcrumb-item"><a href="index.html">Home</a></li>
        <li class="breadcrumb-item"><a href="#" id="breadcrumbExcel">Excel</a></li>
        <li class="breadcrumb-item">${entry.level}</li>
        <li class="breadcrumb-item active">${entry.formula}</li>
      </ol>
    </nav>

    <div class="d-flex align-items-center gap-2 flex-wrap mt-4 mb-2">
      <h1 class="mb-0">${entry.formula} Function in Excel</h1>
    </div>
    <div class="mb-4">
      <span class="level-badge ${levelClass(entry.level)}">${entry.level}</span>
      <span class="category-badge">${escapeHtml(entry.category)}</span>
    </div>

    <h2>Syntax</h2>
    <div class="syntax-box">
      <code>${escapeHtml(entry.syntax)}</code>
      <button class="btn btn-outline-success btn-sm" id="copySyntaxBtn"><i class="bi bi-clipboard"></i> Copy</button>
    </div>

    <h2>Description</h2>
    <p>${escapeHtml(entry.description)}</p>

    ${argTable ? `<h2 class="mt-4">Arguments</h2>${argTable}` : ""}

    <h2 class="mt-5">Examples</h2>
    ${examplesHtml}

    ${hrUsesHtml}
    ${mistakesHtml}
    ${relatedHtml}

    <hr class="my-5">
    <a href="#" id="backToLibrary" class="btn btn-outline-secondary">&larr; Excel Function Library</a>
  `;

  detail.hidden = false;
  window.scrollTo({ top: 0, behavior: "smooth" });
  history.replaceState(null, "", `excel.html#${entry.formula}`);

  const copyBtn = document.getElementById("copySyntaxBtn");
  if (copyBtn) {
    copyBtn.addEventListener("click", () => {
      navigator.clipboard.writeText(entry.syntax).then(() => {
        copyBtn.innerHTML = "Copied!";
        setTimeout(() => (copyBtn.innerHTML = '<i class="bi bi-clipboard"></i> Copy'), 1500);
      });
    });
  }
  document.querySelectorAll(".related-pill").forEach((pill) => {
    pill.addEventListener("click", () => showFormulaDetail(pill.dataset.formula));
  });
  ["breadcrumbExcel", "backToLibrary"].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("click", (e) => { e.preventDefault(); showFormulaHome(); });
  });
}

async function initFormulaApp() {
  const app = document.getElementById("formulaApp");
  if (!app) return;

  try {
    const res = await fetch("data/formulas.json");
    FORMULAS = await res.json();
  } catch (err) {
    app.innerHTML = `<div class="no-match-box">Couldn't load the formula library. Try refreshing the page.</div>`;
    return;
  }

  buildFormulaSidebar();
  renderPopularGrid();

  const searchInput = document.getElementById("formulaSearchInput");
  if (searchInput) {
    searchInput.addEventListener("input", () => runFormulaSearch(searchInput.value));
  }
  const searchForm = document.getElementById("formulaSearchForm");
  if (searchForm) searchForm.addEventListener("submit", (e) => e.preventDefault());

  const params = new URLSearchParams(location.search);
  const q = params.get("q");
  const hash = decodeURIComponent(location.hash.replace("#", ""));

  if (hash && FORMULAS.some((f) => f.formula === hash)) {
    showFormulaDetail(hash);
  } else if (q) {
    if (searchInput) searchInput.value = q;
    runFormulaSearch(q);
  }
}

document.addEventListener("DOMContentLoaded", initFormulaApp);
