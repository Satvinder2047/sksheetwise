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

  const stepsHtml = (entry.steps || []).length
    ? `<h2 class="mt-4">Step by Step</h2><ol class="step-list">${entry.steps.map((s) => `<li>${escapeHtml(s)}</li>`).join("")}</ol>`
    : "";

  const whyUseHtml = entry.whyUse
    ? `<h2 class="mt-4">Why & When to Use It</h2><p>${escapeHtml(entry.whyUse)}</p>`
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

    ${whyUseHtml}

    ${argTable ? `<h2 class="mt-4">Arguments</h2>${argTable}` : ""}

    ${stepsHtml}

    <h2 class="mt-5">Examples</h2>
    ${examplesHtml}

    ${entry.hasPractice ? `<h2 class="mt-5">🧪 Practice It Yourself</h2><div id="practiceMount"></div>` : ""}

    ${hrUsesHtml}
    ${mistakesHtml}
    ${relatedHtml}

    <hr class="my-5">
    <a href="#" id="backToLibrary" class="btn btn-outline-secondary">&larr; Excel Function Library</a>
  `;

  detail.hidden = false;
  window.scrollTo({ top: 0, behavior: "smooth" });
  history.replaceState(null, "", `excel.html#${entry.formula}`);

  if (entry.hasPractice) mountPractice(entry.formula);

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

// =======================================
// Practice widgets — real, live-calculating
// mini spreadsheets for the most-searched formulas
// =======================================

function cellInput(value, width) {
  const input = document.createElement("input");
  input.type = "text";
  input.className = "practice-cell";
  input.value = value;
  if (width) input.style.width = width;
  return input;
}

function practiceWrap(instructions, tableEl, resultLabel, resultEl) {
  const wrap = document.createElement("div");
  wrap.className = "practice-widget";
  const p = document.createElement("p");
  p.className = "practice-instructions";
  p.textContent = instructions;
  wrap.appendChild(p);
  wrap.appendChild(tableEl);
  const resultBox = document.createElement("div");
  resultBox.className = "practice-result";
  resultBox.innerHTML = `<span class="practice-result-label">${resultLabel}</span>`;
  resultBox.appendChild(resultEl);
  wrap.appendChild(resultBox);
  return wrap;
}

function buildIfPractice(mount) {
  const table = document.createElement("table");
  table.className = "table table-bordered practice-table";
  table.innerHTML = `<thead><tr><th>Student</th><th>Marks</th><th>Result</th></tr></thead>`;
  const tbody = document.createElement("tbody");
  const students = ["Amit", "Priya", "Rahul"];
  const marksInputs = [];
  const resultCells = [];
  students.forEach((name, i) => {
    const tr = document.createElement("tr");
    const tdName = document.createElement("td");
    tdName.textContent = name;
    const tdMarks = document.createElement("td");
    const input = cellInput([72, 28, 55][i], "70px");
    marksInputs.push(input);
    tdMarks.appendChild(input);
    const tdResult = document.createElement("td");
    tdResult.className = "fw-bold";
    resultCells.push(tdResult);
    tr.append(tdName, tdMarks, tdResult);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  const formulaLine = document.createElement("div");
  formulaLine.className = "practice-formula-line";

  function recalc() {
    marksInputs.forEach((input, i) => {
      const val = Number(input.value);
      const pass = !isNaN(val) && val >= 35;
      resultCells[i].textContent = pass ? "Pass" : "Fail";
      resultCells[i].style.color = pass ? "#198754" : "#c0392b";
    });
    formulaLine.textContent = `=IF(marks>=35,"Pass","Fail")`;
  }
  marksInputs.forEach((input) => input.addEventListener("input", recalc));
  recalc();

  const wrap = document.createElement("div");
  wrap.className = "practice-widget";
  wrap.innerHTML = `<p class="practice-instructions">Edit the marks — watch the Result column update live using =IF(marks>=35,"Pass","Fail").</p>`;
  wrap.appendChild(table);
  mount.appendChild(wrap);
}

function buildSumPractice(mount) {
  const rows = ["Travel", "Meals", "Supplies", "Software"];
  const table = document.createElement("table");
  table.className = "table table-bordered practice-table";
  table.innerHTML = `<thead><tr><th>Expense</th><th>Amount</th></tr></thead>`;
  const tbody = document.createElement("tbody");
  const inputs = [];
  const defaults = [1200, 450, 300, 850];
  rows.forEach((label, i) => {
    const tr = document.createElement("tr");
    const td1 = document.createElement("td");
    td1.textContent = label;
    const td2 = document.createElement("td");
    const input = cellInput(defaults[i], "90px");
    inputs.push(input);
    td2.appendChild(input);
    tr.append(td1, td2);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  const resultEl = document.createElement("span");
  resultEl.className = "practice-result-value";

  function recalc() {
    const total = inputs.reduce((sum, inp) => sum + (Number(inp.value) || 0), 0);
    resultEl.textContent = `=SUM(range) → ${total.toLocaleString()}`;
  }
  inputs.forEach((inp) => inp.addEventListener("input", recalc));
  recalc();

  mount.appendChild(practiceWrap(
    "Edit any amount — the total recalculates live using =SUM(range).",
    table, "", resultEl
  ));
}

function buildAveragePractice(mount) {
  const rows = ["Week 1", "Week 2", "Week 3", "Week 4"];
  const table = document.createElement("table");
  table.className = "table table-bordered practice-table";
  table.innerHTML = `<thead><tr><th>Week</th><th>Hours worked</th></tr></thead>`;
  const tbody = document.createElement("tbody");
  const inputs = [];
  const defaults = [38, 41, 36, 40];
  rows.forEach((label, i) => {
    const tr = document.createElement("tr");
    const td1 = document.createElement("td");
    td1.textContent = label;
    const td2 = document.createElement("td");
    const input = cellInput(defaults[i], "80px");
    inputs.push(input);
    td2.appendChild(input);
    tr.append(td1, td2);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  const resultEl = document.createElement("span");
  resultEl.className = "practice-result-value";

  function recalc() {
    const vals = inputs.map((inp) => Number(inp.value)).filter((v) => !isNaN(v));
    const avg = vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
    resultEl.textContent = `=AVERAGE(range) → ${avg.toFixed(1)}`;
  }
  inputs.forEach((inp) => inp.addEventListener("input", recalc));
  recalc();

  mount.appendChild(practiceWrap(
    "Edit any week's hours — the average recalculates live using =AVERAGE(range).",
    table, "", resultEl
  ));
}

function buildCountifPractice(mount) {
  const table = document.createElement("table");
  table.className = "table table-bordered practice-table";
  table.innerHTML = `<thead><tr><th>Employee</th><th>Department</th></tr></thead>`;
  const tbody = document.createElement("tbody");
  const names = ["Amit", "Priya", "Rahul", "Sana", "Divya"];
  const defaults = ["Sales", "HR", "Sales", "IT", "Sales"];
  const inputs = [];
  names.forEach((name, i) => {
    const tr = document.createElement("tr");
    const td1 = document.createElement("td");
    td1.textContent = name;
    const td2 = document.createElement("td");
    const input = cellInput(defaults[i], "100px");
    inputs.push(input);
    td2.appendChild(input);
    tr.append(td1, td2);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  const criteriaRow = document.createElement("div");
  criteriaRow.className = "practice-criteria-row";
  const criteriaInput = cellInput("Sales", "120px");
  criteriaRow.innerHTML = `<label class="me-2">Count how many equal:</label>`;
  criteriaRow.appendChild(criteriaInput);

  const resultEl = document.createElement("span");
  resultEl.className = "practice-result-value";

  function recalc() {
    const q = criteriaInput.value.trim().toLowerCase();
    const count = inputs.filter((inp) => inp.value.trim().toLowerCase() === q).length;
    resultEl.textContent = `=COUNTIF(range,"${criteriaInput.value}") → ${count}`;
  }
  inputs.forEach((inp) => inp.addEventListener("input", recalc));
  criteriaInput.addEventListener("input", recalc);
  recalc();

  const wrap = practiceWrap(
    "Edit departments, or change the criteria below — the count updates live.",
    table, "", resultEl
  );
  wrap.insertBefore(criteriaRow, wrap.querySelector(".practice-result"));
  mount.appendChild(wrap);
}

function buildVlookupPractice(mount, useXlookupLabel) {
  const table = document.createElement("table");
  table.className = "table table-bordered practice-table";
  table.innerHTML = `<thead><tr><th>ID</th><th>Name</th><th>Department</th></tr></thead>`;
  const tbody = document.createElement("tbody");
  const data = [
    ["A101", "Amit", "Sales"],
    ["A102", "Priya", "HR"],
    ["A103", "Rahul", "IT"],
  ];
  const rowsInputs = [];
  data.forEach((row) => {
    const tr = document.createElement("tr");
    const rowInputs = row.map((val, i) => {
      const td = document.createElement("td");
      const input = cellInput(val, i === 0 ? "70px" : "90px");
      td.appendChild(input);
      tr.appendChild(td);
      return input;
    });
    rowsInputs.push(rowInputs);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  const criteriaRow = document.createElement("div");
  criteriaRow.className = "practice-criteria-row";
  const lookupInput = cellInput("A102", "90px");
  criteriaRow.innerHTML = `<label class="me-2">Look up ID:</label>`;
  criteriaRow.appendChild(lookupInput);

  const resultEl = document.createElement("span");
  resultEl.className = "practice-result-value";

  function recalc() {
    const q = lookupInput.value.trim().toLowerCase();
    const match = rowsInputs.find((r) => r[0].value.trim().toLowerCase() === q);
    const fnName = useXlookupLabel ? "XLOOKUP" : "VLOOKUP";
    resultEl.textContent = match
      ? `=${fnName}(...) → ${match[2].value}`
      : `=${fnName}(...) → Not found`;
  }
  rowsInputs.flat().forEach((inp) => inp.addEventListener("input", recalc));
  lookupInput.addEventListener("input", recalc);
  recalc();

  const wrap = practiceWrap(
    "Edit the table, or change the ID you're looking up — the result updates live.",
    table, "", resultEl
  );
  wrap.insertBefore(criteriaRow, wrap.querySelector(".practice-result"));
  mount.appendChild(wrap);
}

function buildSumifsPractice(mount) {
  const table = document.createElement("table");
  table.className = "table table-bordered practice-table";
  table.innerHTML = `<thead><tr><th>Region</th><th>Month</th><th>Amount</th></tr></thead>`;
  const tbody = document.createElement("tbody");
  const data = [
    ["West", "Jan", 4200],
    ["East", "Jan", 3100],
    ["West", "Feb", 5000],
    ["West", "Jan", 1800],
  ];
  const rowsInputs = [];
  data.forEach((row) => {
    const tr = document.createElement("tr");
    const rowInputs = row.map((val, i) => {
      const td = document.createElement("td");
      const input = cellInput(val, i === 2 ? "80px" : "80px");
      td.appendChild(input);
      tr.appendChild(td);
      return input;
    });
    rowsInputs.push(rowInputs);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  const criteriaRow = document.createElement("div");
  criteriaRow.className = "practice-criteria-row";
  const regionInput = cellInput("West", "80px");
  const monthInput = cellInput("Jan", "70px");
  criteriaRow.innerHTML = `<label class="me-2">Region equals:</label>`;
  criteriaRow.appendChild(regionInput);
  const monthLabel = document.createElement("label");
  monthLabel.className = "ms-3 me-2";
  monthLabel.textContent = "and Month equals:";
  criteriaRow.appendChild(monthLabel);
  criteriaRow.appendChild(monthInput);

  const resultEl = document.createElement("span");
  resultEl.className = "practice-result-value";

  function recalc() {
    const r = regionInput.value.trim().toLowerCase();
    const m = monthInput.value.trim().toLowerCase();
    const total = rowsInputs
      .filter((row) => row[0].value.trim().toLowerCase() === r && row[1].value.trim().toLowerCase() === m)
      .reduce((sum, row) => sum + (Number(row[2].value) || 0), 0);
    resultEl.textContent = `=SUMIFS(Amount, Region,"${regionInput.value}", Month,"${monthInput.value}") → ${total.toLocaleString()}`;
  }
  rowsInputs.flat().forEach((inp) => inp.addEventListener("input", recalc));
  [regionInput, monthInput].forEach((inp) => inp.addEventListener("input", recalc));
  recalc();

  const wrap = practiceWrap(
    "Edit the table or the two conditions below — the total updates live.",
    table, "", resultEl
  );
  wrap.insertBefore(criteriaRow, wrap.querySelector(".practice-result"));
  mount.appendChild(wrap);
}

const PRACTICE_BUILDERS = {
  IF: buildIfPractice,
  SUM: buildSumPractice,
  AVERAGE: buildAveragePractice,
  COUNTIF: buildCountifPractice,
  VLOOKUP: (mount) => buildVlookupPractice(mount, false),
  XLOOKUP: (mount) => buildVlookupPractice(mount, true),
  SUMIFS: buildSumifsPractice,
};

function mountPractice(formulaName) {
  const mount = document.getElementById("practiceMount");
  if (!mount) return;
  const builder = PRACTICE_BUILDERS[formulaName];
  if (builder) builder(mount);
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
