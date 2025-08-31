/**
 * validation-report (FULL UI)
 * ---------------------------------------------------------------
 * - Reads validation from user-chosen scope/path (msg/flow/global)
 *   Accepts either { logs, counts } or logs[] directly.
 * - Builds the full modern HTML report (filters, pagination, exports, etc.).
 * - Writes HTML to user-chosen scope/path (msg/flow/global).
 * - Optionally writes a filename (fixed) to a user-chosen scope/path.
 *
 * Typical wiring:
 *   inject → data-validation-engine → validation-report → file
 *
 * Node config fields:
 *   inScope,  inPath      : where to read validation from
 *   outScope, outPath     : where to write HTML to
 *   fileScope, filePath   : where to write filename to (optional)
 *   fixedFilename         : a static filename value to write (optional)
 */
module.exports = function(RED){

  // ------ helpers for typed I/O ------
  function readFrom(node, scope, path, msg){
    if (scope === "msg")   return RED.util.getMessageProperty(msg, path);
    if (scope === "flow")  return node.context().flow.get(path);
    /* scope === "global" */return node.context().global.get(path);
  }
  function writeTo(node, scope, path, value, msg){
    if (!path) return;
    if (scope === "msg")   return RED.util.setMessageProperty(msg, path, value, true);
    if (scope === "flow")  return node.context().flow.set(path, value);
    /* scope === "global" */return node.context().global.set(path, value);
  }

  /**
   * Generate the full HTML report (your original function, adapted to accept rules).
   * @param {Array<Object>} results - validation logs array
   * @param {Array<Object>} allRules - Rules
   */
  function generateValidationReport(results, allRules) {
    // ---------- helpers ----------
    const now = new Date().toLocaleString();
    const esc = s => String(s ?? "").replace(/[&<>\"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    const idify = s => esc(s).replace(/[^a-zA-Z0-9_-]+/g, "_");
    const isIssue = lvl => lvl === "error" || lvl === "warning";

    const data = Array.isArray(results) ? results : [];

    // map ruleId -> suggestions[]
    const suggestionMap = Object.fromEntries(
      (Array.isArray(allRules) ? allRules : []).map(r => [r.id, r.suggestions || []])
    );

    // ---------- groupings ----------

    const byRule = {};
    const keyOf = r => r.ruleId ?? r.id ?? r.rule ?? "(unknown)";
    for (const r of data) (byRule[keyOf(r)] ||= []).push(r);

    const ruleBlocks = Object.entries(byRule).map(([ruleKey, rows]) => {
      const description = rows[0]?.description || "";
      const type = rows[0]?.type || "";
      const counts = rows.reduce((a, x) => { a.total++; a[x.level]=(a[x.level]||0)+1; return a; }, {total:0,info:0,warning:0,error:0});
      const status = counts.error ? "error" : counts.warning ? "warning" : "info";
      const issues = (counts.error||0)+(counts.warning||0);
      const anchor = "rule_" + idify(ruleKey);
      return { ruleId: ruleKey, description, type, rows, counts, status, issues, anchor };
    });

    // By sheet (source or target, excluding engine)
    const sheetNames = new Set();
    data.forEach(x => {
      if (x.source_sheet && x.source_sheet !== "(engine)") sheetNames.add(x.source_sheet);
      if (x.target_sheet && x.target_sheet !== "(engine)") sheetNames.add(x.target_sheet);
    });
    const bySheet = {};
    for (const s of sheetNames) {
      const rows = data.filter(x => x.source_sheet === s || x.target_sheet === s);
      const counts = rows.reduce((a, r) => { a.total++; a[r.level]=(a[r.level]||0)+1; return a; }, {total:0,info:0,warning:0,error:0});
      const status = counts.error ? "error" : counts.warning ? "warning" : "info";
      bySheet[s] = { rows, counts, status };
    }

    // ---------- summary ----------
    const totalRules = ruleBlocks.length;
    const totalRows  = data.length;
    const rulesPassed   = ruleBlocks.filter(r => r.status === "info").length;
    const rulesWarn     = ruleBlocks.filter(r => r.status === "warning").length;
    const rulesErr      = ruleBlocks.filter(r => r.status === "error").length;
    const infoRows = data.filter(x=>x.level==="info").length;
    const warnRows = data.filter(x=>x.level==="warning").length;
    const errRows  = data.filter(x=>x.level==="error").length;

    // ---------- CSS ----------
    const style = `
  <style>
    :root{
      --bg:#0f1420; --fg:#e6e9ef; --muted:#a8b0bf;
      --card:#141a2a; --card2:#1a2135; --border:#2b3a57; --chip:#24314d; --mark:#ffe58a;
      --ok:#27c93f; --warn:#f5a524; --err:#ff6b6b; --link:#7db2ff; --badge:#2a3553;
      --thead:#173052; --theadTxt:#dfe8ff;
      --tableOdd:#121a2b; --tableEven:#0f1727;
      --shadow:0 10px 24px rgba(0,0,0,.35);
    }
    /* light mode */
    body.light{
      --bg:#f7f9fc; --fg:#1d2636; --muted:#526079;
      --card:#ffffff; --card2:#f1f5fb; --border:#d8e0ef; --chip:#e8eef9; --mark:#fff2a6;
      --ok:#16a34a; --warn:#d97706; --err:#ef4444; --link:#1d4ed8; --badge:#e9eefb;
      --thead:#e5efff; --theadTxt:#1f2b49;
      --tableOdd:#ffffff; --tableEven:#f6f9ff;
      --shadow:0 8px 16px rgba(16,24,40,.08);
    }

    html,body{height:100%}
    body{ background:var(--bg); color:var(--fg); font:14px/1.45 system-ui,Segoe UI,Roboto,Arial; margin:16px; }
    h1{ margin:0 0 6px; font-size:24px; }
    .muted{ color:var(--muted); }

    /* ===== Fixed header dock (title + toolbar) ===== */
    #headerDock{
      position: fixed;
      top: 12px; left: 12px; right: 12px;
      z-index: 996;
      background: var(--bg);
      border: 1px solid var(--border);
      border-radius: 12px;
      box-shadow: var(--shadow);
      backdrop-filter: blur(2px);
    }
    #headerDock .title-wrap{ padding: 8px 12px 6px; }
    #headerDock .toolbar{ margin: 6px 12px 10px; }
    body.has-header-offset{ padding-top: calc(var(--headerH, 0px) + 16px); }
    section.rule{ scroll-margin-top: calc(var(--headerH, 0px) + 20px); }

    /* ===== Toolbar ===== */
    .toolbar{ display:flex; align-items:center; gap:10px; flex-wrap:wrap; }
    .toolbar .input{ flex:1 1 120px; min-width:100px; }
    .toolbar .group{ flex:0 0 auto; }
    .toolbar .group.export{ margin-left:auto; }

    .btn,.chip,.select,.input{
      background:var(--card2); border:1px solid var(--border); color:var(--fg);
      border-radius:10px; padding:7px 10px; font-size:13px;
    }
    .btn{ cursor:pointer; box-shadow:var(--shadow); }
    .btn:hover{ filter:brightness(1.05); }
    .select{ padding-right:28px; }

    .toolbar .group{
      display:flex; align-items:center; gap:8px;
      background:var(--card2); border:1px solid var(--border);
      padding:6px 8px; border-radius:12px;
    }
    .toolbar .group-title{
      font-size:12px; font-weight:700; text-transform:uppercase;
      letter-spacing:.04em; opacity:.8; margin-right:2px;
    }
    .toolbar .divider{
      align-self:stretch; width:1px; background:var(--border); margin:0 6px;
    }

    .chip{ display:inline-flex; gap:6px; align-items:center; cursor:pointer; user-select:none; background:var(--chip); }
    .chip.active{ outline:2px solid var(--border); box-shadow:0 0 0 2px rgba(255,255,255,.04) inset; }
    .dot{ width:9px; height:9px; border-radius:50%; display:inline-block; }
    .ok{ background:var(--ok); } .warn{ background:var(--warn);} .err{ background:var(--err); }
    .tag{ font-size:12px; padding:2px 6px; border-radius:999px; background:var(--badge); }

    .sr-only{
      position:absolute; width:1px; height:1px; padding:0; margin:-1px;
      overflow:hidden; clip:rect(0,0,0,0); white-space:nowrap; border:0;
    }

    /* ===== Summary cards ===== */
    .summary{ display:grid; gap:10px; grid-template-columns:repeat(6,minmax(140px,1fr)); margin:8px 0 14px; }
    .card{ background:var(--card); border:1px solid var(--border); border-radius:12px; padding:10px 12px; box-shadow:var(--shadow); }
    .big{ font-size:18px; font-weight:700; }
    .pill{ display:inline-flex; gap:6px; align-items:center; padding:2px 8px; border-radius:999px; background:var(--badge); }

    /* ===== Rule / section cards ===== */
    .rule{ background:var(--card); border:1px solid var(--border); border-radius:12px; margin:16px 0; box-shadow:var(--shadow); }
    .rule-hd{ padding:12px 14px; border-bottom:1px solid var(--border); display:flex; align-items:center; gap:10px; }
    .rule-title{ font-size:18px; font-weight:700; margin-right:auto; }
    .status-badge{ padding:2px 10px; border-radius:999px; font-weight:700; background:var(--badge); }
    .status-badge.ok{ border:1px solid var(--ok); } .status-badge.warn{ border:1px solid var(--warn); } .status-badge.err{ border:1px solid var(--err); }
    .rule-desc{ padding:8px 14px; color:var(--muted); }
    details.rule-body{ padding:0 0 10px; }

    .sec{ margin:10px 14px 16px; border:1px solid var(--border); border-radius:10px; overflow:hidden; }
    .sec-h{ background:var(--thead); color:var(--theadTxt); padding:8px 12px; font-weight:700; }
    .table-wrap{ background:transparent; }

    /* ===== Table ===== */
    table{ border-collapse:collapse; width:100%; }
    thead th{
      background:var(--thead); color:var(--theadTxt); padding:8px; text-align:left;
      border-bottom:1px solid var(--border); position:sticky; top:0; z-index:1;
    }
    tbody td{ border-bottom:1px solid var(--border); padding:8px; }
    tbody tr:nth-child(odd){ background:var(--tableOdd); }
    tbody tr:nth-child(even){ background:var(--tableEven); }
    .lvl{ font-weight:700; }
    .lvl.info{ color:var(--ok); } .lvl.warning{ color:var(--warn);} .lvl.error{ color:var(--err); }

    /* inline “why” rows */
    .why{ font-size:12px; color:var(--muted); margin-top:6px; border-left:2px solid var(--border); padding-left:8px; }

    /* ===== Pagination ===== */
    .pager{ display:flex; gap:6px; align-items:center; padding:10px 12px; background:var(--card2); border-top:1px solid var(--border); }
    .pager .btn{ padding:6px 8px; }

    /* ===== Responsive ===== */
    @media (max-width: 900px){
      .toolbar .input{ flex-basis:100%; min-width:0; }
    }

    /* ===== Print ===== */
    @media print{
      #headerDock{ display:none; }
      body{ background:#fff; color:#000; }
      .rule{ page-break-inside:avoid; }
    }

    /* Export dropdown */
    .dropdown{ position:relative; }
    .dropdown .menu-btn{ display:inline-flex; align-items:center; gap:6px; }
    .dropdown .menu-list{
      position:absolute; top:calc(100% + 6px); right:0;
      background:var(--card); border:1px solid var(--border);
      border-radius:12px; padding:6px; min-width:200px;
      box-shadow:var(--shadow); display:none; z-index:1000;
    }
    .dropdown .menu-list.open{ display:block; }
    .dropdown .menu-item{
      display:block; width:100%; text-align:left;
      border-radius:10px; margin:4px 0; padding:8px 10px;
      background:var(--card2); border:1px solid var(--border);
    }
    .dropdown .menu-item:hover{ filter:brightness(1.05); }
    .dropdown .menu-sep{ height:1px; background:var(--border); margin:6px 2px; border-radius:99px; }
    .dropdown .menu-btn, .dropdown .menu-item { color: var(--fg); }
  </style>`;

    // ---------- HTML: toolbar ----------
    const toolbar = `
  <div class="toolbar">
    <input id="searchBox" class="input"
      placeholder="🔎 Search rows (value / sheet / rule)…"
      title="Filters rows only (inside visible tables and rules)"
      oninput="applyFilters()"/>

    <span class="divider"></span>

    <div class="group" role="group" aria-labelledby="lblRulesFilter">
      <span id="lblRulesFilter" class="group-title">Rules</span>
      <label class="sr-only" for="levelSelect">Rule status</label>
      <select id="levelSelect" class="select"
              title="Show rules that contain at least one row of this level"
              onchange="applyFilters()">
        <option value="">All rule statuses</option>
        <option value="info">Rules with Info</option>
        <option value="warning">Rules with Warning</option>
        <option value="error">Rules with Error</option>
      </select>
    </div>

    <span class="divider"></span>

    <div class="group" role="group" aria-labelledby="lblTableFilter">
      <span id="lblTableFilter" class="group-title">Tables</span>
      <span class="chip active" id="chip-info"    title="Toggle Info sub-tables"    onclick="toggleChip('info')"><span class="dot ok"></span><span class="tag" id="count-info">${infoRows}</span></span>
      <span class="chip active" id="chip-warning" title="Toggle Warning sub-tables" onclick="toggleChip('warning')"><span class="dot warn"></span><span class="tag" id="count-warn">${warnRows}</span></span>
      <span class="chip active" id="chip-error"   title="Toggle Error sub-tables"   onclick="toggleChip('error')"><span class="dot err"></span><span class="tag" id="count-err">${errRows}</span></span>
    </div>

    <span class="divider"></span>

    <div class="group" role="group" aria-labelledby="lblView">
      <span id="lblView" class="group-title">View</span>
      <button class="btn" onclick="foldAll()"     title="Unfold all tables in all rules">▼</button>
      <button class="btn" onclick="collapseAll()" title="Fold all tables in all rules">▲</button>
      <button class="btn" onclick="toggleLight()"> ☀ </button>
    </div>

    <div class="group export dropdown" role="group" aria-labelledby="lblExport">
      <span id="lblExport" class="group-title">Export</span>
      <button id="btnExport" class="btn menu-btn" aria-haspopup="true" aria-expanded="false"
              onclick="toggleExportMenu(event)">
        Export ▾
      </button>
      <div id="exportMenu" class="menu-list" role="menu" aria-labelledby="btnExport">
        <button class="menu-item" role="menuitem" onclick="exportIssuesCSV()">↓ Issues CSV</button>
        <button class="menu-item" role="menuitem" onclick="exportVisibleJSON()">↓ Visible JSON</button>
        <button class="menu-item" role="menuitem" onclick="copyVisible()">📋 Copy Visible</button>
        <div class="menu-sep" aria-hidden="true"></div>
        <button class="menu-item" role="menuitem" onclick="window.print()">🖨 Print</button>
      </div>
    </div>
  </div>`;

    // ---------- Summary ----------
    const summary = `
  <div class="summary">
    <div class="card"><div class="muted">Total rules</div><div class="big">${totalRules}</div></div>
    <div class="card"><div class="muted">Total rows</div><div class="big">${totalRows}</div></div>
    <div class="card"><div class="muted">Rules passed</div><div class="big">${rulesPassed}</div></div>
    <div class="card"><div class="muted">Rules with warnings</div><div class="big">${rulesWarn}</div></div>
    <div class="card"><div class="muted">Rules with errors</div><div class="big">${rulesErr}</div></div>
    <div class="card"><div class="muted">Generated</div><div class="big">${esc(now)}</div></div>

    <div class="card"><span class="pill"><span class="dot ok"></span>Info rows</span> <b>${infoRows}</b></div>
    <div class="card"><span class="pill"><span class="dot warn"></span>Warning rows</span> <b>${warnRows}</b></div>
    <div class="card"><span class="pill"><span class="dot err"></span>Error rows</span> <b>${errRows}</b></div>
  </div>`;

    // ---------- Row renderers ----------
    const renderMainRow = (r, idx, withRuleCol = false) => {
      const yesNo = r.level === "info"
        ? `<span class="status-yes">YES</span>`
        : `<span class="status-no">NO</span>`;

      const cells = [
        `<td>${idx}</td>`,
        `<td>${esc(r.source_sheet || "")}</td>`,
        `<td class="val-cell">${esc(r.value ?? r.message ?? "")}</td>`,
        `<td>${esc(r.type)}</td>`,
        `<td>${esc(r.target_sheet || "")}</td>`
      ];

      if (withRuleCol) cells.push(`<td>${esc(r.ruleId)}</td>`);

      // Always render Status THEN Level explicitly
      cells.push(
        `<td>${yesNo}</td>`,
        `<td class="lvl ${esc(r.level)}">${esc(r.level)}</td>`
      );

      // Suggestions
      const suggestions = suggestionMap[r.ruleId] || [];
      const suggestionsRow = isIssue(r.level)
        ? `<tr class="why-row" data-level="${esc(r.level)}">
            <td colspan="${withRuleCol ? 8 : 7}">
              <div class="why">
                <div><b>Suggestions</b></div>
                ${
                  suggestions.length
                    ? `<ul>${suggestions.map(s => `<li>${esc(s)}</li>`).join("")}</ul>`
                    : `<div class="muted">No suggestions provided for this rule.</div>`
                }
              </div>
            </td>
          </tr>`
        : "";

      return `
        <tr data-level="${esc(r.level)}">
          ${cells.join("")}
        </tr>
        ${suggestionsRow}
      `;
    };

    const renderLevelSection = (title, rows, tableId, withRuleCol=false) => {
      if (!rows.length) return "";
      const headCols = withRuleCol
        ? `<th>#</th><th>Source Sheet</th><th>Value</th><th>Type</th><th>Target Sheet</th><th>Rule</th><th>Status</th><th>Level</th>`
        : `<th>#</th><th>Source Sheet</th><th>Value</th><th>Type</th><th>Target Sheet</th><th>Status</th><th>Level</th>`;

      const body = rows.map((r,i)=>renderMainRow(r, i+1, withRuleCol)).join("");

      // derive kind from section title
      const t = title.toLowerCase();
      const kind = t.startsWith('error') ? 'error' : t.startsWith('warning') ? 'warning' : 'info';

      return `
        <div class="sec" data-sec="${esc(tableId)}" data-kind="${kind}">
          <div class="sec-h">${esc(title)}</div>
          <div class="table-wrap">
            <table class="data-table" id="${esc(tableId)}" data-page="1" data-rows="10">
              <thead><tr>${headCols}</tr></thead>
              <tbody>${body}</tbody>
            </table>
            <div class="pager">
              <button class="btn" onclick="chgPage(this,-1)">◀ Prev</button>
              <span class="spacer"></span>
              <span class="muted">Rows per page <span class="rpp">10</span>/<span class="total">0</span></span>
              <select class="select" onchange="setRpp(this)">
                <option selected>10</option><option>25</option><option>50</option><option>100</option>
              </select>
              <span class="spacer"></span>
              <button class="btn" onclick="chgPage(this,1)">Next ▶</button>
            </div>
          </div>
        </div>`;
    };

    // ---------- Sections: by rule ----------
    const renderRule = block => {
      const infos = block.rows.filter(r=>r.level==='info');
      const warns = block.rows.filter(r=>r.level==='warning');
      const errs  = block.rows.filter(r=>r.level==='error');

      const badgeClass = block.status === "error" ? "err" : block.status === "warning" ? "warn" : "ok";

      return `
        <section class="rule"
          data-rule="${esc(block.ruleId)}"
          data-status="${esc(block.status)}"
          data-issues="${block.issues}"
          data-count-info="${block.counts.info||0}"
          data-count-warning="${block.counts.warning||0}"
          data-count-error="${block.counts.error||0}"
          id="${block.anchor}" data-index="0">
          <div class="rule-hd">
            <div class="rule-title">Rule: ${esc(block.ruleId)}</div>
            <span class="status-badge ${badgeClass}">${block.status.toUpperCase()}</span>
            <span class="pill"><span class="dot err"></span>${block.counts.error||0}</span>
            <span class="pill"><span class="dot warn"></span>${block.counts.warning||0}</span>
            <span class="pill"><span class="dot ok"></span>${block.counts.info||0}</span>
          </div>
          <div class="rule-desc">Description: ${esc(block.description)} | <b>Type:</b> ${esc(block.type)}</div>
          <details class="rule-body" open>
            ${renderLevelSection('Errors',   errs,  `tbl_${idify(block.ruleId)}_err`)}
            ${renderLevelSection('Warnings', warns, `tbl_${idify(block.ruleId)}_warn`)}
            ${renderLevelSection('Info',     infos, `tbl_${idify(block.ruleId)}_info`)}
          </details>
        </section>`;
    };
    const sectionsByRule = ruleBlocks.map(renderRule).join("");

    // ---------- Sections: by sheet ----------
    const renderSheet = (name, bundle) => {
      const rows = bundle.rows;
      const infos = rows.filter(r=>r.level==='info');
      const warns = rows.filter(r=>r.level==='warning');
      const errs  = rows.filter(r=>r.level==='error');

      const badgeClass = bundle.status === "error" ? "err" : bundle.status === "warning" ? "warn" : "ok";
      return `
        <section class="rule" data-sheet="${esc(name)}" data-status="${esc(bundle.status)}" data-issues="${(bundle.counts.error||0)+(bundle.counts.warning||0)}" id="sheet_${idify(name)}">
          <div class="rule-hd">
            <div class="rule-title">Sheet: ${esc(name)}</div>
            <span class="status-badge ${badgeClass}">${bundle.status.toUpperCase()}</span>
            <span class="pill"><span class="dot err"></span>${bundle.counts.error||0}</span>
            <span class="pill"><span class="dot warn"></span>${bundle.counts.warning||0}</span>
            <span class="pill"><span class="dot ok"></span>${bundle.counts.info||0}</span>
          </div>
          <details class="rule-body" open>
            ${renderLevelSection('Errors',   errs,  `tbl_${idify(name)}_err`,  true)}
            ${renderLevelSection('Warnings', warns, `tbl_${idify(name)}_warn`, true)}
            ${renderLevelSection('Info',     infos, `tbl_${idify(name)}_info`, true)}
          </details>
        </section>`;
    };
    const sectionsBySheet = Object.entries(bySheet).map(([n,b]) => renderSheet(n,b)).join("");

    // ---------- JS (filters, pagination, exports, theme) ----------
    const script = `
  <script>
    (function(){
      // state
      let chipState = { info:true, warning:true, error:true };
      let grouping = 'rule'; // 'rule' | 'sheet'

      // On load
      window.addEventListener('DOMContentLoaded', () => {
        indexRuleCards();
        paginateAll();
        applyFilters();
        const header = document.getElementById('headerDock');
        if (!header) return;

        const setHeaderOffset = () => {
          const h = header.offsetHeight || 0;
          document.documentElement.style.setProperty('--headerH', h + 'px');
          document.body.classList.add('has-header-offset');
        };

        setHeaderOffset();
        window.addEventListener('resize', setHeaderOffset);
      });

      // Index cards for next/prev
      function indexRuleCards(){
        document.querySelectorAll('#sections-rule section.rule').forEach((el, i)=> el.dataset.index = String(i));
        document.querySelectorAll('#sections-sheet section.rule').forEach((el, i)=> el.dataset.index = String(i));
      }

      // ------------ Pagination (filter-aware) ------------
      function paginateTable(tbl){
        const rpp = parseInt(tbl.getAttribute('data-rows') || '10', 10);
        let page  = parseInt(tbl.getAttribute('data-page') || '1', 10);

        // Main (non-why) rows
        const allRows = Array.from(tbl.querySelectorAll('tbody tr'))
          .filter(tr => !tr.classList.contains('why-row'));

        // Rows that MATCH FILTERS (dataset.pass set by applyFilters)
        const filtered = allRows.filter(tr => tr.dataset.pass !== '0');

        const total = filtered.length;
        const pages = Math.max(1, Math.ceil(total / rpp));
        if (page > pages) page = pages;
        if (page < 1) page = 1;
        tbl.setAttribute('data-page', String(page));

        const start = (page-1) * rpp;
        const end   = start + rpp;

        // Hide all main rows first
        allRows.forEach(tr => { tr.style.display = 'none'; });

        // Show only current page among filtered
        let shown = 0;
        filtered.forEach((tr, i) => {
          const show = (i >= start && i < end);
          if (show) { tr.style.display = ''; shown++; }
        });

        // Sync “why” rows to their preceding main row + also to filter pass
        const bodyRows = Array.from(tbl.querySelectorAll('tbody tr'));
        for (let i = 0; i < bodyRows.length; i++){
          const row = bodyRows[i];
          if (!row.classList.contains('why-row')) continue;
          const prev = bodyRows[i-1];
          const prevPass = prev && prev.dataset.pass !== '0';
          row.style.display = (prev && prev.style.display === '' && prevPass) ? '' : 'none';
        }

        // Update pager label (shown/total)
        const pager   = tbl.closest('.table-wrap')?.querySelector('.pager');
        const rppEl   = pager?.querySelector('.rpp');
        const totalEl = pager?.querySelector('.total');
        if (rppEl)   rppEl.textContent   = String(shown);
        if (totalEl) totalEl.textContent = String(total);

        // Keep dropdown in sync
        const sel = pager?.querySelector('select.select');
        if (sel && sel.value !== String(rpp)) sel.value = String(rpp);
      }
      function paginateAll(){
        document.querySelectorAll('table.data-table').forEach(paginateTable);
      }

      window.chgPage = function(el, dir){
        const tbl = el.closest('.table-wrap').querySelector('table.data-table');
        let page = parseInt(tbl.getAttribute('data-page') || '1', 10);
        page = page + dir;
        tbl.setAttribute('data-page', String(page));
        paginateTable(tbl);
      };

      window.setRpp = function(sel){
        const tbl = sel.closest('.table-wrap').querySelector('table.data-table');
        tbl.setAttribute('data-rows', sel.value);
        tbl.setAttribute('data-page', '1');
        paginateTable(tbl);
      };

      // ------------ Filtering + search highlight ------------
      function clearHighlights(scope){
        scope.querySelectorAll('mark.hl').forEach(m => m.replaceWith(document.createTextNode(m.textContent)));
      }
      function highlight(scope, term){
        if (!term) return;
        const rx = new RegExp(term.replace(/[-\\/^$*+?.()|[\\]{}]/g, '\\\\$&'), 'gi');
        scope.querySelectorAll('.val-cell, .rule-title, .rule-desc, td').forEach(node=>{
          if (!node.childNodes || !node.childNodes.length) return;
          node.childNodes.forEach(n=>{
            if (n.nodeType===3){
              const text = n.textContent;
              if (!text) return;
              const frag = document.createDocumentFragment();
              let last = 0;
              text.replace(rx, (m, idx)=>{
                frag.appendChild(document.createTextNode(text.slice(last, idx)));
                const mark = document.createElement('mark'); mark.className='hl'; mark.textContent = m;
                frag.appendChild(mark);
                last = idx + m.length;
              });
              if (last){
                frag.appendChild(document.createTextNode(text.slice(last)));
                n.parentNode.replaceChild(frag, n);
              }
            }
          });
        });
      }
      
      window.applyFilters = function(){
        const q = (document.getElementById('searchBox').value || '').trim().toLowerCase();
        const ruleFilter = document.getElementById('levelSelect').value; // '', 'info', 'warning', 'error'
        const container = document.getElementById('sections-'+grouping);

        clearHighlights(container);

        // Row-level predicate: search only (chips handled at table level)
        const rowPass = (tr) => {
          const text = tr.innerText.toLowerCase();
          return !q || text.includes(q);
        };

        container.querySelectorAll('section.rule').forEach(sec=>{
          // --- RULE filter: only in 'rule' grouping ---
          let passRule = true;
          if (grouping === 'rule' && ruleFilter) {
            const countInfo    = +(sec.dataset.countInfo    || 0);
            const countWarning = +(sec.dataset.countWarning || 0);
            const countError   = +(sec.dataset.countError   || 0);
            passRule =
              (ruleFilter === 'info'    && countInfo    > 0) ||
              (ruleFilter === 'warning' && countWarning > 0) ||
              (ruleFilter === 'error'   && countError   > 0);
          }

          let ruleHasVisibleRows = false;

          // --- TABLE filter: chips hide/show sub-tables by data-kind ---
          sec.querySelectorAll('.sec').forEach(secBlock=>{
            const kind = secBlock.getAttribute('data-kind'); // 'error' | 'warning' | 'info'

            if (kind && chipState[kind] === false) {
              secBlock.style.display = 'none';
              return;
            }

            const tbl = secBlock.querySelector('table.data-table');
            if (!tbl) { secBlock.style.display = 'none'; return; }

            // apply search to rows
            const mainRows = Array.from(tbl.querySelectorAll('tbody tr'))
              .filter(tr => !tr.classList.contains('why-row'));

            let visCount = 0;
            mainRows.forEach(tr=>{
              const pass = rowPass(tr);
              tr.dataset.pass = pass ? '1' : '0';
              if (pass) visCount++;
            });

            // hide sub-table if empty after search
            if (visCount === 0) {
              secBlock.style.display = 'none';
            } else {
              secBlock.style.display = '';
              ruleHasVisibleRows = true;
            }

            // paginate this table using the flags we just set
            paginateTable(tbl);
          });

          // show/hide the entire rule card
          sec.style.display = (passRule && ruleHasVisibleRows) ? '' : 'none';
        });

        // Highlight AFTER visibility decisions
        if (q) highlight(container, q);
      };

      window.toggleChip = function(kind){
        chipState[kind] = !chipState[kind];
        const el = document.getElementById('chip-'+kind);
        el.classList.toggle('active', chipState[kind]);
        applyFilters();
      }

      // ------------ Grouping & sorting (hooks kept for future) ------------
      window.switchGrouping = function(){
        grouping = document.getElementById('groupSelect').value;
        document.getElementById('sections-rule').style.display = (grouping==='rule') ? '' : 'none';
        document.getElementById('sections-sheet').style.display = (grouping==='sheet') ? '' : 'none';
        applyFilters();
      }

      // ------------ Theme ------------
      window.toggleLight = function(){ document.body.classList.toggle('light'); }

      window.foldAll = function(){ document.querySelectorAll('details.rule-body').forEach(d=>d.open=true); }
      window.collapseAll = function(){ document.querySelectorAll('details.rule-body').forEach(d=>d.open=false); }

      // ------------ Exports ------------
      window.toggleExportMenu = function(){
        const m = document.getElementById('exportMenu');
        if (m) m.classList.toggle('open');
      };
      // close on outside click / Escape
      document.addEventListener('click', (e)=>{
        const menu = document.getElementById('exportMenu');
        if (!menu) return;
        const wrap = menu.closest('.menu');
        if (menu.classList.contains('open') && wrap && !wrap.contains(e.target)){
          menu.classList.remove('open');
        }
      });
      document.addEventListener('keydown', (e)=>{
        if (e.key === 'Escape'){
          const menu = document.getElementById('exportMenu');
          if (menu) menu.classList.remove('open');
        }
      });
      function getVisibleRows(){
        const cont = document.getElementById('sections-'+grouping);
        const rows = [];
        cont.querySelectorAll('section.rule').forEach(sec=>{
          if (sec.style.display==='none') return;
          sec.querySelectorAll('table.data-table tbody tr').forEach(tr=>{
            if (tr.classList.contains('why-row')) return;
            if (tr.classList.contains('hidden-by-filter')) return;
            if (tr.style.display==='none') return;
            const tds = tr.querySelectorAll('td');
            const ctx = grouping==='rule' ? {ruleId: sec.getAttribute('data-rule')} : {sheet: sec.getAttribute('data-sheet')};
            rows.push({
              idx: tds[0]?.innerText.trim(),
              source: tds[1]?.innerText.trim(),
              value: tds[2]?.innerText.trim(),
              type: tds[3]?.innerText.trim(),
              target: tds[4]?.innerText.trim(),
              status: tds[5]?.innerText.trim(),
              level: tds[6]?.innerText.trim(),
              ...ctx
            });
          });
        });
        return rows;
      }
      function download(name, text, mime='text/plain'){
        const blob = new Blob([text], {type:mime});
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob); a.download=name; a.click();
        setTimeout(()=>URL.revokeObjectURL(a.href), 500);
      }
      window.exportIssuesCSV = function(){
        const rows = getVisibleRows().filter(r => /warning|error/i.test(r.level||''));
        const head = Object.keys(rows[0] || {idx:'',source:'',value:'',type:'',target:'',status:'',level:''});
        const escCsv = s => '"' + String(s??'').replace(/"/g,'""') + '"';
        const csv = [head.join(',')].concat(rows.map(r => head.map(k=>escCsv(r[k])).join(','))).join('\\n');
        download('validation_issues.csv', csv, 'text/csv');
      }
      window.exportVisibleJSON = function(){
        download('validation_visible.json', JSON.stringify(getVisibleRows(), null, 2), 'application/json');
      }
      window.copyVisible = async function(){
        const rows = getVisibleRows().map(r => Object.values(r).join('\\t')).join('\\n');
        try { await navigator.clipboard.writeText(rows); alert('Copied visible rows'); }
        catch { alert('Clipboard copy failed'); }
      }
      // --- Export dropdown helpers ---
      window.toggleExportMenu = function(e){
        const menu = document.getElementById('exportMenu');
        const btn  = document.getElementById('btnExport');
        const open = !menu.classList.contains('open');
        menu.classList.toggle('open', open);
        btn.setAttribute('aria-expanded', open ? 'true' : 'false');
        if (open) {
          // close on outside click or ESC
          const close = (ev) => {
            if (ev.type === 'keydown' && ev.key !== 'Escape') return;
            if (ev.type === 'click' && menu.contains(ev.target)) return;
            menu.classList.remove('open');
            btn.setAttribute('aria-expanded', 'false');
            document.removeEventListener('click', close, true);
            document.removeEventListener('keydown', close, true);
          };
          document.addEventListener('click', close, true);
          document.addEventListener('keydown', close, true);
        }
      };
    })();
  </script>`;

    // ---------- assemble ----------
    return `
<!doctype html>
<html>
<head>
<meta charset="UTF-8">
<title>Validation Report</title>
${style}
</head>
<body>
  <div class="title-wrap">
    <h1>Validation Report</h1>
    <div class="muted">Generated on ${esc(now)}</div>
  </div>
  <div id="headerDock">
    ${toolbar}
  </div>

  ${summary}

  <!-- Rule grouping -->
  <div id="sections-rule">
    ${sectionsByRule}
  </div>

  <!-- Sheet grouping (hidden by default) -->
  <div id="sections-sheet" style="display:none">
    ${sectionsBySheet}
  </div>
  ${script}
</body>
</html>`;
  }

  function ValidationReport(config){
    RED.nodes.createNode(this, config);
    const node = this;

    // inputs
    node.inScope = config.inScope || "msg";
    node.inPath  = config.inPath  || "validation";

    // outputs (HTML)
    node.outScope = config.outScope || "msg";
    node.outPath  = config.outPath  || "payload";

    // optional filename out
    node.fileScope = config.fileScope || "msg";
    node.filePath  = config.filePath  || "filename";
    node.fixedFilename = config.fixedFilename || "";

    node.on("input", (msg, send, done)=>{
      try{
        // read validation from chosen scope
        const src = readFrom(node, node.inScope, node.inPath, msg);
        if (!src){
          node.status({fill:"red",shape:"ring",text:"no validation found"});
          send(msg); return done && done();
        }

        // normalize to logs[]
        const logs = Array.isArray(src) ? src : (Array.isArray(src.logs) ? src.logs : []);
        // build HTML
        const html = generateValidationReport(logs, []);

        // write HTML to chosen destination
        writeTo(node, node.outScope, node.outPath, html, msg);

        // filename (if a fixed one was provided, set it; else keep existing)
        const filename = (node.fixedFilename||"").trim();
        if (filename){
          writeTo(node, node.fileScope, node.filePath, filename, msg);
        }

        // status + pass-through
        const counts = src.counts || logs.reduce((a,r)=>{
          const lvl = String(r.level||"info").toLowerCase();
          if (lvl.startsWith("err")) a.error++;
          else if (lvl.startsWith("warn")) a.warning++;
          else a.info++;
          return a;
        }, {info:0,warning:0,error:0,total:0});
        counts.total = counts.info + counts.warning + counts.error;

        const worst = counts.error ? "error" : counts.warning ? "warning" : "info";
        node.status({fill: worst==="error"?"red":worst==="warning"?"yellow":"green", shape:"dot",
          text:`E:${counts.error||0} W:${counts.warning||0} I:${counts.info||0}`});
        send(msg);
        done && done();
      }catch(e){
        node.status({fill:"red",shape:"ring",text:"runtime error"});
        done ? done(e) : node.error(e);
      }
    });
  }

  RED.nodes.registerType("validation-report", ValidationReport);
};
