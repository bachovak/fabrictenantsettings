"""
build.py — Generate docs/index.html from data/settings.json.

The output is a fully self-contained HTML file (inline CSS + JS) styled
to match kristinabachova.com — same colour tokens, fonts, and component
patterns as styles.css.
"""

import json
from datetime import date
from pathlib import Path

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
REPO_ROOT = Path(__file__).resolve().parent.parent
SETTINGS_FILE = REPO_ROOT / "data" / "settings.json"
OUTPUT_FILE = REPO_ROOT / "docs" / "index.html"

MS_DOCS_URL = (
    "https://learn.microsoft.com/en-us/fabric/admin/tenant-settings-index"
)
MAINTAINER_URL = "https://kristinabachova.com"


# ---------------------------------------------------------------------------
# Badge colour classification
# ---------------------------------------------------------------------------

def classify_default(value: str) -> str:
    if not value:
        return "amber"
    v = value.lower()
    if v.startswith("off"):
        return "red"
    if v == "on" or v.startswith("on \u2013 all") or v.startswith("on – all"):
        return "green"
    return "amber"


def classify_recommended(value: str) -> str:
    if not value:
        return "amber"
    v = value.lower()
    if v.startswith("off"):
        return "red"
    if v.startswith("on"):
        return "green"
    return "amber"


def classify_requires(value: str) -> str:
    if not value:
        return "amber"
    return "red" if value.strip().lower() == "yes" else "green"


def classify_impact(value: str) -> str:
    if not value:
        return "amber"
    v = value.strip().lower()
    if v == "high":
        return "red"
    if v == "medium":
        return "amber"
    return "green"


# ---------------------------------------------------------------------------
# HTML helpers
# ---------------------------------------------------------------------------

def escape_html(text: str) -> str:
    if not text:
        return ""
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&#39;")
    )


def build_row(s: dict) -> str:
    status = s.get("status", "current")
    is_removed = status == "removed"
    is_pending = status == "pending_review"

    row_class = ""
    if is_removed:
        row_class = ' class="row-removed"'
    elif is_pending:
        row_class = ' class="row-pending"'

    notes = s.get("notes") or ""
    if is_pending:
        notes = "New setting detected — pending human review."

    name_html = escape_html(s.get("name", ""))
    if is_pending:
        name_html += ' <span class="badge badge--pending">&#9888; New</span>'

    default_val = s.get("default_setting") or ""
    rec_val     = s.get("recommended_setting") or ""
    req_val     = s.get("requires_adjustment") or ""
    impact_val  = s.get("impact") or ""

    def badge(val, colour):
        if not val:
            return '<span class="badge badge--empty">&mdash;</span>'
        return f'<span class="badge badge--{colour}">{escape_html(val)}</span>'

    return f"""    <tr{row_class}
      data-name="{escape_html(s.get('name',''))}"
      data-category="{escape_html(s.get('category',''))}"
      data-requires="{escape_html(req_val)}"
      data-impact="{escape_html(impact_val)}"
      data-status="{escape_html(status)}"
      data-description="{escape_html(s.get('description',''))}"
      data-notes="{escape_html(notes)}">
      <td class="col-name">{name_html}</td>
      <td class="col-category">{escape_html(s.get('category',''))}</td>
      <td class="col-desc">{escape_html(s.get('description',''))}</td>
      <td class="col-default">{badge(default_val, classify_default(default_val))}</td>
      <td class="col-rec">{badge(rec_val, classify_recommended(rec_val))}</td>
      <td class="col-req">{badge(req_val, classify_requires(req_val))}</td>
      <td class="col-impact">{badge(impact_val, classify_impact(impact_val))}</td>
      <td class="col-notes">{escape_html(notes)}</td>
    </tr>"""


def get_categories(settings: list[dict]) -> list[str]:
    seen = []
    for s in settings:
        c = s.get("category", "")
        if c and c not in seen:
            seen.append(c)
    return seen


def build_html(settings: list[dict], build_date: str) -> str:
    # Only show settings that have been reviewed by a human
    visible = [s for s in settings if s.get("status") != "pending_review"]

    rows_html = "\n".join(build_row(s) for s in visible)
    categories = get_categories(visible)
    category_options = "\n".join(
        f'            <option value="{escape_html(c)}">{escape_html(c)}</option>'
        for c in categories
    )
    total = len(visible)

    # Serialise settings for the Excel export (JS-side) — pending items excluded
    export_data = []
    for s in visible:
        notes = s.get("notes") or ""
        if s.get("status") == "pending_review":
            notes = "New setting detected — pending human review."
        export_data.append({
            "name":                 s.get("name", ""),
            "category":             s.get("category", ""),
            "description":          s.get("description", ""),
            "default_setting":      s.get("default_setting", ""),
            "recommended_setting":  s.get("recommended_setting", ""),
            "requires_adjustment":  s.get("requires_adjustment", ""),
            "impact":               s.get("impact", ""),
            "notes":                notes,
            "status":               s.get("status", "current"),
        })
    settings_json = json.dumps(export_data, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <title>Microsoft Fabric Tenant Settings &mdash; Kristina Bachov&aacute;</title>

  <!-- Playfair Display + Source Sans 3 via Google Fonts (matches kristinabachova.com) -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&family=Source+Sans+3:wght@300;400;500;600;700&display=swap"
        rel="stylesheet">

  <style>
    /* ===== Design tokens — mirrors kristinabachova.com/styles.css ===== */
    :root {{
      --colour-primary:    #2A5F8F;
      --colour-dark:       #1A3D5C;
      --colour-mid:        #5B8DB8;
      --colour-bg:         #F4F1EB;
      --colour-card:       #FFFFFF;
      --colour-text:       #2A3A4A;
      --colour-text-muted: #6A7E90;
      --colour-border:     #D9D4C8;
      --colour-green:      #16A34A;
      --colour-warning:    #D4901A;
      --colour-red:        #DC2626;

      --font-heading: 'Playfair Display', Georgia, serif;
      --font-body:    'Source Sans 3', 'Source Sans Pro', system-ui, sans-serif;

      --radius:      8px;
      --radius-pill: 100px;
      --shadow-sm:   0 1px 3px rgba(0,0,0,0.06);
      --shadow-md:   0 4px 12px rgba(0,0,0,0.08);
      --transition:  0.25s ease;
    }}

    /* ===== Reset ===== */
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    html {{ font-size: 14px; scroll-behavior: smooth; }}
    body {{
      font-family: var(--font-body);
      background: var(--colour-bg);
      color: var(--colour-text);
      line-height: 1.6;
      -webkit-font-smoothing: antialiased;
    }}
    a {{ color: var(--colour-primary); text-decoration: none; transition: color var(--transition); }}
    a:hover {{ color: var(--colour-mid); }}
    img {{ max-width: 100%; display: block; }}

    /* ===== Layout ===== */
    .container {{
      max-width: 1120px;
      margin: 0 auto;
      padding: 0 24px;
    }}

    /* ===== Page header — matches .page-header in styles.css ===== */
    .page-header {{
      background: var(--colour-dark);
      padding: 48px 0 40px;
    }}
    .page-header__title {{
      font-family: var(--font-heading);
      font-size: clamp(1.8rem, 4vw, 2.6rem);
      color: #fff;
      font-weight: 700;
      line-height: 1.25;
      margin-bottom: 12px;
    }}
    .page-header__intro {{
      color: rgba(255,255,255,0.7);
      font-size: 1rem;
      max-width: 640px;
      line-height: 1.7;
    }}
    .page-header__meta {{
      display: inline-block;
      margin-top: 16px;
      background: rgba(255,255,255,0.1);
      color: rgba(255,255,255,0.85);
      font-size: 0.8rem;
      font-weight: 600;
      padding: 5px 14px;
      border-radius: var(--radius-pill);
      letter-spacing: 0.02em;
    }}

    /* ===== Main content ===== */
    .page-body {{
      padding: 32px 0 64px;
    }}

    /* ===== Controls bar ===== */
    .controls {{
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      align-items: center;
      margin-bottom: 20px;
    }}
    .search-box {{
      flex: 1 1 260px;
      position: relative;
    }}
    .search-box input {{
      width: 100%;
      padding: 9px 12px 9px 38px;
      border: 1px solid var(--colour-border);
      border-radius: var(--radius);
      font-family: var(--font-body);
      font-size: 0.9rem;
      background: var(--colour-card);
      color: var(--colour-text);
      transition: border-color var(--transition), box-shadow var(--transition);
    }}
    .search-box input:focus {{
      outline: none;
      border-color: var(--colour-primary);
      box-shadow: 0 0 0 3px rgba(42,95,143,0.12);
    }}
    .search-icon {{
      position: absolute;
      left: 11px;
      top: 50%;
      transform: translateY(-50%);
      color: var(--colour-text-muted);
      pointer-events: none;
    }}

    .filter-group {{
      display: flex;
      align-items: center;
      gap: 8px;
    }}
    .filter-group label {{
      font-size: 0.8rem;
      font-weight: 600;
      color: var(--colour-text-muted);
      white-space: nowrap;
    }}
    .filter-group select {{
      font-family: var(--font-body);
      font-size: 0.85rem;
      padding: 8px 32px 8px 12px;
      border: 1px solid var(--colour-border);
      border-radius: var(--radius);
      background: var(--colour-card);
      color: var(--colour-text);
      appearance: none;
      background-image: url("data:image/svg+xml,%3Csvg width='12' height='8' viewBox='0 0 12 8' fill='none' xmlns='http://www.w3.org/2000/svg'%3E%3Cpath d='M1 1.5L6 6.5L11 1.5' stroke='%236A7E90' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E");
      background-repeat: no-repeat;
      background-position: right 10px center;
      transition: border-color var(--transition);
      cursor: pointer;
    }}
    .filter-group select:focus {{
      outline: none;
      border-color: var(--colour-primary);
    }}

    /* Excel button — matches .btn .btn--primary in styles.css */
    .btn-excel {{
      display: inline-flex;
      align-items: center;
      gap: 7px;
      font-family: var(--font-body);
      font-size: 0.9rem;
      font-weight: 600;
      padding: 9px 22px;
      background: var(--colour-primary);
      color: #fff;
      border: none;
      border-radius: var(--radius-pill);
      cursor: pointer;
      white-space: nowrap;
      transition: background var(--transition), transform var(--transition), box-shadow var(--transition);
    }}
    .btn-excel:hover {{
      background: var(--colour-mid);
      transform: translateY(-1px);
      box-shadow: var(--shadow-md);
    }}

    .row-count {{
      font-size: 0.8rem;
      color: var(--colour-text-muted);
      white-space: nowrap;
      margin-left: auto;
    }}

    /* ===== Table wrapper ===== */
    .table-wrapper {{
      overflow-x: auto;
      border-radius: var(--radius);
      box-shadow: var(--shadow-sm);
      background: var(--colour-card);
      border: 1px solid var(--colour-border);
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 0.85rem;
    }}

    /* ===== Sticky header ===== */
    thead th {{
      background: var(--colour-dark);
      color: #fff;
      padding: 12px 14px;
      text-align: left;
      font-family: var(--font-body);
      font-weight: 600;
      font-size: 0.78rem;
      text-transform: uppercase;
      letter-spacing: 0.04em;
      white-space: nowrap;
      position: sticky;
      top: 0;
      z-index: 10;
      cursor: pointer;
      user-select: none;
      border-right: 1px solid rgba(255,255,255,0.1);
    }}
    thead th:last-child {{ border-right: none; }}
    thead th:hover {{ background: #1e4a6e; }}
    thead th .sort-icon {{ margin-left: 4px; opacity: 0.4; font-size: 0.65rem; }}
    thead th.sort-asc  .sort-icon::after {{ content: " \25B2"; opacity: 1; }}
    thead th.sort-desc .sort-icon::after {{ content: " \25BC"; opacity: 1; }}

    /* ===== Rows ===== */
    tbody tr {{
      border-bottom: 1px solid var(--colour-border);
      transition: background var(--transition);
    }}
    tbody tr:hover {{ background: rgba(42,95,143,0.04); }}
    tbody tr:last-child {{ border-bottom: none; }}
    tbody td {{
      padding: 11px 14px;
      vertical-align: top;
      line-height: 1.55;
    }}

    /* ===== Column widths ===== */
    .col-name     {{ min-width: 200px; max-width: 260px; font-weight: 600; color: var(--colour-dark); font-size: 0.88rem; }}
    .col-category {{ min-width: 130px; max-width: 170px; color: var(--colour-text-muted); font-size: 0.82rem; }}
    .col-desc     {{ min-width: 190px; max-width: 280px; color: var(--colour-text-muted); font-size: 0.82rem; }}
    .col-default  {{ min-width: 130px; max-width: 190px; }}
    .col-rec      {{ min-width: 150px; max-width: 220px; }}
    .col-req      {{ min-width: 88px;  max-width: 110px; text-align: center; }}
    .col-impact   {{ min-width: 78px;  max-width: 100px; text-align: center; }}
    .col-notes    {{ min-width: 210px; color: var(--colour-text-muted); font-size: 0.82rem; }}

    /* ===== Badges — matches .badge in styles.css, extended with semantic colours ===== */
    .badge {{
      display: inline-block;
      font-size: 0.75rem;
      font-weight: 600;
      padding: 3px 10px;
      border-radius: var(--radius-pill);
      line-height: 1.5;
      white-space: normal;
      word-break: break-word;
    }}
    .badge--green  {{ background: rgba(22,163,74,0.12);  color: #14532D; }}
    .badge--red    {{ background: rgba(220,38,38,0.10);  color: #7F1D1D; }}
    .badge--amber  {{ background: rgba(212,144,26,0.13); color: #78350F; }}
    .badge--empty  {{ background: rgba(0,0,0,0.04); color: var(--colour-text-muted); }}
    .badge--pending {{
      background: rgba(212,144,26,0.13);
      color: #78350F;
      font-size: 0.7rem;
      margin-left: 6px;
      vertical-align: middle;
    }}

    /* ===== Removed rows ===== */
    .row-removed td {{
      opacity: 0.4;
      text-decoration: line-through;
    }}

    /* ===== Pending rows ===== */
    .row-pending td {{ background: rgba(212,144,26,0.04); }}

    /* ===== No results ===== */
    .no-results {{
      text-align: center;
      padding: 56px 16px;
      color: var(--colour-text-muted);
      font-size: 0.95rem;
    }}

    /* ===== Footer — matches .footer in styles.css ===== */
    .page-footer {{
      background: var(--colour-dark);
      color: rgba(255,255,255,0.6);
      padding: 40px 0;
      margin-top: 48px;
      font-size: 0.82rem;
    }}
    .page-footer__inner {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px 32px;
      justify-content: space-between;
      align-items: center;
    }}
    .page-footer a {{ color: rgba(255,255,255,0.7); transition: color var(--transition); }}
    .page-footer a:hover {{ color: var(--colour-primary); }}
    .page-footer__right {{ color: rgba(255,255,255,0.4); font-size: 0.78rem; }}

    /* ===== Responsive ===== */
    @media (max-width: 768px) {{
      .page-header {{ padding: 32px 0 28px; }}
      .controls {{ flex-direction: column; align-items: stretch; }}
      .btn-excel {{ justify-content: center; }}
      .row-count {{ margin-left: 0; }}
      .filter-group {{ flex-wrap: wrap; }}
    }}
  </style>
</head>
<body>

<!-- Page header -->
<header class="page-header">
  <div class="container">
    <h1 class="page-header__title">Microsoft Fabric Tenant Settings</h1>
    <p class="page-header__intro">
      Governance reference &mdash; recommended settings with rationale
      for enterprise Power BI and Fabric deployments.
    </p>
    <span class="page-header__meta">Last updated: {build_date}</span>
  </div>
</header>

<!-- Main content -->
<main class="page-body">
  <div class="container">

    <!-- Controls -->
    <div class="controls">
      <div class="search-box">
        <svg class="search-icon" width="16" height="16" viewBox="0 0 20 20"
             fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
          <circle cx="9" cy="9" r="6" stroke="currentColor" stroke-width="2"/>
          <path d="M13.5 13.5L17 17" stroke="currentColor" stroke-width="2"
                stroke-linecap="round"/>
        </svg>
        <input type="text" id="searchInput"
               placeholder="Search setting name, description, or notes&hellip;"
               autocomplete="off" spellcheck="false" aria-label="Search settings">
      </div>

      <div class="filter-group">
        <label for="filterCategory">Category:</label>
        <select id="filterCategory" aria-label="Filter by category">
          <option value="">All categories</option>
{category_options}
        </select>
      </div>

      <div class="filter-group">
        <label for="filterRequires">Requires adjustment:</label>
        <select id="filterRequires" aria-label="Filter by requires adjustment">
          <option value="">All</option>
          <option value="Yes">Yes</option>
          <option value="No">No</option>
        </select>
      </div>

      <div class="filter-group">
        <label for="filterImpact">Impact:</label>
        <select id="filterImpact" aria-label="Filter by impact">
          <option value="">All</option>
          <option value="High">High</option>
          <option value="Medium">Medium</option>
          <option value="Low">Low</option>
        </select>
      </div>

      <button class="btn-excel" onclick="downloadExcel()" aria-label="Download as Excel">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none"
             xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
          <path d="M12 16L7 11H10V4H14V11H17L12 16Z" fill="white"/>
          <path d="M5 18H19V20H5V18Z" fill="white"/>
        </svg>
        Download as Excel
      </button>

      <span class="row-count" id="rowCount" aria-live="polite">
        Showing {total} of {total} settings
      </span>
    </div>

    <!-- Table -->
    <div class="table-wrapper" role="region" aria-label="Fabric tenant settings table">
      <table id="settingsTable">
        <thead>
          <tr>
            <th data-col="0">Setting Name<span class="sort-icon" aria-hidden="true"></span></th>
            <th data-col="1">Category<span class="sort-icon" aria-hidden="true"></span></th>
            <th data-col="2">Description<span class="sort-icon" aria-hidden="true"></span></th>
            <th data-col="3">Default Setting<span class="sort-icon" aria-hidden="true"></span></th>
            <th data-col="4">Recommended Setting<span class="sort-icon" aria-hidden="true"></span></th>
            <th data-col="5">Requires Adjustment<span class="sort-icon" aria-hidden="true"></span></th>
            <th data-col="6">Impact<span class="sort-icon" aria-hidden="true"></span></th>
            <th data-col="7">Governance Notes<span class="sort-icon" aria-hidden="true"></span></th>
          </tr>
        </thead>
        <tbody id="tableBody">
{rows_html}
        </tbody>
      </table>
      <div class="no-results" id="noResults" style="display:none;" role="status">
        No settings match your filters.
      </div>
    </div>

  </div><!-- /container -->
</main>

<!-- Footer -->
<footer class="page-footer">
  <div class="container">
    <div class="page-footer__inner">
      <div>
        Source:
        <a href="{MS_DOCS_URL}" target="_blank" rel="noopener noreferrer">
          Microsoft Fabric documentation
        </a>
        &nbsp;&middot;&nbsp;
        Maintained by
        <a href="{MAINTAINER_URL}" target="_blank" rel="noopener noreferrer">
          Kristina Bachov&aacute;
        </a>
      </div>
      <div class="page-footer__right">
        Updated automatically &middot; Last checked: {build_date}
      </div>
    </div>
  </div>
</footer>

<!-- SheetJS from cdnjs (Excel export) -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"
        integrity="sha512-r22gChDnGvBylk90+2e/ycr3RVrDi8DIOkIGNhJlKfuyQM4tIRAI062MaV8sfjQKYVGjOBaZBOA87z+IhZE9DA=="
        crossorigin="anonymous" referrerpolicy="no-referrer"></script>

<script>
// ============================================================
//  Settings data (for Excel export)
// ============================================================
const SETTINGS_DATA = {settings_json};

// ============================================================
//  Filter + Search
// ============================================================
(function() {{
  const searchInput  = document.getElementById('searchInput');
  const filterCat    = document.getElementById('filterCategory');
  const filterReq    = document.getElementById('filterRequires');
  const filterImpact = document.getElementById('filterImpact');
  const tableBody    = document.getElementById('tableBody');
  const rowCount     = document.getElementById('rowCount');
  const noResults    = document.getElementById('noResults');
  const allRows      = Array.from(tableBody.querySelectorAll('tr'));
  const total        = allRows.length;

  function applyFilters() {{
    var q   = searchInput.value.toLowerCase().trim();
    var cat = filterCat.value;
    var req = filterReq.value;
    var imp = filterImpact.value;
    var visible = 0;

    allRows.forEach(function(row) {{
      var name  = (row.dataset.name        || '').toLowerCase();
      var desc  = (row.dataset.description || '').toLowerCase();
      var notes = (row.dataset.notes       || '').toLowerCase();
      var show  = (!q   || name.includes(q) || desc.includes(q) || notes.includes(q))
               && (!cat || row.dataset.category === cat)
               && (!req || row.dataset.requires === req)
               && (!imp || row.dataset.impact   === imp);

      row.style.display = show ? '' : 'none';
      if (show) visible++;
    }});

    rowCount.textContent = 'Showing ' + visible + ' of ' + total + ' settings';
    noResults.style.display = visible === 0 ? '' : 'none';
  }}

  searchInput.addEventListener('input',   applyFilters);
  filterCat.addEventListener('change',    applyFilters);
  filterReq.addEventListener('change',    applyFilters);
  filterImpact.addEventListener('change', applyFilters);
}})();

// ============================================================
//  Column sort
// ============================================================
(function() {{
  var table   = document.getElementById('settingsTable');
  var headers = table.querySelectorAll('thead th');
  var sortCol = -1, sortAsc = true;

  headers.forEach(function(th, colIdx) {{
    th.addEventListener('click', function() {{
      sortAsc = (sortCol === colIdx) ? !sortAsc : true;
      sortCol = colIdx;
      headers.forEach(function(h) {{ h.classList.remove('sort-asc','sort-desc'); }});
      th.classList.add(sortAsc ? 'sort-asc' : 'sort-desc');

      var tbody = table.querySelector('tbody');
      var rows  = Array.from(tbody.querySelectorAll('tr'));
      rows.sort(function(a, b) {{
        var av = (a.cells[colIdx] ? a.cells[colIdx].textContent : '').trim().toLowerCase();
        var bv = (b.cells[colIdx] ? b.cells[colIdx].textContent : '').trim().toLowerCase();
        return sortAsc ? av.localeCompare(bv) : bv.localeCompare(av);
      }});
      rows.forEach(function(r) {{ tbody.appendChild(r); }});
    }});
  }});
}})();

// ============================================================
//  Excel export using SheetJS
// ============================================================
function downloadExcel() {{
  var wb = XLSX.utils.book_new();

  // Colour helpers — match badge colours from styles.css
  function classifyDefault(v) {{
    if (!v) return 'amber';
    var lv = v.toLowerCase();
    if (lv.startsWith('off')) return 'red';
    if (lv === 'on' || lv.startsWith('on \u2013 all') || lv.startsWith('on \u2014 all')) return 'green';
    return 'amber';
  }}
  function classifyRec(v) {{
    if (!v) return 'amber';
    var lv = v.toLowerCase();
    return lv.startsWith('off') ? 'red' : lv.startsWith('on') ? 'green' : 'amber';
  }}
  function classifyReq(v) {{
    return !v ? 'amber' : (v.trim().toLowerCase() === 'yes' ? 'red' : 'green');
  }}
  function classifyImpact(v) {{
    if (!v) return 'amber';
    var lv = v.trim().toLowerCase();
    return lv === 'high' ? 'red' : lv === 'medium' ? 'amber' : 'green';
  }}

  // Cell fill colours aligned to website palette
  var FILLS = {{
    green:   {{ fgColor: {{ rgb: 'D1FAE5' }} }},  // rgba(22,163,74,0.12) solid approx
    red:     {{ fgColor: {{ rgb: 'FEE2E2' }} }},  // rgba(220,38,38,0.10) solid approx
    amber:   {{ fgColor: {{ rgb: 'FEF3C7' }} }},  // rgba(212,144,26,0.13) solid approx
    header:  {{ fgColor: {{ rgb: '1A3D5C' }} }},  // --colour-dark
    removed: {{ fgColor: {{ rgb: 'F0EDEB' }} }},
    pending: {{ fgColor: {{ rgb: 'FFFBEB' }} }},
    white:   {{ fgColor: {{ rgb: 'FFFFFF' }} }},
  }};

  var HEADER_FONT = {{ bold: true, color: {{ rgb: 'FFFFFF' }}, name: 'Calibri', sz: 11 }};
  var BORDER = {{ style: 'thin', color: {{ rgb: 'D9D4C8' }} }};  // --colour-border
  var CELL_BORDER = {{ top: BORDER, bottom: BORDER, left: BORDER, right: BORDER }};

  // ---- Sheet 1: Settings ----
  var headers = [
    'Setting Name','Category','Description',
    'Default Setting','Recommended Setting',
    'Requires Adjustment','Impact','Governance Notes','Status'
  ];

  var ws = {{}};
  var range = {{ s: {{ r:0, c:0 }}, e: {{ r:SETTINGS_DATA.length, c:headers.length-1 }} }};

  headers.forEach(function(h, c) {{
    var ref = XLSX.utils.encode_cell({{ r:0, c:c }});
    ws[ref] = {{
      v: h, t: 's',
      s: {{
        font: HEADER_FONT,
        fill: FILLS.header,
        alignment: {{ horizontal:'center', vertical:'center', wrapText:true }},
        border: CELL_BORDER
      }}
    }};
  }});

  SETTINGS_DATA.forEach(function(row, ri) {{
    var r = ri + 1;
    var status = row.status || 'current';
    var rowFill = status === 'removed' ? FILLS.removed
                : status === 'pending_review' ? FILLS.pending
                : null;

    var colFns = [null, null, null,
      classifyDefault(row.default_setting),
      classifyRec(row.recommended_setting),
      classifyReq(row.requires_adjustment),
      classifyImpact(row.impact),
      null, null
    ];
    var vals = [
      row.name, row.category, row.description,
      row.default_setting, row.recommended_setting,
      row.requires_adjustment, row.impact,
      row.notes, row.status
    ];

    vals.forEach(function(val, c) {{
      var ref = XLSX.utils.encode_cell({{ r:r, c:c }});
      var fill = colFns[c] ? FILLS[colFns[c]] : (rowFill || FILLS.white);
      ws[ref] = {{
        v: val || '', t: 's',
        s: {{
          fill: fill,
          font: {{ name:'Calibri', sz:10 }},
          alignment: {{ vertical:'top', wrapText:true }},
          border: CELL_BORDER
        }}
      }};
    }});
  }});

  ws['!ref']        = XLSX.utils.encode_range(range);
  ws['!autofilter'] = {{ ref: XLSX.utils.encode_range(range) }};
  ws['!cols'] = [
    {{ wch:38 }},  // Setting Name
    {{ wch:22 }},  // Category
    {{ wch:42 }},  // Description
    {{ wch:32 }},  // Default
    {{ wch:36 }},  // Recommended
    {{ wch:12 }},  // Requires Adj
    {{ wch:10 }},  // Impact
    {{ wch:60 }},  // Notes
    {{ wch:14 }},  // Status
  ];
  ws['!rows'] = [{{ hpt:28 }}];
  XLSX.utils.book_append_sheet(wb, ws, 'Fabric Tenant Settings');

  // ---- Sheet 2: Legend ----
  var legendRows = [
    ['Column','Value','Cell colour','Meaning'],
    ['Default Setting','On','Green','Enabled for all users by default'],
    ['Default Setting','On (scoped / conditional)','Amber','Enabled with limitations or conditions'],
    ['Default Setting','Off','Red','Disabled by default'],
    ['Recommended Setting','On \u2026','Green','Recommended to enable'],
    ['Recommended Setting','Off \u2026','Red','Recommended to keep disabled'],
    ['Requires Adjustment','Yes','Red','Action required — default does not match recommendation'],
    ['Requires Adjustment','No','Green','No immediate action required'],
    ['Impact','High','Red','High governance / security impact'],
    ['Impact','Medium','Amber','Moderate governance impact'],
    ['Impact','Low','Green','Low governance impact'],
    ['Status','current','White','Setting is current and reviewed'],
    ['Status','pending_review','Amber (light)','New setting detected — needs human review'],
    ['Status','removed','Grey (light)','No longer present in MS documentation'],
  ];
  var legendFillMap = {{
    'Green':         FILLS.green,
    'Amber':         FILLS.amber,
    'Amber (light)': FILLS.pending,
    'Red':           FILLS.red,
    'White':         FILLS.white,
    'Grey (light)':  FILLS.removed,
  }};

  var wsL = {{}};
  legendRows.forEach(function(row, ri) {{
    row.forEach(function(val, ci) {{
      var ref = XLSX.utils.encode_cell({{ r:ri, c:ci }});
      var isHeader = ri === 0;
      var colourLabel = (ri > 0 && ci === 2) ? val : null;
      var fill = isHeader ? FILLS.header
               : (colourLabel && legendFillMap[colourLabel]) ? legendFillMap[colourLabel]
               : FILLS.white;
      wsL[ref] = {{
        v: val, t: 's',
        s: {{
          font: isHeader ? HEADER_FONT : {{ name:'Calibri', sz:10 }},
          fill: fill,
          alignment: {{ vertical:'center', wrapText:true }},
          border: CELL_BORDER
        }}
      }};
    }});
  }});
  wsL['!ref']  = XLSX.utils.encode_range({{ s:{{r:0,c:0}}, e:{{r:legendRows.length-1,c:3}} }});
  wsL['!cols'] = [{{ wch:22 }},{{ wch:32 }},{{ wch:14 }},{{ wch:55 }}];
  XLSX.utils.book_append_sheet(wb, wsL, 'Legend');

  // ---- Download ----
  var pad = function(n) {{ return String(n).padStart(2,'0'); }};
  var d   = new Date();
  var ds  = d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());
  XLSX.writeFile(wb, 'fabric-tenant-settings-'+ds+'.xlsx');
}}
</script>

</body>
</html>"""


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    settings = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    build_date = date.today().isoformat()
    print(f"Loaded {len(settings)} settings from {SETTINGS_FILE.relative_to(REPO_ROOT)}")

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    html = build_html(settings, build_date)
    OUTPUT_FILE.write_text(html, encoding="utf-8")

    size_kb = OUTPUT_FILE.stat().st_size / 1024
    print(f"Generated {OUTPUT_FILE.relative_to(REPO_ROOT)} ({size_kb:.1f} KB)")


if __name__ == "__main__":
    main()
