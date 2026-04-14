"""
build.py — Generate docs/index.html from data/settings.json.

The output is a fully self-contained HTML file (inline CSS + JS) that:
  • Shows all Fabric tenant settings in a searchable, filterable table
  • Colour-codes badges by value category (green/red/amber)
  • Supports Excel download via SheetJS
  • Works as a standalone local file AND via GitHub Pages
  • Is safely embeddable in an iframe on an external site
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
# Badge colour classification helpers
# ---------------------------------------------------------------------------

def classify_default(value: str) -> str:
    """Return 'green', 'red', or 'amber' for a Default Setting value."""
    if not value:
        return "amber"
    v = value.lower()
    if v.startswith("off"):
        return "red"
    if v.startswith("on – all") or v == "on":
        return "green"
    # Partial / conditional (e.g. "On – scoped to …", "On (if same geo)")
    return "amber"


def classify_recommended(value: str) -> str:
    """Return badge colour for a Recommended Setting value."""
    if not value:
        return "amber"
    v = value.lower()
    if v.startswith("off"):
        return "red"
    if v.startswith("on"):
        return "green"
    return "amber"


def classify_requires(value: str) -> str:
    """'Yes' → red, 'No' → green."""
    if not value:
        return "amber"
    return "red" if value.strip().lower() == "yes" else "green"


def classify_impact(value: str) -> str:
    """'High' → red, 'Medium' → amber, 'Low' → green."""
    if not value:
        return "amber"
    v = value.strip().lower()
    if v == "high":
        return "red"
    if v == "medium":
        return "amber"
    return "green"


# ---------------------------------------------------------------------------
# HTML generation
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
    """Build a single <tr> for one setting."""
    status = s.get("status", "current")
    is_removed = status == "removed"
    is_pending = status == "pending_review"

    row_class = ""
    if is_removed:
        row_class = ' class="row-removed"'
    elif is_pending:
        row_class = ' class="row-pending"'

    # Pending notes override
    notes = s.get("notes") or ""
    if is_pending:
        notes = "⚠️ New setting detected — pending human review."

    name_html = escape_html(s.get("name", ""))
    if is_pending:
        name_html += ' <span class="badge badge-pending">⚠ New</span>'

    default_val = s.get("default_setting") or ""
    rec_val = s.get("recommended_setting") or ""
    req_val = s.get("requires_adjustment") or ""
    impact_val = s.get("impact") or ""

    def badge(val, colour_class, extra_class=""):
        classes = f"badge badge-{colour_class}"
        if extra_class:
            classes += f" {extra_class}"
        return f'<span class="{classes}">{escape_html(val)}</span>' if val else '<span class="badge badge-empty">—</span>'

    default_badge = badge(default_val, classify_default(default_val))
    rec_badge = badge(rec_val, classify_recommended(rec_val))
    req_badge = badge(req_val, classify_requires(req_val))
    impact_badge = badge(impact_val, classify_impact(impact_val))

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
      <td class="col-default">{default_badge}</td>
      <td class="col-rec">{rec_badge}</td>
      <td class="col-req">{req_badge}</td>
      <td class="col-impact">{impact_badge}</td>
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
    rows_html = "\n".join(build_row(s) for s in settings)
    categories = get_categories(settings)
    category_options = "\n".join(
        f'            <option value="{escape_html(c)}">{escape_html(c)}</option>'
        for c in categories
    )
    total = len(settings)

    # Serialise settings as JSON for the Excel export
    # Only include fields needed by Excel
    export_data = []
    for s in settings:
        notes = s.get("notes") or ""
        if s.get("status") == "pending_review":
            notes = "New setting detected — pending human review."
        export_data.append({
            "name": s.get("name", ""),
            "category": s.get("category", ""),
            "description": s.get("description", ""),
            "default_setting": s.get("default_setting", ""),
            "recommended_setting": s.get("recommended_setting", ""),
            "requires_adjustment": s.get("requires_adjustment", ""),
            "impact": s.get("impact", ""),
            "notes": notes,
            "status": s.get("status", "current"),
        })
    settings_json = json.dumps(export_data, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <title>Microsoft Fabric Tenant Settings</title>
  <style>
    /* ===== Reset & Base ===== */
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    html {{ font-size: 14px; }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto,
                   "Helvetica Neue", Arial, sans-serif;
      background: #f8f9fa;
      color: #212529;
      line-height: 1.5;
    }}

    /* ===== Layout ===== */
    .page-wrapper {{
      max-width: 1400px;
      margin: 0 auto;
      padding: 0 16px 40px;
    }}

    /* ===== Header ===== */
    .page-header {{
      background: #1F3864;
      color: #fff;
      padding: 24px 32px 20px;
      margin-bottom: 24px;
    }}
    .page-header h1 {{
      font-size: 1.6rem;
      font-weight: 700;
      letter-spacing: -0.3px;
    }}
    .page-header .subtitle {{
      opacity: 0.85;
      font-size: 0.9rem;
      margin-top: 4px;
    }}
    .last-updated {{
      background: #fff;
      color: #1F3864;
      display: inline-block;
      padding: 4px 12px;
      border-radius: 4px;
      font-size: 0.8rem;
      font-weight: 600;
      margin-top: 12px;
    }}

    /* ===== Controls bar ===== */
    .controls {{
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      align-items: center;
      margin-bottom: 16px;
    }}
    .search-box {{
      flex: 1 1 260px;
      position: relative;
    }}
    .search-box input {{
      width: 100%;
      padding: 8px 12px 8px 36px;
      border: 1px solid #ced4da;
      border-radius: 6px;
      font-size: 0.9rem;
      background: #fff;
    }}
    .search-box input:focus {{ outline: none; border-color: #1F3864; box-shadow: 0 0 0 2px rgba(31,56,100,.15); }}
    .search-icon {{
      position: absolute;
      left: 10px;
      top: 50%;
      transform: translateY(-50%);
      color: #6c757d;
      pointer-events: none;
    }}
    .filter-group {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      align-items: center;
    }}
    .filter-group label {{
      font-size: 0.8rem;
      font-weight: 600;
      color: #495057;
      white-space: nowrap;
    }}
    .filter-group select {{
      padding: 7px 28px 7px 10px;
      border: 1px solid #ced4da;
      border-radius: 6px;
      font-size: 0.85rem;
      background: #fff;
      appearance: none;
      background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%236c757d'/%3E%3C/svg%3E");
      background-repeat: no-repeat;
      background-position: right 8px center;
      cursor: pointer;
    }}
    .filter-group select:focus {{ outline: none; border-color: #1F3864; }}

    .btn-excel {{
      background: #1F7A4E;
      color: #fff;
      border: none;
      padding: 8px 18px;
      border-radius: 6px;
      font-size: 0.875rem;
      font-weight: 600;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 6px;
      white-space: nowrap;
      transition: background 0.15s;
    }}
    .btn-excel:hover {{ background: #155c3a; }}

    .row-count {{
      font-size: 0.8rem;
      color: #6c757d;
      white-space: nowrap;
      margin-left: auto;
    }}

    /* ===== Table wrapper ===== */
    .table-wrapper {{
      overflow-x: auto;
      border-radius: 8px;
      box-shadow: 0 1px 4px rgba(0,0,0,.08);
      background: #fff;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 0.85rem;
    }}

    /* ===== Sticky header ===== */
    thead th {{
      background: #1F3864;
      color: #fff;
      padding: 11px 12px;
      text-align: left;
      font-weight: 600;
      font-size: 0.8rem;
      white-space: nowrap;
      position: sticky;
      top: 0;
      z-index: 10;
      cursor: pointer;
      user-select: none;
      border-right: 1px solid rgba(255,255,255,.12);
    }}
    thead th:last-child {{ border-right: none; }}
    thead th:hover {{ background: #2a4d87; }}
    thead th .sort-icon {{ margin-left: 4px; opacity: 0.5; font-size: 0.7rem; }}
    thead th.sort-asc .sort-icon::after {{ content: " ▲"; opacity: 1; }}
    thead th.sort-desc .sort-icon::after {{ content: " ▼"; opacity: 1; }}

    /* ===== Rows ===== */
    tbody tr {{
      border-bottom: 1px solid #e9ecef;
      transition: background 0.1s;
    }}
    tbody tr:hover {{ background: #f0f4fc; }}
    tbody tr:last-child {{ border-bottom: none; }}
    tbody td {{
      padding: 10px 12px;
      vertical-align: top;
    }}

    /* ===== Column widths ===== */
    .col-name   {{ min-width: 220px; max-width: 280px; font-weight: 600; color: #1F3864; }}
    .col-category {{ min-width: 140px; max-width: 180px; color: #495057; }}
    .col-desc   {{ min-width: 200px; max-width: 300px; color: #495057; font-size: 0.82rem; }}
    .col-default {{ min-width: 130px; max-width: 190px; }}
    .col-rec    {{ min-width: 150px; max-width: 220px; }}
    .col-req    {{ min-width: 90px;  max-width: 110px; text-align: center; }}
    .col-impact {{ min-width: 80px;  max-width: 100px; text-align: center; }}
    .col-notes  {{ min-width: 220px; color: #495057; font-size: 0.82rem; }}

    /* ===== Badges ===== */
    .badge {{
      display: inline-block;
      padding: 2px 8px;
      border-radius: 4px;
      font-size: 0.78rem;
      font-weight: 600;
      line-height: 1.6;
      white-space: normal;
      word-break: break-word;
    }}
    .badge-green  {{ background: #d1f5de; color: #155724; }}
    .badge-red    {{ background: #fde0e0; color: #721c24; }}
    .badge-amber  {{ background: #fff3cd; color: #856404; }}
    .badge-empty  {{ background: #e9ecef; color: #6c757d; }}
    .badge-pending {{ background: #fff3cd; color: #856404; font-size: 0.72rem; }}

    /* ===== Removed rows ===== */
    .row-removed td {{ opacity: 0.45; text-decoration: line-through; }}
    .row-removed:hover {{ background: #fff5f5; }}

    /* ===== Pending rows ===== */
    .row-pending td {{ background: #fffbeb; }}
    .row-pending:hover {{ background: #fef9e7; }}

    /* ===== No results ===== */
    .no-results {{
      text-align: center;
      padding: 48px 16px;
      color: #6c757d;
      font-size: 0.95rem;
    }}

    /* ===== Footer ===== */
    .page-footer {{
      margin-top: 32px;
      padding: 20px 0;
      border-top: 1px solid #dee2e6;
      font-size: 0.78rem;
      color: #6c757d;
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      justify-content: space-between;
      align-items: center;
    }}
    .page-footer a {{ color: #1F3864; text-decoration: none; }}
    .page-footer a:hover {{ text-decoration: underline; }}
    .footer-right {{ text-align: right; }}

    /* ===== Responsive ===== */
    @media (max-width: 768px) {{
      .page-header {{ padding: 16px; }}
      .page-header h1 {{ font-size: 1.2rem; }}
      .controls {{ flex-direction: column; align-items: stretch; }}
      .btn-excel {{ width: 100%; justify-content: center; }}
      .row-count {{ margin-left: 0; }}
    }}
  </style>
</head>
<body>

<div class="page-header">
  <div class="page-wrapper" style="padding-top:0; padding-bottom:0;">
    <h1>Microsoft Fabric Tenant Settings</h1>
    <p class="subtitle">
      Governance reference — recommended settings with rationale for enterprise deployments
    </p>
    <span class="last-updated">Last updated: {build_date}</span>
  </div>
</div>

<div class="page-wrapper">

  <!-- Controls -->
  <div class="controls">
    <div class="search-box">
      <svg class="search-icon" width="16" height="16" viewBox="0 0 20 20" fill="none"
           xmlns="http://www.w3.org/2000/svg">
        <circle cx="9" cy="9" r="6" stroke="currentColor" stroke-width="2"/>
        <path d="M13.5 13.5L17 17" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
      </svg>
      <input type="text" id="searchInput"
             placeholder="Search setting name, description, or notes…"
             autocomplete="off" spellcheck="false">
    </div>

    <div class="filter-group">
      <label for="filterCategory">Category:</label>
      <select id="filterCategory">
        <option value="">All categories</option>
{category_options}
      </select>
    </div>

    <div class="filter-group">
      <label for="filterRequires">Requires adjustment:</label>
      <select id="filterRequires">
        <option value="">All</option>
        <option value="Yes">Yes</option>
        <option value="No">No</option>
      </select>
    </div>

    <div class="filter-group">
      <label for="filterImpact">Impact:</label>
      <select id="filterImpact">
        <option value="">All</option>
        <option value="High">High</option>
        <option value="Medium">Medium</option>
        <option value="Low">Low</option>
      </select>
    </div>

    <button class="btn-excel" onclick="downloadExcel()">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none"
           xmlns="http://www.w3.org/2000/svg">
        <path d="M12 16L7 11H10V4H14V11H17L12 16Z" fill="white"/>
        <path d="M5 18H19V20H5V18Z" fill="white"/>
      </svg>
      Download as Excel
    </button>

    <span class="row-count" id="rowCount">Showing {total} of {total} settings</span>
  </div>

  <!-- Table -->
  <div class="table-wrapper">
    <table id="settingsTable" aria-label="Microsoft Fabric Tenant Settings">
      <thead>
        <tr>
          <th data-col="0">Setting Name<span class="sort-icon"></span></th>
          <th data-col="1">Category<span class="sort-icon"></span></th>
          <th data-col="2">Description<span class="sort-icon"></span></th>
          <th data-col="3">Default Setting<span class="sort-icon"></span></th>
          <th data-col="4">Recommended Setting<span class="sort-icon"></span></th>
          <th data-col="5">Requires Adjustment<span class="sort-icon"></span></th>
          <th data-col="6">Impact<span class="sort-icon"></span></th>
          <th data-col="7">Governance Notes<span class="sort-icon"></span></th>
        </tr>
      </thead>
      <tbody id="tableBody">
{rows_html}
      </tbody>
    </table>
    <div class="no-results" id="noResults" style="display:none;">
      No settings match your filters.
    </div>
  </div>

  <!-- Footer -->
  <div class="page-footer">
    <div>
      Source:
      <a href="{MS_DOCS_URL}" target="_blank" rel="noopener noreferrer">
        Microsoft Fabric documentation
      </a>
      &nbsp;|&nbsp;
      Maintained by
      <a href="{MAINTAINER_URL}" target="_blank" rel="noopener noreferrer">
        Kristina Bachov&aacute; | kristinabachova.com
      </a>
    </div>
    <div class="footer-right">
      This tool is updated automatically. Last checked: {build_date}
    </div>
  </div>

</div><!-- /page-wrapper -->

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
    const q       = searchInput.value.toLowerCase().trim();
    const cat     = filterCat.value;
    const req     = filterReq.value;
    const imp     = filterImpact.value;
    let visible   = 0;

    allRows.forEach(function(row) {{
      const name     = (row.dataset.name        || '').toLowerCase();
      const desc     = (row.dataset.description || '').toLowerCase();
      const notes    = (row.dataset.notes       || '').toLowerCase();
      const rowCat   = row.dataset.category   || '';
      const rowReq   = row.dataset.requires   || '';
      const rowImp   = row.dataset.impact     || '';

      const matchSearch = !q || name.includes(q) || desc.includes(q) || notes.includes(q);
      const matchCat    = !cat || rowCat === cat;
      const matchReq    = !req || rowReq === req;
      const matchImp    = !imp || rowImp === imp;

      const show = matchSearch && matchCat && matchReq && matchImp;
      row.style.display = show ? '' : 'none';
      if (show) visible++;
    }});

    rowCount.textContent = 'Showing ' + visible + ' of ' + total + ' settings';
    noResults.style.display = visible === 0 ? '' : 'none';
  }}

  searchInput.addEventListener('input',  applyFilters);
  filterCat.addEventListener('change',   applyFilters);
  filterReq.addEventListener('change',   applyFilters);
  filterImpact.addEventListener('change', applyFilters);
}})();

// ============================================================
//  Column sort
// ============================================================
(function() {{
  const table = document.getElementById('settingsTable');
  const headers = table.querySelectorAll('thead th');
  let sortCol = -1, sortAsc = true;

  headers.forEach(function(th, colIdx) {{
    th.addEventListener('click', function() {{
      if (sortCol === colIdx) {{
        sortAsc = !sortAsc;
      }} else {{
        sortCol = colIdx;
        sortAsc = true;
      }}
      headers.forEach(function(h) {{
        h.classList.remove('sort-asc', 'sort-desc');
      }});
      th.classList.add(sortAsc ? 'sort-asc' : 'sort-desc');

      const tbody = table.querySelector('tbody');
      const rows  = Array.from(tbody.querySelectorAll('tr'));

      rows.sort(function(a, b) {{
        const aVal = (a.cells[colIdx] ? a.cells[colIdx].textContent : '').trim().toLowerCase();
        const bVal = (b.cells[colIdx] ? b.cells[colIdx].textContent : '').trim().toLowerCase();
        if (aVal < bVal) return sortAsc ? -1 : 1;
        if (aVal > bVal) return sortAsc ?  1 : -1;
        return 0;
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

  // ---- Colour helpers ----
  function classifyDefault(v) {{
    if (!v) return 'amber';
    var lv = v.toLowerCase();
    if (lv.startsWith('off')) return 'red';
    if (lv === 'on' || lv.startsWith('on \u2013 all')) return 'green';
    return 'amber';
  }}
  function classifyRec(v) {{
    if (!v) return 'amber';
    var lv = v.toLowerCase();
    if (lv.startsWith('off')) return 'red';
    if (lv.startsWith('on')) return 'green';
    return 'amber';
  }}
  function classifyReq(v) {{
    if (!v) return 'amber';
    return v.trim().toLowerCase() === 'yes' ? 'red' : 'green';
  }}
  function classifyImpact(v) {{
    if (!v) return 'amber';
    var lv = v.trim().toLowerCase();
    if (lv === 'high')   return 'red';
    if (lv === 'medium') return 'amber';
    return 'green';
  }}

  var COLOURS = {{
    green:  {{ fgColor: {{ rgb: 'D1F5DE' }} }},
    red:    {{ fgColor: {{ rgb: 'FDE0E0' }} }},
    amber:  {{ fgColor: {{ rgb: 'FFF3CD' }} }},
    header: {{ fgColor: {{ rgb: '1F3864' }} }},
    removed:{{ fgColor: {{ rgb: 'E9ECEF' }} }},
    pending:{{ fgColor: {{ rgb: 'FFFBEB' }} }},
  }};

  var HEADER_FONT  = {{ bold: true, color: {{ rgb: 'FFFFFF' }}, name: 'Calibri', sz: 11 }};
  var BORDER_STYLE = {{ style: 'thin', color: {{ rgb: 'D0D7E0' }} }};
  var CELL_BORDER  = {{ top: BORDER_STYLE, bottom: BORDER_STYLE, left: BORDER_STYLE, right: BORDER_STYLE }};

  // ---- Sheet 1: Settings ----
  var headers = [
    'Setting Name', 'Category', 'Description',
    'Default Setting', 'Recommended Setting',
    'Requires Adjustment', 'Impact', 'Governance Notes', 'Status'
  ];

  var ws = {{}};
  var range = {{ s: {{ r:0, c:0 }}, e: {{ r:SETTINGS_DATA.length, c:headers.length-1 }} }};

  // Write header row
  headers.forEach(function(h, c) {{
    var cellRef = XLSX.utils.encode_cell({{ r:0, c:c }});
    ws[cellRef] = {{
      v: h, t: 's',
      s: {{
        font: HEADER_FONT,
        fill: COLOURS.header,
        alignment: {{ horizontal: 'center', vertical: 'center', wrapText: true }},
        border: CELL_BORDER
      }}
    }};
  }});

  // Write data rows
  SETTINGS_DATA.forEach(function(row, ri) {{
    var r = ri + 1;
    var status = row.status || 'current';
    var rowFill = status === 'removed' ? COLOURS.removed
                : status === 'pending_review' ? COLOURS.pending
                : null;

    var colClassifiers = [
      null, null, null,
      classifyDefault(row.default_setting),
      classifyRec(row.recommended_setting),
      classifyReq(row.requires_adjustment),
      classifyImpact(row.impact),
      null, null
    ];
    var values = [
      row.name, row.category, row.description,
      row.default_setting, row.recommended_setting,
      row.requires_adjustment, row.impact,
      row.notes, row.status
    ];

    values.forEach(function(val, c) {{
      var cellRef = XLSX.utils.encode_cell({{ r:r, c:c }});
      var fill = colClassifiers[c] ? COLOURS[colClassifiers[c]] : (rowFill || {{ fgColor: {{ rgb: 'FFFFFF' }} }});
      ws[cellRef] = {{
        v: val || '', t: 's',
        s: {{
          fill: fill,
          font: {{ name: 'Calibri', sz: 10 }},
          alignment: {{ vertical: 'top', wrapText: true }},
          border: CELL_BORDER
        }}
      }};
    }});
  }});

  ws['!ref'] = XLSX.utils.encode_range(range);
  ws['!autofilter'] = {{ ref: XLSX.utils.encode_range(range) }};
  ws['!cols'] = [
    {{ wch: 38 }}, // Setting Name
    {{ wch: 22 }}, // Category
    {{ wch: 42 }}, // Description
    {{ wch: 32 }}, // Default
    {{ wch: 36 }}, // Recommended
    {{ wch: 12 }}, // Requires Adj
    {{ wch: 10 }}, // Impact
    {{ wch: 60 }}, // Notes
    {{ wch: 14 }}, // Status
  ];
  ws['!rows'] = [{{ hpt: 28 }}]; // header row height

  XLSX.utils.book_append_sheet(wb, ws, 'Fabric Tenant Settings');

  // ---- Sheet 2: Legend ----
  var legendRows = [
    ['Column', 'Value', 'Colour', 'Meaning'],
    ['Default Setting', 'On', 'Green', 'Enabled for all users by default'],
    ['Default Setting', 'On (scoped/conditional)', 'Amber', 'Enabled with limitations or conditions'],
    ['Default Setting', 'Off', 'Red', 'Disabled by default'],
    ['Recommended Setting', 'On …', 'Green', 'Recommended to be enabled'],
    ['Recommended Setting', 'Off …', 'Red', 'Recommended to remain disabled'],
    ['Requires Adjustment', 'Yes', 'Red', 'Action required — default does not match recommendation'],
    ['Requires Adjustment', 'No', 'Green', 'No immediate action required'],
    ['Impact', 'High', 'Red', 'High governance/security impact'],
    ['Impact', 'Medium', 'Amber', 'Moderate governance impact'],
    ['Impact', 'Low', 'Green', 'Low governance impact'],
    ['Status', 'current', 'White', 'Setting is current and reviewed'],
    ['Status', 'pending_review', 'Light yellow', 'New setting detected — human review needed'],
    ['Status', 'removed', 'Light grey', 'Setting no longer present in MS documentation'],
  ];

  var legendFillMap = {{
    'Green': COLOURS.green,
    'Red': COLOURS.red,
    'Amber': COLOURS.amber,
    'White': {{ fgColor: {{ rgb: 'FFFFFF' }} }},
    'Light yellow': COLOURS.pending,
    'Light grey': COLOURS.removed,
  }};

  var wsL = {{}};
  legendRows.forEach(function(row, ri) {{
    row.forEach(function(val, ci) {{
      var cellRef = XLSX.utils.encode_cell({{ r:ri, c:ci }});
      var isHeader = ri === 0;
      var colourLabel = ri > 0 && ci === 2 ? val : null;
      var fill = isHeader ? COLOURS.header
               : colourLabel && legendFillMap[colourLabel] ? legendFillMap[colourLabel]
               : {{ fgColor: {{ rgb: 'FFFFFF' }} }};
      wsL[cellRef] = {{
        v: val, t: 's',
        s: {{
          font: isHeader ? HEADER_FONT : {{ name: 'Calibri', sz: 10 }},
          fill: fill,
          alignment: {{ vertical: 'center', wrapText: true }},
          border: CELL_BORDER
        }}
      }};
    }});
  }});
  wsL['!ref'] = XLSX.utils.encode_range({{ s:{{r:0,c:0}}, e:{{r:legendRows.length-1, c:3}} }});
  wsL['!cols'] = [{{ wch: 22 }}, {{ wch: 32 }}, {{ wch: 14 }}, {{ wch: 55 }}];

  XLSX.utils.book_append_sheet(wb, wsL, 'Legend');

  // ---- Download ----
  var today = new Date();
  var pad = function(n) {{ return String(n).padStart(2,'0'); }};
  var dateStr = today.getFullYear() + '-' + pad(today.getMonth()+1) + '-' + pad(today.getDate());
  XLSX.writeFile(wb, 'fabric-tenant-settings-' + dateStr + '.xlsx');
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
