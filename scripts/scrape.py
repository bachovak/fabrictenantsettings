"""
scrape.py — Fetch the Microsoft Fabric tenant settings documentation page,
compare the list of settings against data/settings.json, and report any diff.

Exit codes:
  0 — No changes detected (workflow passes silently)
  1 — Changes detected (workflow fails → GitHub sends notification email)

The diff is printed to stdout (visible in the GitHub Actions log) and also
written to data/last_diff.txt for reference.
"""

import json
import sys
import re
from datetime import date
from pathlib import Path

import requests
from bs4 import BeautifulSoup

# ---------------------------------------------------------------------------
# Paths (relative to the repository root — works locally and in Actions)
# ---------------------------------------------------------------------------
REPO_ROOT = Path(__file__).resolve().parent.parent
SETTINGS_FILE = REPO_ROOT / "data" / "settings.json"
DIFF_FILE = REPO_ROOT / "data" / "last_diff.txt"

MS_DOCS_URL = (
    "https://learn.microsoft.com/en-us/fabric/admin/tenant-settings-index"
)

# H2 headings that are not setting categories — stop parsing when we hit these
STOP_HEADINGS = {"Related content", "Feedback", "Additional resources", "In this article"}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def load_settings() -> list[dict]:
    with open(SETTINGS_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_settings(settings: list[dict]) -> None:
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2, ensure_ascii=False)


def make_slug(name: str) -> str:
    slug = name.lower()
    slug = re.sub(r"[^a-z0-9\s-]", "", slug)
    slug = re.sub(r"\s+", "-", slug.strip())
    slug = re.sub(r"-+", "-", slug)
    return slug[:80]


def fetch_page(url: str) -> str:
    """Download the MS Docs page and return the HTML text."""
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (compatible; FabricSettingsScraper/1.0; "
            "+https://github.com/bachovak/fabrictenantsettings)"
        )
    }
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.text


def parse_settings_from_page(html: str) -> dict[str, str]:
    """
    Extract setting names and their category headings from the MS Docs page.

    Page structure:
      <h2>Category name</h2>
      <table>
        <thead><tr><th>Setting name</th><th>Description</th></tr></thead>
        <tbody>
          <tr><td>Setting Name</td><td>Description text…</td></tr>
          …
        </tbody>
      </table>
      <h2>Next category</h2>
      …

    Returns: dict of {setting_name: category_name}
    """
    soup = BeautifulSoup(html, "lxml")
    article = soup.find("main") or soup.find("article") or soup.body
    scraped: dict[str, str] = {}

    for h2 in article.find_all("h2"):
        heading = h2.get_text(strip=True)
        if heading in STOP_HEADINGS:
            continue

        # The settings table immediately follows the H2 heading
        table = h2.find_next_sibling("table")
        if not table:
            continue

        # Verify this is a settings table (thead contains "Setting" or "Name")
        thead = table.find("thead")
        if thead and "setting" not in thead.get_text().lower():
            continue

        tbody = table.find("tbody")
        if not tbody:
            continue

        for row in tbody.find_all("tr"):
            cells = row.find_all("td")
            if cells:
                name = cells[0].get_text(strip=True)
                if name:
                    scraped[name] = heading

    return scraped


# ---------------------------------------------------------------------------
# Main diff logic
# ---------------------------------------------------------------------------

def main() -> int:
    print(f"Fetching {MS_DOCS_URL} …")
    try:
        html = fetch_page(MS_DOCS_URL)
    except requests.RequestException as exc:
        # Don't exit 1 for a transient network failure — avoid spurious emails
        print(f"ERROR: Failed to fetch documentation page: {exc}", file=sys.stderr)
        sys.exit(0)

    print("Parsing settings from page …")
    scraped = parse_settings_from_page(html)
    print(f"Found {len(scraped)} settings on the page.")

    if not scraped:
        print(
            "WARNING: No settings parsed from page. "
            "The page structure may have changed. Skipping diff.",
            file=sys.stderr,
        )
        sys.exit(0)

    settings = load_settings()

    # Compare by ms_docs_anchor (the exact name from the MS Docs page).
    # Settings already marked "removed" are excluded from comparison — they
    # are kept in the JSON for historical reference only.
    active_settings = [s for s in settings if s.get("status") != "removed"]
    known_anchors = {s["ms_docs_anchor"] for s in active_settings}
    scraped_names = set(scraped.keys())

    new_names = scraped_names - known_anchors       # on page, not in JSON
    removed_names = known_anchors - scraped_names   # in JSON, not on page

    if not new_names and not removed_names:
        print("No changes detected.")
        if DIFF_FILE.exists():
            DIFF_FILE.unlink()
        return 0

    # ------------------------------------------------------------------
    # Build human-readable diff summary
    # ------------------------------------------------------------------
    today = date.today().isoformat()
    lines = [
        f"Fabric Tenant Settings — diff detected on {today}",
        "=" * 60,
        "",
    ]

    if new_names:
        lines.append(
            f"NEW SETTINGS ({len(new_names)}) — "
            "added to settings.json with status 'pending_review':"
        )
        for name in sorted(new_names):
            lines.append(f"  + {name}  [category: {scraped[name]}]")
        lines.append("")

    if removed_names:
        lines.append(
            f"REMOVED SETTINGS ({len(removed_names)}) — "
            "marked as 'removed' in settings.json:"
        )
        for name in sorted(removed_names):
            lines.append(f"  - {name}")
        lines.append("")

    lines.append("ACTION REQUIRED:")
    if new_names:
        lines.append(
            "  • Open data/settings.json and fill in default_setting, "
            "recommended_setting, requires_adjustment, impact, and notes "
            "for each new setting, then change status to 'current'."
        )
    if removed_names:
        lines.append(
            "  • Review settings marked 'removed' and confirm they are "
            "no longer present in the Microsoft documentation."
        )
    lines.append(
        "  • Run scripts/build.py (or trigger workflow_dispatch) to "
        "regenerate docs/index.html once the JSON is updated."
    )

    summary = "\n".join(lines)
    print(summary)

    # ------------------------------------------------------------------
    # Update settings.json
    # ------------------------------------------------------------------

    # Add new settings with placeholder values (for human review)
    for name in sorted(new_names):
        settings.append({
            "id": make_slug(name),
            "name": name,
            "category": scraped[name],
            "description": "",
            "default_setting": "",
            "recommended_setting": "",
            "requires_adjustment": "",
            "impact": "",
            "notes": "New setting detected — pending human review.",
            "ms_docs_anchor": name,
            "status": "pending_review",
            "last_verified": today,
        })

    # Mark removed settings
    for s in settings:
        if s.get("ms_docs_anchor") in removed_names:
            s["status"] = "removed"
            s["last_verified"] = today

    save_settings(settings)
    print(f"\ndata/settings.json updated ({len(settings)} entries).")

    # Write diff summary to file
    DIFF_FILE.parent.mkdir(parents=True, exist_ok=True)
    DIFF_FILE.write_text(summary, encoding="utf-8")
    print(f"Diff summary written to {DIFF_FILE.relative_to(REPO_ROOT)}.")

    return 1  # Signal GitHub Actions to treat this run as a failure


if __name__ == "__main__":
    sys.exit(main())
