# Microsoft Fabric Tenant Settings

An automatically maintained reference tool for Microsoft Fabric tenant settings.

**Live tool → https://bachovak.github.io/fabrictenantsettings/**

---

## What this does

| Component | Description |
|-----------|-------------|
| `data/settings.json` | Master data file — 152 settings with governance recommendations |
| `scripts/scrape.py` | Fetches the MS Docs page weekly and diffs against the JSON |
| `scripts/build.py` | Generates the interactive HTML table from the JSON |
| `docs/index.html` | The client-facing tool — searchable table with Excel export |
| `.github/workflows/update.yml` | GitHub Actions workflow — runs every Monday at 09:00 UTC |

---

## 1. Enable GitHub Pages

1. Go to **Repository Settings → Pages**
2. Under **Source**, select **Deploy from a branch**
3. Set **Branch** to `main` and **Folder** to `/docs`
4. Click **Save**
5. After a minute, the tool will be live at:
   `https://bachovak.github.io/fabrictenantsettings/`

---

## 2. Test the workflow manually

Before waiting for Monday's scheduled run, trigger it manually:

1. Go to the **Actions** tab in this repository
2. Click **Update Fabric Tenant Settings** in the left sidebar
3. Click **Run workflow → Run workflow**
4. Watch the run complete — it should exit successfully (no changes on first run)

---

## 3. How notifications work

No email secrets or SMTP setup is needed.

When the scraper detects changes (new or removed settings), it exits with
code 1, which causes the GitHub Actions job to **fail**. GitHub automatically
sends a failure notification email to the repository owner.

**To confirm your notification settings:**

1. Go to **GitHub → Settings → Notifications**
2. Under **Actions**, make sure **Failed workflows** email notifications
   are enabled for this repository (or for all repositories)

The diff summary (which settings changed) will be visible in the workflow
run log on the **Actions** tab.

---

## 4. Reviewing new or removed settings

When you receive a notification email that the workflow failed:

### New settings (`status: "pending_review"`)

1. Open `data/settings.json`
2. Find entries where `"status": "pending_review"`
3. Fill in these fields based on the Microsoft documentation:
   - `default_setting`
   - `recommended_setting`
   - `requires_adjustment` (`"Yes"` or `"No"`)
   - `impact` (`"High"`, `"Medium"`, or `"Low"`)
   - `notes` (governance rationale)
4. Change `"status"` to `"current"`
5. Update `"last_verified"` to today's date (ISO format: `YYYY-MM-DD`)

### Removed settings (`status: "removed"`)

Settings no longer found on the MS Docs page are automatically marked
`"removed"`. They remain in the JSON for historical reference and appear
greyed out with strikethrough text in the HTML table.

If a setting reappears after being marked removed, update its `status`
back to `"current"`.

### Regenerate the HTML

After updating `settings.json`, regenerate the HTML table:

**Option A — Locally:**
```bash
pip install -r requirements.txt
python scripts/build.py
```

**Option B — GitHub Actions:**
1. Go to **Actions → Update Fabric Tenant Settings**
2. Click **Run workflow**

---

## 5. Run scripts locally

```bash
# Clone and install
git clone https://github.com/bachovak/fabrictenantsettings.git
cd fabrictenantsettings
pip install -r requirements.txt

# Build the HTML table from the current JSON
python scripts/build.py

# Scrape the MS Docs page and check for changes
python scripts/scrape.py
```

---

## Repository structure

```
fabrictenantsettings/
├── .github/
│   └── workflows/
│       └── update.yml          # GitHub Actions — weekly schedule + manual trigger
├── data/
│   ├── settings.json           # Master data file (152 settings)
│   └── last_diff.txt           # Written when changes are detected (auto-generated)
├── scripts/
│   ├── scrape.py               # Scrapes MS Docs; diffs against settings.json
│   └── build.py                # Generates docs/index.html from settings.json
├── docs/
│   └── index.html              # GitHub Pages output — the interactive tool
├── requirements.txt            # Pinned Python dependencies
└── README.md
```

---

## Embedding in an external website

The `docs/index.html` page is safely embeddable in an iframe:

```html
<iframe
  src="https://bachovak.github.io/fabrictenantsettings/"
  width="100%"
  height="800"
  style="border: none;"
  title="Microsoft Fabric Tenant Settings"
></iframe>
```

---

## Data sources

- [Microsoft Fabric tenant settings index](https://learn.microsoft.com/en-us/fabric/admin/tenant-settings-index)
- Governance recommendations by [Kristina Bachová](https://kristinabachova.com)
