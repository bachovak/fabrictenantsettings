# Maintenance Guide

This document explains what to do when GitHub detects changes to the
Microsoft Fabric tenant settings documentation and opens a Pull Request.

---

## When does this happen?

The tool runs automatically on the **1st of every month at 09:00 UTC**.
It scrapes the Microsoft Fabric documentation, compares it to the current
`data/settings.json`, and opens a Pull Request if anything changed.

You will receive a **GitHub notification email** with a link to the PR.

---

## Step 1 — Open the Pull Request

1. Click the link in the notification email, or go to
   [github.com/bachovak/fabrictenantsettings/pulls](https://github.com/bachovak/fabrictenantsettings/pulls)
2. Open the PR titled **"auto: Fabric tenant settings update YYYY-MM-DD"**
3. Read the PR description — it lists exactly what changed:
   - **New settings** — appear in the documentation for the first time
   - **Removed settings** — no longer present in the documentation
   - **Changed settings** — description or name updated by Microsoft

---

## Step 2 — Decide what needs your attention

### Removed or changed settings → no action needed

These are handled automatically. Removed settings show as struck-through
in the live table. Changed descriptions are updated in place. You can
skip straight to Step 4.

### New settings → you need to fill in the governance data

New settings are added to `data/settings.json` with `"status": "pending_review"`
and are **hidden from the live site** until you review them. Go to Step 3.

---

## Step 3 — Fill in the governance data for new settings

1. In the PR, click the **Files changed** tab
2. Find `data/settings.json` and click the **...** menu → **Edit file**
   (or open the file directly and click the pencil icon)
3. Search for `"pending_review"` — each match is a new setting
4. For each new setting, fill in the four governance fields and change the status:

```json
{
  "name": "Example Setting Name",
  "ms_docs_anchor": "Example Setting Name",
  "category": "Export and sharing settings",
  "description": "Allows users to ...",

  "default_setting":     "On",
  "recommended_setting": "Off – disable for all users",
  "requires_adjustment": "Yes",
  "impact":              "High",
  "notes":               "Reason for recommendation or any context.",

  "status": "current"
}
```

**Field guidance:**

| Field | What to write |
|-------|--------------|
| `default_setting` | Copy from the Microsoft documentation or leave as scraped |
| `recommended_setting` | What a well-governed tenant should have, e.g. `"Off – disable for all users"` or `"On – enable for the entire organisation"` |
| `requires_adjustment` | `"Yes"` if the default does not match the recommendation; `"No"` if it already matches |
| `impact` | `"High"`, `"Medium"`, or `"Low"` — how much governance risk leaving the default creates |
| `notes` | Optional. Rationale, exceptions, or links to relevant policies |
| `status` | Change from `"pending_review"` to `"current"` when done |

5. Commit the changes directly to the PR branch
   (select **"Commit directly to the `auto/settings-...` branch"**)

---

## Step 4 — Rebuild the HTML

After editing `settings.json`, the live page (`docs/index.html`) needs to
be regenerated to reflect your changes.

**Option A — Let Claude Code do it (recommended)**

Open Claude Code in the tool's local folder and ask:
> "New settings were reviewed — please run build.py and commit the result to the PR branch."

**Option B — Run it yourself locally**

```bash
# Pull the PR branch
git fetch origin
git checkout auto/settings-YYYY-MM-DD

# Regenerate the HTML
python scripts/build.py

# Commit and push
git add docs/index.html
git commit -m "rebuild: regenerate index.html after review"
git push
```

---

## Step 5 — Merge the Pull Request

1. Go back to the PR page
2. Confirm the **Files changed** tab shows both `data/settings.json` and
   `docs/index.html` with your updates
3. Click **Merge pull request** → **Confirm merge**

GitHub Pages will pick up the new `docs/index.html` within a minute and
`fabric.kristinabachova.com` will be live with the updated table.

---

## Manual trigger

If you want to run the scraper outside the monthly schedule (e.g. after
Microsoft announces a major Fabric update):

1. Go to the [Actions tab](https://github.com/bachovak/fabrictenantsettings/actions)
2. Select **Update Fabric Tenant Settings** in the left sidebar
3. Click **Run workflow** → **Run workflow**

---

## Reference: status values in settings.json

| Status | Meaning | Visible on site |
|--------|---------|-----------------|
| `current` | Reviewed and approved | Yes |
| `pending_review` | Newly detected — not yet reviewed | **No** |
| `removed` | No longer in Microsoft documentation | Yes (struck through) |
