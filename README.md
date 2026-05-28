# Alpha Schools — SY26/27 Budget Portfolio

Built by Kadmi (May 2026). Handed off to the team on his last day.

---

## What's in This Folder

| File | What it is |
|------|-----------|
| `Alpha_Consolidation_SY2627.xlsx` | Master portfolio Excel file — all 43 schools, 3 scenarios |
| `Alpha_Portfolio_Dashboard.html` | Interactive web dashboard — open in any browser |
| `refresh_consolidation.py` | Script that rebuilds the Excel consolidation from school files |
| `Alpha_[School]_SY2627_Budget.xlsx` | Individual school budget models (22 schools) |
| `CLAUDE.md` | Full technical briefing for Claude AI — read this if you're using Claude to make changes |

---

## Viewing the Dashboard

Just open `Alpha_Portfolio_Dashboard.html` in Chrome or any browser. No install needed.

If it's hosted on GitHub Pages, use the team URL instead.

**What you can do in the dashboard:**
- Switch between BASE / CONSTRAINED / STRETCH scenarios
- Filter to specific schools using the left sidebar
- View all schools side by side in the P&L Table
- See Q1/Q2/Q3/Q4 quarterly breakdown plus full 5-year view
- Check the Issues tab for schools with losses or concerns
- Click any school in Overview to drill into its full P&L

---

## Making Changes (with Claude)

The fastest way to update anything is to open this folder in **Claude Cowork** (or upload files to Claude.ai) and ask in plain English:

> "Update the Chicago budget — rent increased to $45K/month"

> "Add a new school: Alpha Scottsdale, Arizona, 25 students, $52K tuition"

> "Rebuild the dashboard with the latest numbers"

> "Show me which schools turn profitable by Year 3"

Claude knows this entire project — just read `CLAUDE.md` if starting a new session and it will brief Claude automatically.

---

## Updating Numbers Manually

If you prefer to edit directly without Claude:

1. Open the relevant school file (e.g. `Alpha_Charlotte_SY2627_Budget.xlsx`)
2. Go to the **Assumptions** tab
3. Edit the **blue cells** only (hardcoded inputs)
4. Save the file
5. Run the consolidation script:
   ```
   python3 refresh_consolidation.py
   ```
6. Ask Claude to rebuild the dashboard HTML, or run the build script if one exists

---

## Adding a New School

1. Copy `Alpha_Kirkland_SY2627_Budget.xlsx` as your template
2. Rename it `Alpha_[NewCity]_SY2627_Budget.xlsx`
3. Fill in the Assumptions tab for the new school
4. Add the school to the `SCHOOLS` list in `refresh_consolidation.py`
5. Run `python3 refresh_consolidation.py`
6. Rebuild the dashboard

---

## Deploying the Dashboard (GitHub Pages)

When you update the dashboard HTML and want to publish it:

1. Push `Alpha_Portfolio_Dashboard.html` to the GitHub repo
2. GitHub Pages auto-publishes it at the team URL
3. Anyone with the link sees the updated version immediately

---

## School Status Summary

**28 active schools with models** (show data in dashboard)
Chantilly VA · Charlotte NC · Tampa FL · Palo Alto CA · Boston MA · Bethesda MD · New York NY · The Woodlands TX · Keller TX · Santa Monica CA · Dorado PR · Raleigh NC · Atlanta GA · Santa Barbara CA · Chicago IL · La Jolla CA · Palm Beach Gardens FL · Kirkland WA · Plano TX · Houston TX · Austin TX · Miami FL · San Francisco CA · Tulsa OK · Fort Worth TX · Boca Raton FL · Greenwich CT · Park City UT

**15 locations without models yet** (show as "No Model" in dashboard)
Orange County CA · Scottsdale AZ · East Bay CA · Armonk NY · Dallas TX · Edmond OK · Jamaica Plain MA · Lexington MA · Los Angeles CA · Malibu CA · New York (775 Columbus) · San Juan PR · Santa Clara CA · Tulsa (421 E 11th) · Williston VT

---

## Google Drive

All Assumptions documents (narrative docs per school) live here:
https://drive.google.com/drive/folders/1Y-C00Wf4DaOiIQTrEKOKcmdKiyQXAkoT

---

## Questions?

Read `CLAUDE.md` — it has the full technical detail. Or open Claude and ask.
