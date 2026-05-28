# CLAUDE.md — Alpha Schools SY26/27 Budget Portfolio
> This file is written for Claude. Read it fully before doing anything. It contains everything you need to work on this project without asking questions that were already answered.

---

## What This Project Is

This is the **Alpha Schools SY2026/27 annual budget portfolio** — a system for building, managing, and viewing financial models for every Alpha school location. It was built by Kadmi (kadmiel.vonspringer@trilogy.com) and handed off to the team in May 2026.

The portfolio covers **43 school locations** across the US, each with three financial scenarios: BASE (expected), CONSTRAINED (downside), STRETCH (upside). Each school has a 5-year P&L model covering enrollment, revenue, headcount costs, facility costs, and EBITDA.

---

## File Map — What Everything Does

### Core Files (always keep these in sync)
| File | What it is |
|------|-----------|
| `Alpha_Consolidation_SY2627.xlsx` | Master portfolio workbook. 7 tabs: Config, BASE, CONSTRAINED, STRETCH, Summary, _Data, WORKFLOW. The `_Data` tab feeds the dashboard. **Never edit manually — always regenerate with the script.** |
| `refresh_consolidation.py` | The main script. Reads all school xlsx files, builds the consolidation Excel workbook from scratch. Run this whenever any school budget changes. |
| `Alpha_Portfolio_Dashboard.html` | Self-contained web app. All portfolio data embedded as JSON. Regenerate this after running refresh_consolidation.py. |

### School Budget Files (22 active models)
All named `Alpha_[School]_SY2627_Budget.xlsx`. Each has 5 tabs: Assumptions, BASE, CONSTRAINED, STRETCH, Summary.

| File | School | Status |
|------|--------|--------|
| Alpha_Chantilly_SY2627_Budget.xlsx | Chantilly VA | Active |
| Alpha_Charlotte_SY2627_Budget.xlsx | Charlotte NC | Active |
| Alpha_Tampa_SY2627_Budget.xlsx | Tampa FL | Active |
| Alpha_PaloAlto_SY2627_Budget.xlsx | Palo Alto CA | Active |
| Alpha_Boston_SY2627_Budget.xlsx | Boston MA | Active |
| Alpha_Bethesda_SY2627_Budget.xlsx | Bethesda MD | Active |
| Alpha_NewYork_SY2627_Budget.xlsx | New York NY | Active |
| Alpha_Woodlands_SY2627_Budget.xlsx | The Woodlands TX | Active |
| Alpha_Keller_SY2627_Budget.xlsx | Keller TX | Active |
| Alpha_SantaMonica_SY2627_Budget.xlsx | Santa Monica CA | Active |
| Alpha_Dorado_SY2627_Budget.xlsx | Dorado PR | Active |
| Alpha_Raleigh_SY2627_Budget.xlsx | Raleigh NC | Active |
| Alpha_Atlanta_Roswell_SY2627_Budget.xlsx | Atlanta (Roswell) GA | Active |
| Alpha_SantaBarbara_SY2627_Budget.xlsx | Santa Barbara CA | Active — flagship, high cost |
| Alpha_Chicago_SY2627_Budget.xlsx | Chicago IL | Active — flagship, high cost |
| Alpha_LaJolla_SY2627_Budget.xlsx | La Jolla CA | Active |
| Alpha_PalmBeachGardens_SY2627_Budget.xlsx | Palm Beach Gardens FL | Active |
| Alpha_Kirkland_SY2627_Budget.xlsx | Kirkland WA | Active — original template school |
| Alpha_Plano_SY2627_Budget.xlsx | Plano TX | Active |
| Alpha_Houston_SY2627_Budget.xlsx | Houston TX | Active |
| Alpha_Austin_SY2627_Budget.xlsx | Austin TX | Active — very high cost location |
| Alpha_Miami_SY2627_Budget.xlsx | Miami FL | Active — very high cost location |
| Alpha_SanFrancisco_SY2627_Budget.xlsx | San Francisco CA | Active — large school, 80 students |
| Alpha_Tulsa_SY2627_Budget.xlsx | Tulsa OK | Active |
| Alpha_FortWorth_SY2627_Budget.xlsx | Fort Worth TX | Active |
| Alpha_BocaRaton_SY2627_Budget.xlsx | Boca Raton FL | Active |
| Alpha_Greenwich_SY2627_Budget.xlsx | Greenwich CT | Active |
| Alpha_ParkCity_SY2627_Budget.xlsx | Park City UT | Active |

### Schools Without Models Yet (show in dashboard as "No Model")
Orange County CA, Scottsdale AZ, East Bay CA, Armonk NY, Dallas TX, Edmond OK, Jamaica Plain MA, Lexington MA, Los Angeles CA, Malibu CA, New York (775 Columbus), San Juan PR, Santa Clara CA, Tulsa (421 E 11th), Williston VT

### Other Files
| File | What it is |
|------|-----------|
| `WPB_Public_School_Cost_Model.xlsx` | One-off cost model for WPB public school operator pitch (3 schools, 4 enrollment scenarios). Not part of the main portfolio. |
| `Alpha_[School]_Assumptions_SY2627.docx` | Assumptions narrative docs for each school. Linked from the dashboard Config tab. All stored on Google Drive at: https://drive.google.com/drive/folders/1Y-C00Wf4DaOiIQTrEKOKcmdKiyQXAkoT |

---

## How the System Works

### The data flow
```
Individual school xlsx files
        ↓
refresh_consolidation.py
        ↓
Alpha_Consolidation_SY2627.xlsx  (especially the _Data tab)
        ↓
build_dashboard.py  (or manual Python extraction)
        ↓
Alpha_Portfolio_Dashboard.html
```

### School budget structure (every school file is identical)
- **Assumptions tab**: All inputs — tuition, enrollment ramp, rent, headcount, etc. Blue text = hardcoded inputs. Black text = formulas.
- **BASE tab**: Expected scenario. Columns: Actuals | Q1 | Q2 | Q3 | Q4 | Y1 Full | Y2 | Y3 | Y4 | Y5 | 5yr Total | Notes
- **CONSTRAINED tab**: Downside scenario (lower enrollment, same costs)
- **STRETCH tab**: Upside scenario (higher enrollment, same costs)
- **Summary tab**: One-page summary across all 3 scenarios

### Key assumptions (standard across most schools)
- Enrollment cap: 25 students Y1, ramp to capacity Y2-Y4
- Tuition: varies by market ($40K–$75K/yr typical range)
- Lead Guide: 1 per school, ~$120K fully loaded
- Guides: 1 per 25 students, ~$70K fully loaded (115% loaded salaries)
- Timeback: 20% of core tuition revenue
- Programs: ~$800/student/yr
- Facility costs: vary by lease (see Assumptions tab)

---

## How to Update Things

### When a school's numbers change
1. Open the school's xlsx file (e.g. `Alpha_Charlotte_SY2627_Budget.xlsx`)
2. Update the Assumptions tab (blue cells only)
3. Save the file
4. Run: `python3 refresh_consolidation.py`
5. Rebuild the dashboard (see below)

### How to rebuild the dashboard HTML
The dashboard is built by extracting data from the consolidation xlsx and baking it into the HTML. The quickest way:

```python
# Run this in the project folder
python3 build_dashboard.py
```

If `build_dashboard.py` doesn't exist yet, tell Claude to "rebuild the Alpha_Portfolio_Dashboard.html from the consolidation file" — Claude knows how to do this from scratch.

Alternatively, you can ask Claude directly: *"The consolidation file has been updated. Please regenerate Alpha_Portfolio_Dashboard.html."*

### How to add a new school
1. Copy an existing school file as a template (Kirkland is the cleanest template)
2. Rename it `Alpha_[SchoolName]_SY2627_Budget.xlsx`
3. Update the Assumptions tab with the new school's data
4. Add the school to the `SCHOOLS` list in `refresh_consolidation.py`:
```python
{"name": "Alpha [City]", "file": "Alpha_[City]_SY2627_Budget.xlsx", 
 "location": "City, ST", 
 "assumptions_url": "https://drive.google.com/file/d/[GDRIVE_ID]/view"},
```
5. Run `refresh_consolidation.py`
6. Rebuild the dashboard

### How to update the school filter (Y/N) in the consolidation
The BASE tab has a school filter in columns A-B (rows 3–62). Set column B to Y/N to include/exclude schools from portfolio totals. This is interactive in Excel — the consolidation file reads these dropdowns.

---

## Dashboard — How It Was Built

The `Alpha_Portfolio_Dashboard.html` is a single self-contained HTML file (~380KB). All data is embedded as a JSON object called `RAW` at the top of the script section. The app uses:
- **Chart.js** (CDN: cdnjs.cloudflare.com) for charts — requires internet
- **Vanilla JS** — no frameworks, no build step
- **Embedded JSON** — all school and P&L data baked in at build time

### Dashboard data structure
```javascript
RAW.schools[]           // 43 schools with metadata (name, location, y1_enrollment, y1_ebitda, etc.)
RAW.pnl[school][scenario][rowkey][period]  // P&L values
  // periods: Y1, Y2, Y3, Y4, Y5, Actuals, Q1, Q2, Q3, Q4
  // rowkeys: enrollment, core_rev, ah_rev, ss_rev, total_rev, 
  //          lead_guides, hos, guides, admin, total_hc,
  //          timeback, programs, misc, ext_cost,
  //          base_rent, nnn, utilities, rm_fac, rm_grd, cleaning, it_net, security,
  //          total_fac, da, total_costs, ebitda, ebit
```

### To add a new view or feature to the dashboard
The HTML is structured as a single `<script>` block at the bottom of the file. All rendering functions follow the pattern `render[ViewName]()`. The main entry point is `renderAll()`, called at the very bottom of the script (after all `const` declarations — this ordering is important, do not move `renderAll()` up).

---

## What's Still Pending / Known Issues

### Schools needing attention
- **Fort Worth**: Model exists but needs target date confirmed
- **Orange County, Scottsdale, East Bay, Armonk, Dallas, Edmond, Jamaica Plain, Lexington, Los Angeles, Malibu, NY 775 Columbus, San Juan, Santa Clara, Tulsa (421 E 11th), Williston**: No budget models built yet — show as "No Model" in dashboard
- **Assumptions docs audit**: Not all Assumptions docs have been verified against their corresponding Excel files (task was in progress)

### Known model quirks
- **Chantilly**: Only 3 students enrolled (actuals) vs 25 projected — the Y1 revenue looks low because actuals pulled through
- **Chicago & Santa Barbara**: Flagship schools — much higher costs (HoS salary, larger facility). Y1 EBITDA deeply negative by design
- **Austin & Miami**: High-cost markets. Y1 losses are expected — check Y2/Y3 for profitability path
- **San Francisco**: 80-student model — much larger than typical 25-student school

### Google Drive
All Assumptions docs live here: https://drive.google.com/drive/folders/1Y-C00Wf4DaOiIQTrEKOKcmdKiyQXAkoT
School budget xlsx files are also uploaded there (folder managed by the team).

---

## Key Conventions to Preserve

- **Color coding in Excel**: Blue text = hardcoded inputs users change. Black text = formulas. Green text = cross-sheet links.
- **Row 3 height** in BASE/CONSTRAINED/STRETCH tabs: set to 36pt (wraps two-line headers)
- **The `_Data` tab** in the consolidation is the single source of truth for the dashboard — never edit it manually, it's fully generated by `refresh_consolidation.py`
- **School names must match exactly** between the SCHOOLS list in `refresh_consolidation.py` and the school files — the lookup is case-sensitive
- **EBIT vs EBITDA**: EBITDA = operating profit before D&A. EBIT = EBITDA minus D&A. D&A is in each school's Assumptions tab.

---

## Glossary
| Term | Meaning |
|------|---------|
| SY26/27 | School Year 2026–2027 (Jul 2026 – Jun 2027) |
| Y1–Y5 | Year 1 (SY26/27) through Year 5 (SY30/31) |
| BASE | Expected scenario |
| CONSTRAINED | Downside — lower enrollment, same costs |
| STRETCH | Upside — higher enrollment, same costs |
| HoS | Head of School |
| LG | Lead Guide |
| NNN / CAM | Triple-net / Common Area Maintenance lease charges |
| Timeback | Alpha's platform fee — 20% of core tuition revenue |
| GCC | Gross Cost per Child |
| DDR | Debt to Revenue ratio (used in some older models) |
| Loaded salary | Fully loaded = base salary × 115% (benefits, taxes, etc.) |
