# Engineering KPI Dashboard (Excel + VBA Automation)

**What it does:** Tracks engineering throughput and quality KPIs for transformer projects.  
**Why it matters:** Gives managers a single source of truth for delivery visibility and bottlenecks.

## Tech
Excel (Power Query, Pivot), VBA (auto-refresh/export), Optional: PostgreSQL

## KPIs
- On-time completion %
- WIP aging (days)
- First-pass yield %
- Rework rate
- Throughput (tasks/week)

## Data Flow
`/data/mock/*.csv` → Power Query → Pivot/Measures → KPI Cards (Dashboard sheet)

## Run
1) Open `Engineering_KPI_Dashboard.xlsx`
2) Press **Ctrl+Shift+R** (macro) → refreshes data & exports PDF to `/exports/`

## Files
- `/data/mock/`: sample CSVs (projects.csv, tasks.csv, defects.csv)
- `/assets/`: screenshots
- `/exports/`: auto-exports (PDF/PNG)
- `/vba/RefreshAndExport.bas`: macro source

## Business Impact
- Cut monthly prep time by 70% through one-click refresh & export.
- Standardized KPI definitions for exec reporting.

## Next
- Connect to PostgreSQL via ODBC.
- Publish summary PDF automatically (task scheduler).
