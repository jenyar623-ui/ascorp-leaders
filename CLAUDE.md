# Teams & Customers Dashboard

## Project Description

Interactive HTML/JSON operational and client dashboard for an IT service company. Displays team workload, productivity, tickets, SLA, and field visits across multiple teams and customers. Single-page application with embedded data, styles, and logic. Only external dependency is Chart.js (loaded via CDN).

The dashboard has two main views:
- **Operational (daily)** — team workload, hours, tickets, field visits by employee/team/day
- **Client (monthly)** — customer-level metrics: hours by client, ticket flow, SLA compliance

## Tech Stack

- **HTML5** — single-file dashboard (`dashboard_v7.html`)
- **CSS3** — dark theme, responsive layout (embedded in HTML)
- **JavaScript (vanilla)** — all filtering, charting, table rendering logic (embedded)
- **Chart.js** — charting library (loaded from CDN)
- **JSON** — all dashboard data embedded inline (sourced from `v3_data.json`)
- **Python** — build script to assemble the final HTML from component files

## Key Files

| File | Description |
|---|---|
| `dashboard_v7.html` | The assembled dashboard — open this in a browser to view |
| `v3_data.json` | All dashboard data in JSON format (~920 KB) |
| `new_css.txt` | CSS source (dark theme, responsive) — ~510 lines |
| `new_body.txt` | HTML body markup (both tabs + sidebar) — ~160 lines |
| `new_js.txt` | JavaScript logic (filters, charts, tables) — ~1180 lines |
| `DASHBOARD_DOCUMENTATION.md` | Full technical documentation (in Russian) |

## Data Sources (OneDrive, read via temp-copy)

- `Операционные отчеты (ежедневные).xlsx` — daily operational reports by team (Dec 2025 – Feb 2026)
- `Отчет по клиентам (ежемесячный).xlsx` — monthly client reports (Aug 2025 – Jan 2026)
- Path: `~/Library/CloudStorage/OneDrive-AscorpSP/My Obsidian/FinancesDocs/`

## Leaders Dashboards

After each build, `dashboard_v7.html` is automatically copied to OneDrive `Leaders Dashboards/` as `teams-customers-dashboard.html` for sharing with colleagues.

## How to Run / View

1. Open `dashboard_v7.html` in any modern browser (Chrome, Firefox, Edge, Safari).
2. No server or build step needed — everything is self-contained.
3. Use the "Фильтры" (Filters) button to filter by period, team, or client.
4. Switch between "Операционный" (Operational) and "Клиентский" (Client) tabs.

## How to Rebuild

If you modify the component files (`new_css.txt`, `new_body.txt`, `new_js.txt`, `v3_data.json`), reassemble the dashboard with Python:

```python
import json

with open('new_css.txt') as f: css = f.read()
with open('new_body.txt') as f: body = f.read()
with open('new_js.txt') as f: js = f.read()
with open('v3_data.json') as f: data = json.load(f)

data_json = json.dumps(data, ensure_ascii=False)

html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Информационная панель</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
{css}
</style>
{body}
<script>
const D = {data_json};
{js}
</script>
</html>"""

with open('dashboard_v7.html', 'w') as f:
    f.write(html)
```

## Safety

- Excel files are never opened directly — `safe_load_workbook()` copies to a temp file first, reads, then deletes the temp
- This prevents any interaction with the original (OneDrive locks, sync conflicts, Data Validation preservation)
- Retry logic (3 attempts, 5s delay) handles transient OneDrive sync issues (e.g. `BadZipFile` errors during sync)
