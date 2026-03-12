# Protest Tracker

Automatically monitors **Mobilize.us** and **Action Network** for protests and events. Runs on a schedule via GitHub Actions, publishes results to GitHub Pages, and sends the Excel report as a workflow artifact.

---

## Accessing the Report

**Live web report (HTML):**
👉 https://lordnader2002-sudo.github.io/Protest-Tracker/

The page includes:
- **Summary tab** — event counts by distance, property, and organization
- **Protests tab** — general events within the search radius
- **No Kings tab** — No Kings–specific events
- **Map tab** — interactive map with property and event pins (click any pin for details)
- **Download Excel** button at the top of the page

The Excel file (`docs/latest.xlsx`) is also available directly from the page or as a GitHub Actions artifact attached to each run.

> If you see a 404, wait 1–2 minutes and refresh — GitHub Pages can take a moment after a new run.

---

## Excel Workbook

When you open the downloaded `.xlsx` file you'll see the following sheets:

| Sheet | Contents |
|---|---|
| **Summary** | Counts by distance band, property, and organization |
| **Protests** | General events found within the next ~7 days |
| **NoKings** | No Kings events found within the next 30 days (configurable) |
| **AllMatches** *(hidden)* | Every property/event pair within radius (right-click any tab → Unhide) |

### Row color coding

| Color | Meaning |
|---|---|
| 🔴 Red | Event is **< 1 mile** from a property |
| 🟠 Amber | Event is **1–2 miles** from a property |
| 🟢 Green | Event is **> 2 miles** from a property |

### Special columns

| Column | Meaning |
|---|---|
| **Is New** | `TRUE` = first time this event has been detected (highlighted gold) |
| **Is Duplicate** | `TRUE` = another row shares the same title, date, and location (shown in italic gray) |
| **Event URL** | Clickable hyperlink to the event page |

Each sheet also has:
- **Column auto-filter** dropdowns for quick filtering
- **Thick borders** separating each distance group
- **Bold** on key columns (name, date, distance)

---

## Live HTML Report Features

- **Sortable columns** — click any column header to sort ▲▼
- **Live text filter** — type in the filter box above a table to narrow rows instantly
- **Interactive map** — color-coded circle markers per event, 🏢 markers for properties; click any pin for a popup
- **Tabs** for Summary, Protests, No Kings, and Map
- Updates automatically after every run

---

## Schedule & Manual Runs

The tracker runs automatically **Monday, Wednesday, Friday at 5:00 PM UTC**.

You can also trigger it manually from the **Actions** tab → **Protest Tracker** → **Run workflow**, with optional overrides:

| Input | Default | Description |
|---|---|---|
| `radius_miles` | `3` | Search radius around each property |
| `no_kings_window_days` | `30` | How many days ahead to search for No Kings events |
| `workers` | `8` | Parallel threads for geocoding/fetching |

---

## Data Sources

| Source | Used For |
|---|---|
| [Mobilize.us](https://www.mobilize.us) | General protest/event search |
| [Action Network](https://actionnetwork.org) | No Kings event pages |

---

## Repository Structure

```
.
├── .github/
│   └── workflows/
│       └── protest_tracker.yml   # Scheduled + manual GitHub Actions workflow
├── data/
│   └── properties.csv            # Properties to monitor (name, address, lat, lon, zip)
├── docs/
│   ├── index.html                # Auto-generated HTML report (GitHub Pages)
│   ├── latest.xlsx               # Most recent Excel output
│   ├── previous.xlsx             # Previous run's Excel output
│   ├── map_data.json             # Event + property coordinates for the map
│   └── last_updated.txt          # Timestamp of last successful run
├── scripts/
│   ├── Simon OIC Intel - Protest Tracker Script v9.1.py   # Main tracker script
│   └── generate_report.py        # Generates docs/index.html from the Excel + map data
├── seen_events.json              # Persistent store of previously seen event IDs
└── requirements.txt              # Python dependencies
```

---

## Properties File (`data/properties.csv`)

Each row represents a property to monitor. Required columns:

| Column | Description |
|---|---|
| `property_id` | Unique identifier |
| `name` | Display name |
| `address` | Full street address |
| `lat` / `lon` | Coordinates (used for distance calculation) |
| `postal_code` | Property zip code |
| `query_zipcode` | Zip code used to query Mobilize (usually same as postal_code) |

---

## Troubleshooting

| Symptom | Likely Cause |
|---|---|
| Page shows 404 | GitHub Pages still deploying — wait 1–2 min and refresh |
| Excel opens empty | No events found within the radius/time window |
| Property skipped in output | Missing or invalid `lat`, `lon`, or `query_zipcode` in properties.csv |
| Map tab shows no pins | `map_data.json` not yet generated — run the workflow once |
| "Is New" always TRUE | `seen_events.json` was reset or is missing |
