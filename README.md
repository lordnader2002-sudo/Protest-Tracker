# Protest Tracker (Daily Report)

## ✅ Download the Latest Tracker
Open this link and click **“Download latest report (.xlsx)”**:

**GitHub Pages report page:**  
https://lordnader2002-sudo.github.io/protest-tracker/

> If you ever see “File not found”, wait 1–2 minutes and refresh. Sometimes GitHub Pages takes a moment to update after the daily run.

## Screenshot: Where to Download
![Download latest tracker button](docs/images/download.png)

---

## What this is
This repo automatically generates a daily Excel report of:

- **Protests/events near company properties** (General view)
- **“No Kings” related events** (NoKings tab)

It runs automatically using GitHub Actions and publishes the newest report to a simple webpage so anyone can download it.

---

## What’s inside the Excel file
The downloaded Excel file contains multiple tabs:

### 1) `Protests` (General)
Events found in the **next 7 days**.

### 2) `NoKings`
No Kings events found in the **next 30 days** (for collection/tracking).

### 3) Property tabs (optional)
If enabled, you may also see one tab per property showing matches for that property.

---

## What “Is New” means
The tracker keeps a memory of previously discovered events.

- **Is New = TRUE** → this event was discovered for the first time ever
- **Is New = FALSE** → it was already found on a previous run

New events are highlighted so you can spot them quickly.

---

## When it runs
The tracker runs **once per day** via GitHub Actions.

Note: GitHub schedules are based on UTC time. The job is configured to approximately match a “daily noon ET” cadence.

---

## Where the report comes from
The tracker collects event data from:
- **Mobilize.us** (main source)
- **Action Network** (No Kings public event pages)

---

## For technical users (optional)
- Script: `scripts/Simon OIC Intel - Protest Tracker Script v9.1.py`
- Properties list: `data/properties.csv`
- Workflow: `.github/workflows/protest_tracker.yml`
- Published report file:
  - `docs/latest.xlsx` (latest)
  - `docs/previous.xlsx` (previous)
  - `docs/last_updated.txt` (timestamp)
- Event history store:
  - `seen_events.json` (used to flag “Is New”)

---

## Troubleshooting (simple)
- **Download link shows 404:** wait a minute and refresh; GitHub Pages may still be deploying.
- **Excel opens but looks empty:** that usually means there were no events within the configured radius/time window.
- **Some properties are skipped:** usually caused by invalid/missing zip codes or missing lat/lon in the properties file.

---
