#!/usr/bin/env python3
"""Generate docs/index.html from docs/latest.xlsx and docs/map_data.json."""

import html as html_mod
import json
import sys
from pathlib import Path

import pandas as pd

XLSX_PATH = Path("docs/latest.xlsx")
MAP_DATA_PATH = Path("docs/map_data.json")
TIMESTAMP_PATH = Path("docs/last_updated.txt")
HTML_PATH = Path("docs/index.html")

DIST_COL_MAIN = "Distance to Nearest Property (miles)"
DIST_COL_MATCH = "Distance to Property (miles)"

TAB_LABELS = {
    "Summary": "Summary",
    "Protests": "Protests",
    "NoKings": "No Kings",
    "AllMatches": "All Matches",
}


def row_bg(dist_val: object) -> str:
    try:
        d = float(dist_val)
        if d < 1.0:
            return "#FFCCCC"
        if d < 2.0:
            return "#FFE5B4"
        return "#CCFFCC"
    except (TypeError, ValueError):
        return ""


def df_to_html_table(df: pd.DataFrame, dist_col: str, table_id: str) -> str:
    if df.empty:
        return "<p style='padding:16px;color:#666'>No events found.</p>"

    header_cells = "".join(
        f"<th onclick=\"sortTable('{table_id}',{i})\" title='Click to sort'>"
        f"{html_mod.escape(str(c))} <span class='sort-arrow'>⇅</span></th>"
        for i, c in enumerate(df.columns)
    )

    rows: list[str] = []
    for _, row in df.iterrows():
        bg = row_bg(row.get(dist_col))
        is_dup = row.get("Is Duplicate", False)
        is_dup = is_dup is True or str(is_dup).upper() == "TRUE"
        is_new = row.get("Is New", False)
        is_new = is_new is True or str(is_new).upper() == "TRUE"

        styles: list[str] = []
        if bg:
            styles.append(f"background:{bg}")
        if is_dup:
            styles.append("color:#999;font-style:italic")
        elif is_new:
            styles.append("font-weight:bold")

        row_style = f' style="{";".join(styles)}"' if styles else ""

        cells: list[str] = []
        for col, v in zip(df.columns, row):
            text = str(v) if pd.notna(v) else ""
            if col == "Event URL" and text.startswith("http"):
                cells.append(
                    f'<td><a href="{html_mod.escape(text)}" target="_blank" '
                    f'rel="noopener">Link</a></td>'
                )
            else:
                cells.append(f"<td>{html_mod.escape(text)}</td>")
        rows.append(f"<tr{row_style}>{''.join(cells)}</tr>")

    return (
        f"<div class='table-controls'>"
        f"<input type='text' class='filter-input' placeholder='Filter rows...' "
        f"oninput=\"filterTable('{table_id}', this.value)\">"
        f"</div>"
        f"<div class='table-wrap'>"
        f"<table id='{table_id}'>"
        f"<thead><tr>{header_cells}</tr></thead>"
        f"<tbody>{''.join(rows)}</tbody>"
        f"</table></div>"
    )


def build_map_section(map_data: dict) -> str:
    """Return the Leaflet map HTML."""
    events = map_data.get("events", []) + map_data.get("no_kings_events", [])
    properties = map_data.get("properties", [])

    if not events and not properties:
        return ""

    events_js = json.dumps(events)
    properties_js = json.dumps(properties)

    return f"""
<div id="map-panel" class="tab-panel" hidden>
  <div id="map" style="height:80vh;min-height:500px;width:100%;border-radius:0 6px 6px 6px;"></div>
</div>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script>
var _mapInit = false;
function initMap() {{
  if (_mapInit) return;
  _mapInit = true;
  var events = {events_js};
  var properties = {properties_js};

  // Center on first property or first event
  var center = [39.5, -98.35];
  var zoom = 4;
  if (properties.length) {{ center = [properties[0].lat, properties[0].lon]; zoom = 10; }}
  else if (events.length) {{ center = [events[0].lat, events[0].lon]; zoom = 10; }}

  var map = L.map('map').setView(center, zoom);
  L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
    attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
    maxZoom: 19
  }}).addTo(map);

  function distColor(dist) {{
    if (dist === null || dist === undefined) return '#888';
    if (dist < 1.0) return '#cc0000';
    if (dist < 2.0) return '#cc7700';
    return '#007700';
  }}

  // Property markers (blue star)
  var propIcon = L.divIcon({{
    html: '<div style="font-size:22px;line-height:1">🏢</div>',
    className: '', iconAnchor: [11, 11]
  }});
  properties.forEach(function(p) {{
    L.marker([p.lat, p.lon], {{icon: propIcon}})
      .addTo(map)
      .bindPopup('<strong>' + p.name + '</strong><br>' + (p.address || ''));
  }});

  // Event markers (colored circles)
  events.forEach(function(ev) {{
    var color = distColor(ev.distance);
    var circle = L.circleMarker([ev.lat, ev.lon], {{
      radius: 8, color: color, fillColor: color,
      fillOpacity: 0.75, weight: 2,
      opacity: ev.is_duplicate ? 0.4 : 0.9
    }}).addTo(map);
    var badge = ev.is_new ? ' <span style="color:gold">★ NEW</span>' : '';
    var dup = ev.is_duplicate ? ' <em style="color:#999">(duplicate)</em>' : '';
    circle.bindPopup(
      '<strong>' + ev.title + '</strong>' + badge + dup +
      '<br>📅 ' + ev.date + (ev.time ? ' ' + ev.time : '') +
      '<br>📍 ' + ev.location +
      '<br>🏢 ' + ev.nearest_property +
      (ev.distance != null ? '<br>📏 ' + ev.distance.toFixed(2) + ' mi away' : '') +
      (ev.url ? '<br><a href="' + ev.url + '" target="_blank">View Event</a>' : '')
    );
  }});
}}
</script>
"""


def main() -> int:
    try:
        last_updated = TIMESTAMP_PATH.read_text(encoding="utf-8").strip()
    except OSError:
        from datetime import datetime, timezone
        last_updated = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    try:
        xl = pd.ExcelFile(XLSX_PATH, engine="openpyxl")
    except Exception as exc:
        print(f"ERROR: could not open {XLSX_PATH}: {exc}", file=sys.stderr)
        return 1

    sheets: dict[str, pd.DataFrame] = {}
    for name in xl.sheet_names:
        # Skip the hidden AllMatches sheet in the HTML (still visible via download)
        if name == "AllMatches":
            continue
        sheets[name] = pd.read_excel(xl, sheet_name=name)

    map_data: dict = {}
    try:
        map_data = json.loads(MAP_DATA_PATH.read_text(encoding="utf-8"))
    except OSError:
        pass  # map data optional

    has_map = bool(map_data.get("events") or map_data.get("properties"))

    # Build tab buttons and content panels
    tab_buttons: list[str] = []
    tab_panels: list[str] = []

    for i, (name, df) in enumerate(sheets.items()):
        label = TAB_LABELS.get(name, name)
        is_summary = name == "Summary"
        active_cls = " active" if i == 0 else ""
        hidden_attr = "" if i == 0 else " hidden"
        dist_col = DIST_COL_MAIN if DIST_COL_MAIN in df.columns else DIST_COL_MATCH
        count = f" ({len(df)})" if not is_summary else ""

        tab_buttons.append(
            f'<button class="tab-btn{active_cls}" onclick="showTab(this,\'{name}\')">'
            f'{html_mod.escape(label)}<span class="count">{count}</span></button>'
        )

        if is_summary:
            # Render summary as styled sections, not a sortable table
            inner = df_to_summary_html(df)
        else:
            table_id = f"tbl-{name}"
            inner = df_to_html_table(df, dist_col, table_id)

        tab_panels.append(
            f'<div id="tab-{name}" class="tab-panel"{hidden_attr}>{inner}</div>'
        )

    # Map tab
    map_section = ""
    if has_map:
        tab_buttons.append(
            '<button class="tab-btn" onclick="showTab(this,\'map\');initMap()">Map</button>'
        )
        map_section = build_map_section(map_data)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Protest Tracker</title>
<style>
  *, *::before, *::after {{ box-sizing: border-box; }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
    margin: 0; padding: 16px 20px; background: #f0f2f5; color: #222;
  }}
  h1 {{ margin: 0 0 4px; font-size: 1.4em; color: #111; }}
  .meta {{ color: #555; font-size: 0.88em; margin-bottom: 14px; }}
  .meta a {{ color: #0056b3; text-decoration: none; font-weight: 600; }}
  .meta a:hover {{ text-decoration: underline; }}
  .legend {{
    display: flex; gap: 14px; flex-wrap: wrap;
    margin-bottom: 14px; font-size: 0.82em; align-items: center;
  }}
  .legend-item {{ display: flex; align-items: center; gap: 5px; }}
  .swatch {{
    width: 14px; height: 14px; border-radius: 3px;
    border: 1px solid rgba(0,0,0,0.15); flex-shrink: 0;
  }}
  .tabs {{ display: flex; gap: 3px; flex-wrap: wrap; margin-bottom: -1px; }}
  .tab-btn {{
    padding: 7px 16px; border: 1px solid #ccc; border-bottom: none;
    background: #dde1e7; cursor: pointer; border-radius: 6px 6px 0 0;
    font-size: 0.88em; transition: background 0.15s;
  }}
  .tab-btn:hover {{ background: #cbd0d8; }}
  .tab-btn.active {{ background: #fff; font-weight: 700; border-bottom: 1px solid #fff; }}
  .count {{ font-weight: 400; color: #666; }}
  .tab-btn.active .count {{ color: #888; }}
  .tab-panel {{
    background: #fff; border: 1px solid #ccc;
    border-radius: 0 6px 6px 6px; overflow: hidden;
  }}
  .table-controls {{
    padding: 8px 12px; border-bottom: 1px solid #eee; background: #fafafa;
  }}
  .filter-input {{
    width: 280px; max-width: 100%; padding: 5px 10px;
    border: 1px solid #ccc; border-radius: 4px; font-size: 0.88em;
  }}
  .table-wrap {{ overflow-x: auto; max-height: 78vh; overflow-y: auto; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 0.82em; }}
  thead tr {{ position: sticky; top: 0; z-index: 1; }}
  th {{
    background: #1e2a3a; color: #fff; padding: 8px 11px;
    text-align: left; white-space: nowrap; font-weight: 600;
    cursor: pointer; user-select: none;
  }}
  th:hover {{ background: #2d3d52; }}
  .sort-arrow {{ opacity: 0.5; font-size: 0.8em; }}
  td {{ padding: 5px 11px; border-bottom: 1px solid #e8e8e8; white-space: nowrap; }}
  tbody tr:hover td {{ filter: brightness(0.93); }}
  a {{ color: #0056b3; }}
  /* Summary sheet */
  .summary-wrap {{ padding: 16px 20px; }}
  .summary-section {{ margin-bottom: 28px; }}
  .summary-section h2 {{ font-size: 1.1em; margin: 0 0 10px; color: #1e2a3a; }}
  .summary-table {{ border-collapse: collapse; font-size: 0.88em; min-width: 260px; }}
  .summary-table th {{
    background: #1e2a3a; color: #fff; padding: 6px 14px;
    text-align: left; font-weight: 600; cursor: default;
  }}
  .summary-table td {{ padding: 5px 14px; border-bottom: 1px solid #e0e0e0; }}
  @media (prefers-color-scheme: dark) {{
    body {{ background: #0f1117; color: #dde1e8; }}
    h1 {{ color: #e8eaf0; }}
    .meta {{ color: #8a9bb0; }}
    .meta a {{ color: #5b9bd5; }}
    .tab-btn {{ background: #1e2635; border-color: #333d4f; color: #c5cdd9; }}
    .tab-btn:hover {{ background: #28334a; }}
    .tab-btn.active {{ background: #0f1117; border-color: #444f63; border-bottom-color: #0f1117; color: #e8eaf0; }}
    .count {{ color: #5e6e84; }}
    .tab-btn.active .count {{ color: #6b7d96; }}
    .tab-panel {{ background: #0f1117; border-color: #2a3347; }}
    .table-controls {{ background: #141921; border-bottom-color: #2a3347; }}
    .filter-input {{ background: #1a2236; border-color: #2a3347; color: #dde1e8; }}
    td {{ border-bottom-color: #1e2635; }}
    tbody tr:hover td {{ filter: brightness(1.15); }}
    a {{ color: #5b9bd5; }}
    .summary-wrap {{ background: #0f1117; }}
    .summary-section h2 {{ color: #7b9dc7; }}
    .summary-table td {{ border-bottom-color: #1e2635; }}
    .swatch {{ border-color: rgba(255,255,255,0.15); }}
  }}
</style>
</head>
<body>
<h1>&#128205; Protest Tracker</h1>
<div class="meta">
  Updated: <strong>{html_mod.escape(last_updated)}</strong>
  &nbsp;&bull;&nbsp;
  <a href="latest.xlsx">&#11015; Download Excel</a>
</div>
<div class="legend">
  <div class="legend-item"><div class="swatch" style="background:#FFCCCC"></div> &lt;1 mile</div>
  <div class="legend-item"><div class="swatch" style="background:#FFE5B4"></div> 1&ndash;2 miles</div>
  <div class="legend-item"><div class="swatch" style="background:#CCFFCC"></div> &gt;2 miles</div>
  <div class="legend-item" style="color:#999;font-style:italic">italic gray = duplicate</div>
  <div class="legend-item"><strong>bold = new this run</strong></div>
</div>
<div class="tabs">{"".join(tab_buttons)}</div>
{"".join(tab_panels)}
{map_section}
<script>
function showTab(btn, name) {{
  document.querySelectorAll('.tab-panel, #map-panel').forEach(el => {{ el.hidden = true; }});
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
  var panel = document.getElementById('tab-' + name) || document.getElementById('map-panel');
  if (panel) panel.hidden = false;
  btn.classList.add('active');
}}

function filterTable(tableId, query) {{
  var q = query.toLowerCase();
  var rows = document.getElementById(tableId).querySelectorAll('tbody tr');
  rows.forEach(function(row) {{
    row.style.display = row.textContent.toLowerCase().includes(q) ? '' : 'none';
  }});
}}

var _sortState = {{}};
function sortTable(tableId, colIdx) {{
  var table = document.getElementById(tableId);
  var tbody = table.querySelector('tbody');
  var rows = Array.from(tbody.querySelectorAll('tr'));
  var key = tableId + ':' + colIdx;
  var asc = !_sortState[key];
  _sortState[key] = asc;

  rows.sort(function(a, b) {{
    var av = a.cells[colIdx] ? a.cells[colIdx].textContent.trim() : '';
    var bv = b.cells[colIdx] ? b.cells[colIdx].textContent.trim() : '';
    var an = parseFloat(av), bn = parseFloat(bv);
    if (!isNaN(an) && !isNaN(bn)) return asc ? an - bn : bn - an;
    return asc ? av.localeCompare(bv) : bv.localeCompare(av);
  }});
  rows.forEach(function(r) {{ tbody.appendChild(r); }});

  // Update arrow indicators
  table.querySelectorAll('th .sort-arrow').forEach(function(el, i) {{
    el.textContent = i === colIdx ? (asc ? '▲' : '▼') : '⇅';
  }});
}}
</script>
</body>
</html>"""

    HTML_PATH.write_text(html, encoding="utf-8")
    print(f"HTML report written to {HTML_PATH} ({len(sheets)} sheet(s), map={'yes' if has_map else 'no'})")
    return 0


def df_to_summary_html(df: pd.DataFrame) -> str:
    """Render the Summary sheet as styled HTML sections rather than a raw table."""
    if df.empty:
        return "<p style='padding:16px;color:#666'>No summary data.</p>"

    band_colors = {"< 1 mile": "#FFCCCC", "1–2 miles": "#FFE5B4", "> 2 miles": "#CCFFCC"}

    # The Summary sheet has alternating header/data blocks; parse into sections
    sections: list[tuple[str, list[tuple[str, str]]]] = []  # (section_title, [(label, count)])
    current_section: str = ""
    current_sub: str = ""
    current_rows: list[tuple[str, str]] = []

    for _, row in df.iterrows():
        vals = [str(v).strip() if pd.notna(v) else "" for v in row]
        non_empty = [v for v in vals if v]
        if not non_empty:
            continue
        first = vals[0] if vals else ""
        second = vals[1] if len(vals) > 1 else ""

        # Detect section heading (single non-numeric value in col A, col B empty)
        if first and not second and not any(c.isdigit() for c in second):
            if "Distance" in first or "Property" in first or "Organization" in first or "Band" in first:
                if current_sub and current_rows:
                    sections.append((current_section + " — " + current_sub, current_rows))
                current_sub = first
                current_rows = []
            else:
                if current_sub and current_rows:
                    sections.append((current_section + " — " + current_sub, current_rows))
                current_section = first
                current_sub = ""
                current_rows = []
        elif first and second:
            current_rows.append((first, second))

    if current_sub and current_rows:
        sections.append((current_section + " — " + current_sub, current_rows))

    if not sections:
        # Fallback: just render a basic table
        header = "".join(f"<th>{html_mod.escape(str(c))}</th>" for c in df.columns)
        rows_html = ""
        for _, row in df.iterrows():
            cells = "".join(f"<td>{html_mod.escape(str(v) if pd.notna(v) else '')}</td>" for v in row)
            rows_html += f"<tr>{cells}</tr>"
        return (
            f"<div style='padding:16px'><table class='summary-table'>"
            f"<thead><tr>{header}</tr></thead><tbody>{rows_html}</tbody></table></div>"
        )

    out = "<div class='summary-wrap'>"
    for title, rows in sections:
        out += f"<div class='summary-section'><h2>{html_mod.escape(title)}</h2>"
        out += "<table class='summary-table'><tbody>"
        for label, count in rows:
            bg = band_colors.get(label, "")
            style = f' style="background:{bg}"' if bg else ""
            out += f"<tr{style}><td>{html_mod.escape(label)}</td><td><strong>{html_mod.escape(count)}</strong></td></tr>"
        out += "</tbody></table></div>"
    out += "</div>"
    return out


if __name__ == "__main__":
    sys.exit(main())
