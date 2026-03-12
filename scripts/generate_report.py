#!/usr/bin/env python3
"""Generate docs/index.html from docs/latest.xlsx."""

import html as html_mod
import sys
from pathlib import Path

import pandas as pd

XLSX_PATH = Path("docs/latest.xlsx")
TIMESTAMP_PATH = Path("docs/last_updated.txt")
HTML_PATH = Path("docs/index.html")

DIST_COL_MAIN = "Distance to Nearest Property (miles)"
DIST_COL_MATCH = "Distance to Property (miles)"

# Map sheet names to user-friendly tab labels
TAB_LABELS = {
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


def df_to_html_table(df: pd.DataFrame, dist_col: str) -> str:
    if df.empty:
        return "<p style='padding:16px;color:#666'>No events found.</p>"

    header_cells = "".join(
        f"<th>{html_mod.escape(str(c))}</th>" for c in df.columns
    )

    rows: list[str] = []
    for _, row in df.iterrows():
        bg = row_bg(row.get(dist_col))
        is_dup = row.get("Is Duplicate", False)
        is_dup = is_dup is True or str(is_dup).upper() == "TRUE"

        styles: list[str] = []
        if bg:
            styles.append(f"background:{bg}")
        if is_dup:
            styles.append("color:#999")
            styles.append("font-style:italic")

        row_style = f' style="{";".join(styles)}"' if styles else ""

        cells = "".join(
            f"<td>{html_mod.escape(str(v) if pd.notna(v) else '')}</td>"
            for v in row
        )
        rows.append(f"<tr{row_style}>{cells}</tr>")

    return (
        f"<div class='table-wrap'>"
        f"<table>"
        f"<thead><tr>{header_cells}</tr></thead>"
        f"<tbody>{''.join(rows)}</tbody>"
        f"</table></div>"
    )


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
        sheets[name] = pd.read_excel(xl, sheet_name=name)

    # Build tab buttons and content panels
    tab_buttons: list[str] = []
    tab_panels: list[str] = []
    for i, (name, df) in enumerate(sheets.items()):
        label = TAB_LABELS.get(name, name)
        active_cls = " active" if i == 0 else ""
        hidden_attr = "" if i == 0 else " hidden"
        dist_col = DIST_COL_MAIN if DIST_COL_MAIN in df.columns else DIST_COL_MATCH
        count = len(df)

        tab_buttons.append(
            f'<button class="tab-btn{active_cls}" onclick="showTab(this,\'{name}\')">'
            f'{html_mod.escape(label)} <span class="count">({count})</span></button>'
        )
        tab_panels.append(
            f'<div id="tab-{name}" class="tab-panel"{hidden_attr}>'
            f"{df_to_html_table(df, dist_col)}</div>"
        )

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
  .table-wrap {{ overflow-x: auto; max-height: 75vh; overflow-y: auto; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 0.82em; }}
  thead tr {{ position: sticky; top: 0; z-index: 1; }}
  th {{
    background: #1e2a3a; color: #fff; padding: 8px 11px;
    text-align: left; white-space: nowrap; font-weight: 600;
  }}
  td {{ padding: 5px 11px; border-bottom: 1px solid #e8e8e8; white-space: nowrap; }}
  tbody tr:hover td {{ filter: brightness(0.93); }}
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
  <div class="legend-item">
    <div class="swatch" style="background:#FFCCCC"></div> &lt;1 mile
  </div>
  <div class="legend-item">
    <div class="swatch" style="background:#FFE5B4"></div> 1&ndash;2 miles
  </div>
  <div class="legend-item">
    <div class="swatch" style="background:#CCFFCC"></div> &gt;2 miles
  </div>
  <div class="legend-item" style="color:#999;font-style:italic">
    italic gray = duplicate event
  </div>
</div>
<div class="tabs">{"".join(tab_buttons)}</div>
{"".join(tab_panels)}
<script>
function showTab(btn, name) {{
  document.querySelectorAll('.tab-panel').forEach(el => {{ el.hidden = true; }});
  document.querySelectorAll('.tab-btn').forEach(el => {{ el.classList.remove('active'); }});
  document.getElementById('tab-' + name).hidden = false;
  btn.classList.add('active');
}}
</script>
</body>
</html>"""

    HTML_PATH.write_text(html, encoding="utf-8")
    print(f"HTML report written to {HTML_PATH} ({len(sheets)} sheet(s))")
    return 0


if __name__ == "__main__":
    sys.exit(main())
