"""Builds a self-contained Sankey-style HTML visualization of one day's reconciliation flow.

Hand-rolled inline SVG (no external JS/CDN) so the output is a fully portable, offline
file that can be attached to an email and opened anywhere. Colored using the fixed
status palette (good/warning/serious/critical) from the `dataviz` skill, since every
category here is literally a status - each with an icon + label so color is never the
only signal. Light/dark mode via prefers-color-scheme / data-theme, per the same skill.
"""
from __future__ import annotations

import html
from datetime import date
from pathlib import Path

import pandas as pd

from .config import Config

# Status palette (fixed, from the dataviz skill's reference palette - never themed).
COLOR_GOOD = "#0ca30c"
COLOR_WARNING = "#fab219"
COLOR_SERIOUS = "#ec835a"
COLOR_CRITICAL = "#d03b3b"
COLOR_MUTED = "#898781"

NODE_WIDTH = 22
COLUMN_GAP = 230
NODE_GAP = 10
MARGIN_TOP = 70
MARGIN_LEFT = 40
PLOT_HEIGHT = 520


def _counts_from_summary(summary_df: pd.DataFrame) -> dict:
    counts = dict(zip(summary_df["metric"], summary_df["count"]))
    total_ice2 = int(counts.get("TOTAL_ICE2_ROWS", 0))
    missing_mapping = int(counts.get("ICE2_MISSING_API_MAPPING", 0))
    mismatch = int(counts.get("ICE2_API_STATUS_MISMATCH", 0))
    missing_recent = int(counts.get("MISSING_FROM_HUBSCAPE", 0))
    missing_historic = int(counts.get("MISSING_FROM_HUBSCAPE_HISTORIC", 0))
    legit_absent = int(counts.get("LEGITIMATELY_ABSENT_FROM_HUBSCAPE", 0))
    orphans = int(counts.get("ORPHAN_IN_HUBSCAPE_NOT_IN_ICE2", 0))
    # Directly-computed (not derived by subtraction) from reconcile.reconcile - see
    # ReconciliationCompletenessError, which guarantees these are self-consistent with the
    # counts above before this function ever runs.
    matched_ok = int(counts.get("ICE2_MATCHED_OK", 0))
    missing_api_id = int(counts.get("HUBSCAPE_MISSING_API_ID", 0))

    # mapped is an intermediate quantity only (not independently asserted); match/visible use
    # the directly-computed ICE2_MATCHED_OK instead of subtracting through mismatch/missing/
    # legit_absent, so the diagram reads the same numbers summary.csv reports.
    mapped = total_ice2 - missing_mapping
    match = matched_ok
    visible = matched_ok

    return {
        "total_ice2": total_ice2, "missing_mapping": missing_mapping, "mapped": mapped,
        "mismatch": mismatch, "match": match,
        "missing_recent": missing_recent, "missing_historic": missing_historic,
        "legit_absent": legit_absent, "visible": max(visible, 0), "orphans": orphans,
        "in_hubscape_total": max(visible, 0) + orphans, "missing_api_id": missing_api_id,
    }


def _node(node_id, label, column, value, color, icon):
    return {"id": node_id, "label": label, "column": column, "value": value,
            "color": color, "icon": icon}


def _build_nodes_and_links(c: dict) -> tuple[list[dict], list[dict]]:
    nodes = [
        _node("ice2", "ICe2 Visits", 0, c["total_ice2"], COLOR_MUTED, "●"),
        _node("missing_mapping", "Missing API Mapping", 1, c["missing_mapping"], COLOR_SERIOUS, "⚠"),
        _node("mapped", "Mapped to API", 1, c["mapped"], COLOR_MUTED, "●"),
        _node("mismatch", "Status Mismatch", 2, c["mismatch"], COLOR_WARNING, "⚠"),
        _node("match", "Status Match", 2, c["match"], COLOR_MUTED, "●"),
        _node("orphan_src", "Untraced in ICe2", 2, c["orphans"], COLOR_CRITICAL, "✖"),
        _node("missing_recent", "Missing from Hubscape", 3, c["missing_recent"], COLOR_SERIOUS, "⚠"),
        _node("missing_historic", "Missing (Historic)", 3, c["missing_historic"], COLOR_MUTED, "●"),
        _node("legit_absent", "Legitimately Absent", 3, c["legit_absent"], COLOR_MUTED, "●"),
        _node("in_hubscape", "Visible in Hubscape", 3, c["in_hubscape_total"], COLOR_GOOD, "✓"),
    ]
    links = [
        {"source": "ice2", "target": "missing_mapping", "value": c["missing_mapping"], "color": COLOR_SERIOUS},
        {"source": "ice2", "target": "mapped", "value": c["mapped"], "color": COLOR_MUTED},
        {"source": "mapped", "target": "mismatch", "value": c["mismatch"], "color": COLOR_WARNING},
        {"source": "mapped", "target": "match", "value": c["match"], "color": COLOR_MUTED},
        {"source": "match", "target": "missing_recent", "value": c["missing_recent"], "color": COLOR_SERIOUS},
        {"source": "match", "target": "missing_historic", "value": c["missing_historic"], "color": COLOR_MUTED},
        {"source": "match", "target": "legit_absent", "value": c["legit_absent"], "color": COLOR_MUTED},
        {"source": "match", "target": "in_hubscape", "value": c["visible"], "color": COLOR_GOOD},
        {"source": "orphan_src", "target": "in_hubscape", "value": c["orphans"], "color": COLOR_CRITICAL},
    ]
    # Drop zero-value nodes/links so an empty category doesn't clutter the diagram.
    nodes = [n for n in nodes if n["value"] > 0]
    valid_ids = {n["id"] for n in nodes}
    links = [link_ for link_ in links if link_["value"] > 0 and link_["source"] in valid_ids and link_["target"] in valid_ids]
    return nodes, links


def _layout(nodes: list[dict], links: list[dict]) -> tuple[list[dict], list[dict], int]:
    scale_denominator = max(
        sum(n["value"] for n in nodes if n["column"] == col)
        for col in {n["column"] for n in nodes}
    ) or 1
    px_per_unit = PLOT_HEIGHT / scale_denominator

    columns: dict[int, list[dict]] = {}
    for n in nodes:
        columns.setdefault(n["column"], []).append(n)

    positioned = {}
    for col, col_nodes in columns.items():
        y = MARGIN_TOP
        for n in col_nodes:
            height = max(n["value"] * px_per_unit, 3)
            x = MARGIN_LEFT + col * COLUMN_GAP
            positioned[n["id"]] = {**n, "x": x, "y": y, "height": height}
            y += height + NODE_GAP

    # Stack each node's outgoing links (in the order given) within its height, and each
    # node's incoming links within its height too, so multi-link nodes render as clean
    # contiguous bands rather than overlapping.
    out_offset = {nid: 0.0 for nid in positioned}
    in_offset = {nid: 0.0 for nid in positioned}
    laid_out_links = []
    for link_ in links:
        src, tgt = positioned[link_["source"]], positioned[link_["target"]]
        band = max(link_["value"] * px_per_unit, 2)
        sy = src["y"] + out_offset[link_["source"]] + band / 2
        ty = tgt["y"] + in_offset[link_["target"]] + band / 2
        out_offset[link_["source"]] += band
        in_offset[link_["target"]] += band
        laid_out_links.append({**link_, "sx": src["x"] + NODE_WIDTH, "sy": sy,
                                "tx": tgt["x"], "ty": ty, "band": band})

    max_col = max(n["column"] for n in nodes)
    total_width = MARGIN_LEFT + max_col * COLUMN_GAP + NODE_WIDTH + 220
    return list(positioned.values()), laid_out_links, total_width


def _render_svg(nodes: list[dict], links: list[dict], width: int) -> str:
    parts = []
    for link_ in links:
        sx, sy, tx, ty, band = link_["sx"], link_["sy"], link_["tx"], link_["ty"], link_["band"]
        mx = (sx + tx) / 2
        path = f"M {sx},{sy - band/2} C {mx},{sy - band/2} {mx},{ty - band/2} {tx},{ty - band/2} " \
               f"L {tx},{ty + band/2} C {mx},{ty + band/2} {mx},{sy + band/2} {sx},{sy + band/2} Z"
        title = html.escape(f"{link_['value']:,}")
        parts.append(
            f'<path d="{path}" fill="{link_["color"]}" opacity="0.55" class="viz-link">'
            f'<title>{title}</title></path>'
        )
    for n in nodes:
        label = html.escape(n["label"])
        value = f"{n['value']:,}"
        parts.append(
            f'<g class="viz-node">'
            f'<rect x="{n["x"]}" y="{n["y"]}" width="{NODE_WIDTH}" height="{n["height"]:.1f}" '
            f'rx="3" fill="{n["color"]}"><title>{html.escape(label)}: {value}</title></rect>'
            f'<text x="{n["x"] + NODE_WIDTH + 8}" y="{n["y"] + n["height"]/2 - 6:.1f}" class="viz-node-label">'
            f'{n["icon"]} {label}</text>'
            f'<text x="{n["x"] + NODE_WIDTH + 8}" y="{n["y"] + n["height"]/2 + 12:.1f}" class="viz-node-value">'
            f'{value}</text>'
            f'</g>'
        )
    height = PLOT_HEIGHT + MARGIN_TOP + 20
    return f'<svg viewBox="0 0 {width} {height}" width="100%" style="max-width:{width}px">' + "".join(parts) + "</svg>"


def _render_issues_bar_chart(c: dict) -> str:
    """A separate, properly-scaled comparison of just today's issue counts.

    The Sankey below is honest about true proportions (so the historic/completed mass -
    typically 10-100x bigger than any issue - renders as a thin sliver, per the dataviz
    skill's own 'compare magnitude' guidance a bar chart, not a forced non-linear Sankey
    scale, is the right form for comparing these smaller figures against each other.
    """
    bars = [
        ("Missing from Hubscape", c["missing_recent"], COLOR_SERIOUS, "serious"),
        ("Missing API mapping", c["missing_mapping"], COLOR_SERIOUS, "serious"),
        ("Orphans in Hubscape", c["orphans"], COLOR_CRITICAL, "critical"),
        ("Status mismatch", c["mismatch"], COLOR_WARNING, "warning"),
        ("Missing from Hubscape (historic)", c["missing_historic"], COLOR_MUTED, "muted"),
    ]
    bars = [b for b in bars if b[1] > 0]
    bars.sort(key=lambda b: b[1], reverse=True)
    if not bars:
        return "<p>No issues found in this run.</p>"

    max_val = max(b[1] for b in bars)
    bar_area_width = 420
    row_height = 30
    label_width = 210
    svg_width = label_width + bar_area_width + 90
    svg_height = row_height * len(bars) + 10

    rows = []
    for i, (label, value, color, role) in enumerate(bars):
        y = i * row_height + 6
        bar_w = max((value / max_val) * bar_area_width, 3)
        rows.append(
            f'<g class="viz-bar-row">'
            f'<text x="{label_width - 10}" y="{y + 15}" text-anchor="end" class="viz-node-label">{html.escape(label)}</text>'
            f'<rect x="{label_width}" y="{y}" width="{bar_w:.1f}" height="18" rx="2" fill="{color}">'
            f'<title>{html.escape(label)}: {value:,}</title></rect>'
            f'<text x="{label_width + bar_w + 8:.1f}" y="{y + 14}" class="viz-node-value">{value:,}</text>'
            f'</g>'
        )
    svg = (
        f'<svg viewBox="0 0 {svg_width} {svg_height}" width="100%" style="max-width:{svg_width}px">'
        + "".join(rows) + "</svg>"
    )
    return svg


def _render_table(c: dict) -> str:
    rows = [
        ("ICe2 active visits (total)", c["total_ice2"], "muted"),
        ("Missing API mapping", c["missing_mapping"], "serious"),
        ("Status mismatch (ICe2 vs Visits API)", c["mismatch"], "warning"),
        ("Missing from Hubscape (recent)", c["missing_recent"], "serious"),
        ("Missing from Hubscape (historic, informational)", c["missing_historic"], "muted"),
        ("Legitimately absent (completed/cancelled, informational)", c["legit_absent"], "muted"),
        ("Visible in Hubscape, correctly", c["visible"], "good"),
        ("Orphans in Hubscape (untraced in ICe2)", c["orphans"], "critical"),
        ("Hubscape rows missing API_ID (informational)", c["missing_api_id"], "muted"),
    ]
    body = "".join(
        f'<tr><td>{html.escape(label)}</td><td class="viz-num">{value:,}</td>'
        f'<td><span class="viz-pill viz-pill-{role}"></span>{role}</td></tr>'
        for label, value, role in rows
    )
    return f'<table class="viz-table"><thead><tr><th>Category</th><th>Count</th><th>Status</th></tr></thead><tbody>{body}</tbody></table>'


_PAGE_TEMPLATE = """<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Visit Reconciliation Flow - {run_date}</title>
<style>
  .viz-root {{
    color-scheme: light;
    --surface-1: #fcfcfb;
    --text-primary: #0b0b0b;
    --text-secondary: #52514e;
    --muted: #898781;
    --gridline: #e1e0d9;
  }}
  @media (prefers-color-scheme: dark) {{
    :root:not([data-theme="light"]) .viz-root {{
      color-scheme: dark;
      --surface-1: #1a1a19;
      --text-primary: #ffffff;
      --text-secondary: #c3c2b7;
      --muted: #898781;
      --gridline: #2c2c2a;
    }}
  }}
  :root[data-theme="dark"] .viz-root {{
    color-scheme: dark;
    --surface-1: #1a1a19;
    --text-primary: #ffffff;
    --text-secondary: #c3c2b7;
    --muted: #898781;
    --gridline: #2c2c2a;
  }}
  body {{ margin: 0; background: var(--surface-1); font-family: system-ui, -apple-system, "Segoe UI", sans-serif; }}
  .viz-root {{ background: var(--surface-1); color: var(--text-primary); padding: 24px; }}
  h1 {{ font-size: 18px; font-weight: 600; margin: 0 0 4px; }}
  .viz-subtitle {{ color: var(--text-secondary); font-size: 13px; margin: 0 0 20px; }}
  .viz-chart-wrap {{ overflow-x: auto; }}
  .viz-node-label {{ fill: var(--text-primary); font-size: 13px; }}
  .viz-node-value {{ fill: var(--text-secondary); font-size: 12px; font-variant-numeric: tabular-nums; }}
  .viz-link {{ transition: opacity 0.15s; }}
  .viz-link:hover {{ opacity: 0.85; }}
  .viz-legend {{ display: flex; gap: 16px; flex-wrap: wrap; margin: 16px 0 24px; font-size: 12px; color: var(--text-secondary); }}
  .viz-legend-item {{ display: flex; align-items: center; gap: 6px; }}
  .viz-pill {{ display: inline-block; width: 10px; height: 10px; border-radius: 50%; }}
  .viz-pill-good {{ background: {good}; }}
  .viz-pill-warning {{ background: {warning}; }}
  .viz-pill-serious {{ background: {serious}; }}
  .viz-pill-critical {{ background: {critical}; }}
  .viz-pill-muted {{ background: {muted}; }}
  table.viz-table {{ border-collapse: collapse; font-size: 13px; margin-top: 8px; width: 100%; max-width: 620px; }}
  table.viz-table th, table.viz-table td {{ text-align: left; padding: 6px 10px; border-bottom: 1px solid var(--gridline); }}
  table.viz-table .viz-num {{ text-align: right; font-variant-numeric: tabular-nums; }}
  table.viz-table th {{ color: var(--text-secondary); font-weight: 500; }}
  h2 {{ font-size: 14px; font-weight: 600; margin: 24px 0 8px; }}
</style>
</head>
<body>
<div class="viz-root">
  <h1>Visit Reconciliation Flow</h1>
  <p class="viz-subtitle">{run_date} - ICe2 &rarr; Visits API &rarr; Hubscape</p>
  <div class="viz-legend">
    <span class="viz-legend-item"><span class="viz-pill viz-pill-good"></span>Good / visible</span>
    <span class="viz-legend-item"><span class="viz-pill viz-pill-warning"></span>Warning</span>
    <span class="viz-legend-item"><span class="viz-pill viz-pill-serious"></span>Serious / actionable gap</span>
    <span class="viz-legend-item"><span class="viz-pill viz-pill-critical"></span>Critical / untraced</span>
    <span class="viz-legend-item"><span class="viz-pill viz-pill-muted"></span>Informational / neutral</span>
  </div>
  <h2>Today's issues at a glance</h2>
  <p class="viz-subtitle">Scaled to compare issue counts against each other - see the full-scale flow below for how these sit within the whole ICe2 population.</p>
  <div class="viz-chart-wrap">{issues_svg}</div>

  <h2>Full reconciliation flow</h2>
  <p class="viz-subtitle">True-to-scale - historic completed/cancelled visits (correctly absent from Hubscape) dominate ICe2's overall history, so smaller issue bands may render thin here; see the chart above for those compared on their own scale.</p>
  <div class="viz-chart-wrap">{svg}</div>

  <h2>Table view</h2>
  {table}
</div>
</body>
</html>
"""


def build_sankey_html(summary_df: pd.DataFrame, cfg: Config, run_date: date, output_dir: Path) -> Path:
    counts = _counts_from_summary(summary_df)
    nodes, links = _build_nodes_and_links(counts)
    positioned_nodes, laid_out_links, width = _layout(nodes, links)
    svg = _render_svg(positioned_nodes, laid_out_links, width)
    issues_svg = _render_issues_bar_chart(counts)
    table = _render_table(counts)

    page = _PAGE_TEMPLATE.format(
        run_date=run_date.isoformat(), svg=svg, issues_svg=issues_svg, table=table,
        good=COLOR_GOOD, warning=COLOR_WARNING, serious=COLOR_SERIOUS,
        critical=COLOR_CRITICAL, muted=COLOR_MUTED,
    )

    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / "sankey.html"
    path.write_text(page, encoding="utf-8")
    return path
