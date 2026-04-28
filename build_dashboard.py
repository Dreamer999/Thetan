#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "openpyxl>=3.1",
#   "matplotlib>=3.8",
# ]
# ///
"""
Read Financial_Model.xlsx and emit a single, fully self-contained
dashboard.html that runs offline on any device. FinTech-style red
palette, fully responsive, no horizontal scroll, no external assets.
"""

from __future__ import annotations

import io
import json
import re
from html import escape
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib import font_manager
import openpyxl

# ── theme ────────────────────────────────────────────────────────────────────

RED_PRIMARY   = "#E63946"
RED_LIGHT     = "#FF6B73"
RED_DEEPER    = "#B91C2C"
RED_DARKEST   = "#7F1D1D"
RED_GLOW      = "#FF3B47"
ACCENT_GOLD   = "#F4B942"

BG_PAGE       = "#0A0E1A"
BG_CARD       = "#13182A"
BG_CARD_ALT   = "#1A2138"
BORDER        = "#252B3F"
TEXT_PRIMARY  = "#F1F5F9"
TEXT_MUTED    = "#94A3B8"
TEXT_DIM      = "#64748B"
GRID_LINE     = "#1E2538"

CHART_PALETTE = [RED_PRIMARY, RED_LIGHT, RED_DEEPER, ACCENT_GOLD, "#9B7DAB"]

# Use a plain sans-serif that's likely on every device
plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["axes.unicode_minus"] = False


# ── data extraction ──────────────────────────────────────────────────────────

def col_idx(letter: str) -> int:
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch.upper()) - ord("A") + 1)
    return n


def row_values(ws, row: int, start_col: int, end_col: int) -> list:
    return [ws.cell(row=row, column=c).value for c in range(start_col, end_col + 1)]


def read_model(xlsx_path: Path) -> dict:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    cover = wb["Cover"]
    meta = {
        "title":       cover["A1"].value,
        "version":     cover["B3"].value,
        "status":      cover["B4"].value,
        "date":        cover["B5"].value,
        "author":      cover["B6"].value,
        "currency":    cover["B7"].value,
        "fiscal_year": cover["B8"].value,
    }

    # period headers row 2, columns B..I
    rev = wb["Revenue"]
    periods = [c for c in row_values(rev, 2, 2, 9) if c]
    n = len(periods)

    revenue = {
        "saas":         row_values(rev, 5, 2, 1 + n),
        "services":     row_values(rev, 6, 2, 1 + n),
        "marketplace":  row_values(rev, 7, 2, 1 + n),
        "total":        row_values(rev, 9, 2, 1 + n),
    }

    pl = wb["P&L"]
    income = {
        "revenue":      row_values(pl, 4,  2, 1 + n),
        "cogs":         row_values(pl, 5,  2, 1 + n),
        "gross_profit": row_values(pl, 6,  2, 1 + n),
        "rnd":          row_values(pl, 9,  2, 1 + n),
        "sm":           row_values(pl, 10, 2, 1 + n),
        "ga":           row_values(pl, 11, 2, 1 + n),
        "ebitda":       row_values(pl, 12, 2, 1 + n),
        "ebit":         row_values(pl, 14, 2, 1 + n),
        "net_income":   row_values(pl, 16, 2, 1 + n),
    }

    cf = wb["Cash Flow"]
    cashflow = {
        "net_income": row_values(cf, 4,  2, 1 + n),
        "depn":       row_values(cf, 5,  2, 1 + n),
        "wc":         row_values(cf, 6,  2, 1 + n),
        "cfo":        row_values(cf, 7,  2, 1 + n),
        "capex":      row_values(cf, 8,  2, 1 + n),
        "fcf":        row_values(cf, 9,  2, 1 + n),
        "cash_bal":   row_values(cf, 11, 2, 1 + n),
    }

    val = wb["Valuation"]
    # row indices found from build script:
    valuation = {
        "pv_fcfs":        val.cell(row=4,  column=2).value,
        "pv_terminal":    val.cell(row=5,  column=2).value,
        "ev_dcf":         val.cell(row=6,  column=2).value,
        "exit_revenue":   val.cell(row=9,  column=2).value,
        "ev_rev_mult":    val.cell(row=11, column=2).value,
        "ev_low":         val.cell(row=14, column=2).value,
        "ev_high":        val.cell(row=15, column=2).value,
    }

    # Determine where actuals end vs estimates
    actuals_count = sum(1 for p in periods if isinstance(p, str) and p.endswith("A"))

    return {
        "meta":      meta,
        "periods":   periods,
        "actuals_count": actuals_count,
        "revenue":   revenue,
        "income":    income,
        "cashflow":  cashflow,
        "valuation": valuation,
    }


# ── chart helpers ────────────────────────────────────────────────────────────

def style_axes(ax):
    ax.set_facecolor("none")
    ax.tick_params(colors=TEXT_MUTED, labelsize=10)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_color(BORDER)
        ax.spines[s].set_linewidth(0.8)
    ax.yaxis.grid(True, color=GRID_LINE, linewidth=0.6, linestyle="-", zorder=0)
    ax.set_axisbelow(True)


def fmt_thousands(x, _pos):
    if abs(x) >= 1_000:
        return f"{x/1000:.1f}M"
    return f"{int(x):,}"


def fig_to_svg(fig) -> str:
    buf = io.StringIO()
    fig.savefig(buf, format="svg", transparent=True, bbox_inches="tight", pad_inches=0.05)
    plt.close(fig)
    svg = buf.getvalue()

    # Strip XML / DOCTYPE preamble; keep only the <svg>…</svg> element
    svg = re.sub(r"<\?xml[^>]+\?>", "", svg)
    svg = re.sub(r"<!DOCTYPE[^>]+>", "", svg)
    svg = re.sub(r"<!--.*?-->", "", svg, flags=re.DOTALL)

    # Make SVG fully responsive: drop fixed width/height, keep viewBox
    svg = re.sub(r'(<svg\b[^>]*?)\swidth="[^"]+"', r"\1", svg, count=1)
    svg = re.sub(r'(<svg\b[^>]*?)\sheight="[^"]+"', r"\1", svg, count=1)
    svg = svg.replace(
        "<svg ",
        '<svg preserveAspectRatio="xMidYMid meet" style="width:100%;height:auto;display:block;" ',
        1,
    )
    return svg.strip()


# ── individual charts ────────────────────────────────────────────────────────

def chart_revenue_trajectory(data: dict) -> str:
    periods = data["periods"]
    rev = data["revenue"]["total"]
    cutoff = data["actuals_count"]

    fig, ax = plt.subplots(figsize=(8.5, 4.0))
    style_axes(ax)

    actual_x = list(range(cutoff))
    est_x    = list(range(cutoff - 1, len(periods)))

    ax.plot(actual_x, rev[:cutoff],
            color=RED_PRIMARY, linewidth=3.0, marker="o",
            markersize=8, markerfacecolor=RED_PRIMARY,
            markeredgecolor=BG_PAGE, markeredgewidth=2,
            label="Actuals", zorder=3)
    ax.plot(est_x, rev[cutoff - 1:],
            color=RED_LIGHT, linewidth=2.5, linestyle="--", marker="s",
            markersize=7, markerfacecolor=RED_LIGHT,
            markeredgecolor=BG_PAGE, markeredgewidth=2,
            label="Estimates", zorder=3)

    # area fill
    ax.fill_between(range(len(periods)), rev,
                    color=RED_PRIMARY, alpha=0.10, zorder=1)

    ax.set_xticks(range(len(periods)))
    ax.set_xticklabels(periods, rotation=0)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(fmt_thousands))
    ax.set_ylabel("Revenue ($K)", color=TEXT_MUTED, fontsize=11)

    leg = ax.legend(loc="upper left", frameon=False, fontsize=10)
    for txt in leg.get_texts():
        txt.set_color(TEXT_PRIMARY)

    fig.tight_layout()
    return fig_to_svg(fig)


def chart_segment_stacked(data: dict) -> str:
    periods = data["periods"]
    saas         = data["revenue"]["saas"]
    services     = data["revenue"]["services"]
    marketplace  = data["revenue"]["marketplace"]

    fig, ax = plt.subplots(figsize=(8.5, 4.0))
    style_axes(ax)

    x = range(len(periods))
    bar_w = 0.62
    ax.bar(x, saas,        bar_w, color=RED_PRIMARY, label="SaaS Subscriptions", zorder=2)
    ax.bar(x, services,    bar_w, bottom=saas, color=RED_LIGHT,
           label="Professional Services", zorder=2)
    ax.bar(x, marketplace, bar_w,
           bottom=[a + b for a, b in zip(saas, services)],
           color=ACCENT_GOLD, label="Marketplace & Other", zorder=2)

    ax.set_xticks(list(x))
    ax.set_xticklabels(periods)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(fmt_thousands))
    ax.set_ylabel("Revenue ($K)", color=TEXT_MUTED, fontsize=11)

    leg = ax.legend(loc="upper left", frameon=False, fontsize=10, ncol=3)
    for txt in leg.get_texts():
        txt.set_color(TEXT_PRIMARY)

    fig.tight_layout()
    return fig_to_svg(fig)


def chart_margin_trend(data: dict) -> str:
    periods = data["periods"]
    rev = data["income"]["revenue"]
    gp  = data["income"]["gross_profit"]
    eb  = data["income"]["ebitda"]
    ni  = data["income"]["net_income"]

    gm = [g / r for g, r in zip(gp, rev)]
    em = [e / r for e, r in zip(eb, rev)]
    nm = [n / r for n, r in zip(ni, rev)]

    fig, ax = plt.subplots(figsize=(8.5, 4.0))
    style_axes(ax)

    x = list(range(len(periods)))
    ax.plot(x, gm, color=RED_PRIMARY, linewidth=2.5, marker="o", markersize=6,
            markeredgecolor=BG_PAGE, markeredgewidth=1.5, label="Gross Margin")
    ax.plot(x, em, color=ACCENT_GOLD, linewidth=2.5, marker="s", markersize=6,
            markeredgecolor=BG_PAGE, markeredgewidth=1.5, label="EBITDA Margin")
    ax.plot(x, nm, color=RED_LIGHT, linewidth=2.5, marker="^", markersize=6,
            markeredgecolor=BG_PAGE, markeredgewidth=1.5, label="Net Margin")

    ax.set_xticks(x)
    ax.set_xticklabels(periods)
    ax.yaxis.set_major_formatter(mticker.PercentFormatter(xmax=1, decimals=0))
    ax.set_ylabel("Margin", color=TEXT_MUTED, fontsize=11)

    leg = ax.legend(loc="lower right", frameon=False, fontsize=10)
    for txt in leg.get_texts():
        txt.set_color(TEXT_PRIMARY)

    fig.tight_layout()
    return fig_to_svg(fig)


def chart_fcf_bridge(data: dict) -> str:
    periods = data["periods"]
    cfo   = data["cashflow"]["cfo"]
    capex = data["cashflow"]["capex"]
    fcf   = data["cashflow"]["fcf"]

    fig, ax = plt.subplots(figsize=(8.5, 4.0))
    style_axes(ax)

    x = list(range(len(periods)))
    bar_w = 0.28
    ax.bar([i - bar_w for i in x], cfo,   bar_w, color=RED_PRIMARY, label="Cash from Ops", zorder=2)
    ax.bar(x,                       capex, bar_w, color=RED_DEEPER, label="CapEx", zorder=2)
    ax.bar([i + bar_w for i in x], fcf,   bar_w, color=ACCENT_GOLD, label="Free Cash Flow", zorder=2)

    ax.axhline(0, color=BORDER, linewidth=0.8)
    ax.set_xticks(x)
    ax.set_xticklabels(periods)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(fmt_thousands))
    ax.set_ylabel("Cash ($K)", color=TEXT_MUTED, fontsize=11)

    leg = ax.legend(loc="upper left", frameon=False, fontsize=10, ncol=3)
    for txt in leg.get_texts():
        txt.set_color(TEXT_PRIMARY)

    fig.tight_layout()
    return fig_to_svg(fig)


def chart_valuation(data: dict) -> str:
    v = data["valuation"]
    labels = ["DCF", "Revenue\nMultiple", "Range\n(Low → High)"]
    lows  = [v["ev_dcf"], v["ev_rev_mult"], v["ev_low"]]
    highs = [v["ev_dcf"], v["ev_rev_mult"], v["ev_high"]]
    colors = [RED_PRIMARY, ACCENT_GOLD, RED_LIGHT]

    fig, ax = plt.subplots(figsize=(8.5, 4.0))
    style_axes(ax)

    x = list(range(len(labels)))
    for i, (lo, hi, c, lbl) in enumerate(zip(lows, highs, colors, labels)):
        if lo == hi:
            ax.bar(i, hi, 0.55, color=c, zorder=2)
            ax.text(i, hi, f"  ${hi:,.0f}K", ha="center", va="bottom",
                    color=TEXT_PRIMARY, fontsize=10, fontweight="bold")
        else:
            # Range bar
            ax.bar(i, hi - lo, 0.55, bottom=lo, color=c, alpha=0.55, zorder=2)
            ax.bar(i, 0, 0.7, bottom=lo, color=c, edgecolor=c, linewidth=2, zorder=3)
            ax.bar(i, 0, 0.7, bottom=hi, color=c, edgecolor=c, linewidth=2, zorder=3)
            ax.text(i, hi, f"  ${hi:,.0f}K", ha="center", va="bottom",
                    color=TEXT_PRIMARY, fontsize=10, fontweight="bold")
            ax.text(i, lo, f"  ${lo:,.0f}K", ha="center", va="top",
                    color=TEXT_MUTED, fontsize=9)

    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(fmt_thousands))
    ax.set_ylabel("Enterprise Value ($K)", color=TEXT_MUTED, fontsize=11)

    fig.tight_layout()
    return fig_to_svg(fig)


# ── KPI computation ──────────────────────────────────────────────────────────

def compute_kpis(data: dict) -> list[dict]:
    rev = data["income"]["revenue"]
    ni  = data["income"]["net_income"]
    eb  = data["income"]["ebitda"]
    fcf = data["cashflow"]["fcf"]
    v   = data["valuation"]

    # CAGR
    n = len(rev) - 1
    cagr = (rev[-1] / rev[0]) ** (1 / n) - 1 if rev[0] else 0
    last_em = eb[-1] / rev[-1] if rev[-1] else 0
    last_nm = ni[-1] / rev[-1] if rev[-1] else 0

    return [
        {"label": "FY2029E Revenue",   "value": f"${rev[-1]/1000:.1f}M",  "delta": f"+{cagr*100:.1f}% CAGR", "trend": "up"},
        {"label": "FY2029E EBITDA",    "value": f"${eb[-1]/1000:.1f}M",   "delta": f"{last_em*100:.1f}% margin", "trend": "up"},
        {"label": "FY2029E Net Income","value": f"${ni[-1]/1000:.1f}M",   "delta": f"{last_nm*100:.1f}% margin", "trend": "up"},
        {"label": "FY2029E Free Cash Flow","value": f"${fcf[-1]/1000:.1f}M",  "delta": "5-yr cumulative", "trend": "up"},
        {"label": "Enterprise Value (DCF)",  "value": f"${v['ev_dcf']/1000:.1f}M","delta": "@ 12% WACC", "trend": "neutral"},
        {"label": "EV / Revenue Multiple",   "value": f"${v['ev_rev_mult']/1000:.1f}M","delta": "@ 6.0x exit", "trend": "neutral"},
    ]


# ── HTML rendering ───────────────────────────────────────────────────────────

CSS = f"""
*, *::before, *::after {{ box-sizing: border-box; }}
html, body {{ margin: 0; padding: 0; }}
html {{ overflow-x: hidden; }}
body {{
    background: {BG_PAGE};
    color: {TEXT_PRIMARY};
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto,
                 "Helvetica Neue", Arial, "Noto Sans TC", "PingFang TC",
                 "Microsoft JhengHei", sans-serif;
    font-size: 15px;
    line-height: 1.55;
    overflow-x: hidden;
    min-height: 100vh;
    background-image:
        radial-gradient(1200px 600px at 0% -10%, rgba(230,57,70,0.10), transparent 60%),
        radial-gradient(900px 500px at 100% 0%, rgba(244,185,66,0.05), transparent 60%);
}}

.container {{
    width: 100%;
    max-width: 1400px;
    margin: 0 auto;
    padding: clamp(16px, 3vw, 32px);
}}

/* Header */
.header {{
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
    gap: 16px;
    padding-bottom: 24px;
    border-bottom: 1px solid {BORDER};
    margin-bottom: 28px;
}}
.brand {{ display: flex; align-items: center; gap: 14px; min-width: 0; }}
.brand-mark {{
    width: 44px; height: 44px;
    border-radius: 12px;
    background: linear-gradient(135deg, {RED_PRIMARY} 0%, {RED_DEEPER} 100%);
    display: grid; place-items: center;
    font-weight: 800; font-size: 18px; color: white;
    box-shadow: 0 6px 20px rgba(230,57,70,0.35);
    flex-shrink: 0;
}}
.brand-text {{ min-width: 0; }}
.brand-text h1 {{
    margin: 0;
    font-size: clamp(1.1rem, 2.4vw, 1.45rem);
    font-weight: 700;
    letter-spacing: -0.01em;
    overflow-wrap: anywhere;
}}
.brand-text p {{
    margin: 2px 0 0;
    color: {TEXT_MUTED};
    font-size: 0.85rem;
}}
.meta-pills {{
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
}}
.pill {{
    background: {BG_CARD};
    border: 1px solid {BORDER};
    color: {TEXT_MUTED};
    font-size: 0.75rem;
    padding: 6px 12px;
    border-radius: 999px;
    font-weight: 500;
    white-space: nowrap;
}}
.pill.live::before {{
    content: "";
    display: inline-block;
    width: 6px; height: 6px;
    background: {RED_GLOW};
    border-radius: 50%;
    margin-right: 6px;
    box-shadow: 0 0 10px {RED_GLOW};
    animation: pulse 1.6s infinite;
}}
@keyframes pulse {{
    0%, 100% {{ opacity: 1; }}
    50% {{ opacity: 0.4; }}
}}

/* KPI grid */
.kpi-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(min(220px, 100%), 1fr));
    gap: 14px;
    margin-bottom: 28px;
}}
.kpi {{
    background: {BG_CARD};
    border: 1px solid {BORDER};
    border-radius: 14px;
    padding: 18px 18px 16px;
    position: relative;
    overflow: hidden;
    transition: transform 0.18s ease, border-color 0.18s ease;
}}
.kpi:hover {{
    transform: translateY(-2px);
    border-color: {RED_DEEPER};
}}
.kpi::before {{
    content: "";
    position: absolute;
    top: 0; left: 0;
    width: 100%; height: 3px;
    background: linear-gradient(90deg, {RED_PRIMARY}, {RED_LIGHT});
}}
.kpi-label {{
    color: {TEXT_MUTED};
    font-size: 0.78rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin-bottom: 8px;
}}
.kpi-value {{
    font-size: clamp(1.3rem, 3vw, 1.75rem);
    font-weight: 700;
    color: {TEXT_PRIMARY};
    letter-spacing: -0.02em;
    overflow-wrap: anywhere;
}}
.kpi-delta {{
    margin-top: 8px;
    font-size: 0.78rem;
    color: {RED_LIGHT};
    font-weight: 500;
}}
.kpi-delta.neutral {{ color: {TEXT_MUTED}; }}

/* Cards */
.section-title {{
    font-size: 0.78rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: {TEXT_MUTED};
    margin: 32px 0 14px;
    padding-left: 10px;
    border-left: 3px solid {RED_PRIMARY};
}}
.card-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(min(420px, 100%), 1fr));
    gap: 18px;
}}
.card {{
    background: {BG_CARD};
    border: 1px solid {BORDER};
    border-radius: 14px;
    padding: 20px;
    overflow: hidden;
}}
.card.full {{ grid-column: 1 / -1; }}
.card-head {{
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    align-items: baseline;
    justify-content: space-between;
    margin-bottom: 12px;
}}
.card-title {{
    font-size: 1rem;
    font-weight: 600;
    color: {TEXT_PRIMARY};
    margin: 0;
}}
.card-sub {{
    font-size: 0.78rem;
    color: {TEXT_DIM};
}}
.chart-wrap {{
    width: 100%;
    overflow: hidden;
}}
.chart-wrap svg {{
    width: 100% !important;
    height: auto !important;
    max-width: 100%;
    display: block;
}}

/* Data table — responsive: cards on small screens */
.data-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.85rem;
    table-layout: fixed;
}}
.data-table th, .data-table td {{
    padding: 9px 8px;
    text-align: right;
    border-bottom: 1px solid {BORDER};
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}}
.data-table th:first-child, .data-table td:first-child {{
    text-align: left;
    color: {TEXT_MUTED};
    font-weight: 500;
}}
.data-table thead th {{
    color: {TEXT_DIM};
    font-size: 0.72rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    border-bottom: 1px solid {RED_DEEPER};
}}
.data-table tbody tr:hover {{
    background: rgba(230,57,70,0.04);
}}
.data-table .total-row td {{
    font-weight: 700;
    color: {TEXT_PRIMARY};
    border-top: 1px solid {RED_DEEPER};
}}
.table-scroll {{
    width: 100%;
    overflow-x: auto;
    -webkit-overflow-scrolling: touch;
}}

/* Footer */
.footer {{
    margin-top: 36px;
    padding-top: 20px;
    border-top: 1px solid {BORDER};
    text-align: center;
    color: {TEXT_DIM};
    font-size: 0.78rem;
}}

/* Responsive tweaks */
@media (max-width: 720px) {{
    .header {{ flex-direction: column; align-items: flex-start; }}
    .meta-pills {{ width: 100%; }}
    .data-table {{ font-size: 0.78rem; }}
    .data-table th, .data-table td {{ padding: 7px 5px; }}
}}
@media (max-width: 480px) {{
    .container {{ padding: 14px; }}
    .kpi {{ padding: 14px; }}
    .card {{ padding: 14px; }}
    .data-table {{ font-size: 0.72rem; }}
    .data-table th:first-child, .data-table td:first-child {{
        max-width: 110px;
        white-space: normal;
    }}
}}
"""


def render_kpi_card(kpi: dict) -> str:
    delta_class = "neutral" if kpi.get("trend") == "neutral" else ""
    return (
        f'<div class="kpi">'
        f'<div class="kpi-label">{escape(kpi["label"])}</div>'
        f'<div class="kpi-value">{escape(kpi["value"])}</div>'
        f'<div class="kpi-delta {delta_class}">{escape(kpi["delta"])}</div>'
        f'</div>'
    )


def render_table(periods: list, rows: list[tuple[str, list, bool]]) -> str:
    head = "".join(f"<th>{escape(p)}</th>" for p in periods)
    body_rows = []
    for label, vals, is_total in rows:
        cls = ' class="total-row"' if is_total else ""
        cells = "".join(
            f"<td>{f'{v:,.0f}' if isinstance(v,(int,float)) else escape(str(v))}</td>"
            for v in vals
        )
        body_rows.append(f'<tr{cls}><td>{escape(label)}</td>{cells}</tr>')
    return (
        '<div class="table-scroll"><table class="data-table">'
        f'<thead><tr><th>Line Item</th>{head}</tr></thead>'
        f'<tbody>{"".join(body_rows)}</tbody>'
        '</table></div>'
    )


def build_html(data: dict, charts: dict) -> str:
    meta = data["meta"]
    periods = data["periods"]
    kpis = compute_kpis(data)

    pl_rows = [
        ("Total Revenue",   data["income"]["revenue"],     True),
        ("Gross Profit",    data["income"]["gross_profit"], False),
        ("EBITDA",          data["income"]["ebitda"],      False),
        ("EBIT",            data["income"]["ebit"],        False),
        ("Net Income",      data["income"]["net_income"],  True),
    ]
    cf_rows = [
        ("Cash from Operations", data["cashflow"]["cfo"],     False),
        ("Capital Expenditure",  data["cashflow"]["capex"],   False),
        ("Free Cash Flow",       data["cashflow"]["fcf"],     True),
        ("Ending Cash Balance",  data["cashflow"]["cash_bal"], False),
    ]

    title = escape(meta.get("title") or "Thetan — Financial Model")

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="color-scheme" content="dark">
<title>{title} · Dashboard</title>
<style>{CSS}</style>
</head>
<body>
<div class="container">

  <header class="header">
    <div class="brand">
      <div class="brand-mark">TH</div>
      <div class="brand-text">
        <h1>{title}</h1>
        <p>Integrated financial dashboard · FY2022A – FY2029E</p>
      </div>
    </div>
    <div class="meta-pills">
      <span class="pill live">{escape(str(meta.get('status') or 'Draft'))}</span>
      <span class="pill">{escape(str(meta.get('version') or 'v1.0'))}</span>
      <span class="pill">{escape(str(meta.get('currency') or 'USD'))}</span>
      <span class="pill">As of {escape(str(meta.get('date') or ''))}</span>
    </div>
  </header>

  <div class="kpi-grid">
    {''.join(render_kpi_card(k) for k in kpis)}
  </div>

  <div class="section-title">Revenue Performance</div>
  <div class="card-grid">
    <div class="card">
      <div class="card-head">
        <h3 class="card-title">Revenue Trajectory</h3>
        <span class="card-sub">Actuals + 5-yr Estimates</span>
      </div>
      <div class="chart-wrap">{charts['revenue']}</div>
    </div>
    <div class="card">
      <div class="card-head">
        <h3 class="card-title">Revenue by Segment</h3>
        <span class="card-sub">SaaS · Services · Marketplace</span>
      </div>
      <div class="chart-wrap">{charts['segments']}</div>
    </div>
  </div>

  <div class="section-title">Profitability</div>
  <div class="card-grid">
    <div class="card full">
      <div class="card-head">
        <h3 class="card-title">Margin Trend</h3>
        <span class="card-sub">Gross · EBITDA · Net Margin</span>
      </div>
      <div class="chart-wrap">{charts['margins']}</div>
    </div>
    <div class="card full">
      <div class="card-head">
        <h3 class="card-title">P&amp;L Summary</h3>
        <span class="card-sub">$ in thousands</span>
      </div>
      {render_table(periods, pl_rows)}
    </div>
  </div>

  <div class="section-title">Cash Flow</div>
  <div class="card-grid">
    <div class="card full">
      <div class="card-head">
        <h3 class="card-title">Cash Flow Bridge</h3>
        <span class="card-sub">CFO · CapEx · FCF</span>
      </div>
      <div class="chart-wrap">{charts['fcf']}</div>
    </div>
    <div class="card full">
      <div class="card-head">
        <h3 class="card-title">Cash Flow Statement</h3>
        <span class="card-sub">$ in thousands</span>
      </div>
      {render_table(periods, cf_rows)}
    </div>
  </div>

  <div class="section-title">Valuation</div>
  <div class="card-grid">
    <div class="card full">
      <div class="card-head">
        <h3 class="card-title">Enterprise Value Summary</h3>
        <span class="card-sub">DCF · Revenue Multiple · Range</span>
      </div>
      <div class="chart-wrap">{charts['valuation']}</div>
    </div>
  </div>

  <footer class="footer">
    Generated locally from <code>Financial_Model.xlsx</code> · No internet required ·
    {escape(str(meta.get('author') or 'Finance Team'))}
  </footer>
</div>
</body>
</html>
"""
    return html


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    root = Path(__file__).parent
    xlsx = root / "Financial_Model.xlsx"
    out  = root / "dashboard.html"

    if not xlsx.exists():
        raise SystemExit(f"Excel not found: {xlsx}")

    print(f"Reading {xlsx.name} …")
    data = read_model(xlsx)

    print("Rendering charts (SVG, embedded) …")
    charts = {
        "revenue":   chart_revenue_trajectory(data),
        "segments":  chart_segment_stacked(data),
        "margins":   chart_margin_trend(data),
        "fcf":       chart_fcf_bridge(data),
        "valuation": chart_valuation(data),
    }

    print("Building dashboard.html …")
    html = build_html(data, charts)
    out.write_text(html, encoding="utf-8")
    size_kb = out.stat().st_size / 1024
    print(f"  Saved → {out}  ({size_kb:.1f} KB)")
    print("Open the file directly in a browser — no internet required.")


if __name__ == "__main__":
    main()
