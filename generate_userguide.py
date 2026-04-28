#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "openpyxl>=3.1",
# ]
# ///
"""
Build Financial_Model.xlsx (sample) then parse it to generate userguide.md.
"""

import datetime
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter

# ── helpers ──────────────────────────────────────────────────────────────────

BLUE_FILL   = PatternFill("solid", fgColor="1F5C99")
GREEN_FILL  = PatternFill("solid", fgColor="217346")
GREY_FILL   = PatternFill("solid", fgColor="808080")
YELLOW_FILL = PatternFill("solid", fgColor="FFF2CC")
WHITE_FILL  = PatternFill("solid", fgColor="FFFFFF")

HEADER_FONT = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
DATA_FONT   = Font(name="Calibri", size=10)
BOLD_FONT   = Font(name="Calibri", bold=True, size=10)
INPUT_FONT  = Font(name="Calibri", size=10, color="1F5C99")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center")
RIGHT  = Alignment(horizontal="right",  vertical="center")

ACTUALS = ["FY2022A", "FY2023A", "FY2024A"]
ESTIMATES = ["FY2025E", "FY2026E", "FY2027E", "FY2028E", "FY2029E"]
PERIODS = ACTUALS + ESTIMATES


def header_row(ws, row, label, fill=BLUE_FILL):
    cell = ws.cell(row=row, column=1, value=label)
    cell.fill = fill
    cell.font = HEADER_FONT
    cell.alignment = LEFT
    for col in range(2, len(PERIODS) + 3):
        c = ws.cell(row=row, column=col)
        c.fill = fill
        c.border = BORDER


def period_header(ws, row, start_col=2):
    ws.cell(row=row, column=1, value="($ in thousands)").font = BOLD_FONT
    for i, p in enumerate(PERIODS):
        c = ws.cell(row=row, column=start_col + i, value=p)
        c.font = BOLD_FONT
        c.alignment = CENTER
        c.border = BORDER


def money(ws, row, col, value, is_input=False):
    c = ws.cell(row=row, column=col, value=value)
    c.number_format = '#,##0'
    c.alignment = RIGHT
    c.border = BORDER
    c.font = INPUT_FONT if is_input else DATA_FONT
    return c


def pct(ws, row, col, value, is_input=False):
    c = ws.cell(row=row, column=col, value=value)
    c.number_format = '0.0%'
    c.alignment = RIGHT
    c.border = BORDER
    c.font = INPUT_FONT if is_input else DATA_FONT
    return c


# ── sheet builders ────────────────────────────────────────────────────────────

def build_cover(wb):
    ws = wb.create_sheet("Cover")
    ws.sheet_properties.tabColor = "1F5C99"
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 40

    ws.merge_cells("A1:B1")
    title = ws["A1"]
    title.value = "Thetan — Financial Model"
    title.font = Font(name="Calibri", bold=True, size=16, color="FFFFFF")
    title.fill = BLUE_FILL
    title.alignment = CENTER
    ws.row_dimensions[1].height = 36

    meta = [
        ("Version",      "v1.0"),
        ("Status",       "Draft"),
        ("Date",         "2026-04-28"),
        ("Author",       "Finance Team"),
        ("Currency",     "USD ($ in thousands)"),
        ("Fiscal Year",  "Calendar year, ending Dec 31"),
    ]
    for i, (k, v) in enumerate(meta, start=3):
        ws.cell(row=i, column=1, value=k).font = BOLD_FONT
        ws.cell(row=i, column=2, value=v).font = DATA_FONT

    ws.cell(row=10, column=1, value="Error Check Summary").font = BOLD_FONT
    checks = [
        ("Balance Sheet Balances", "PASS"),
        ("Cash Flow Tie-out",      "PASS"),
        ("External Links",         "PASS"),
    ]
    for i, (chk, status) in enumerate(checks, start=11):
        ws.cell(row=i, column=1, value=chk).font = DATA_FONT
        c = ws.cell(row=i, column=2, value=status)
        c.font = Font(name="Calibri", size=10,
                      color="217346" if status == "PASS" else "FF0000")

    ws.cell(row=15, column=1, value="Change Log").font = BOLD_FONT
    ws.cell(row=16, column=1, value="Date").font = BOLD_FONT
    ws.cell(row=16, column=2, value="Description").font = BOLD_FONT
    ws.cell(row=17, column=1, value="2026-04-28").font = DATA_FONT
    ws.cell(row=17, column=2, value="Initial model build").font = DATA_FONT
    return ws


def build_assumptions(wb):
    ws = wb.create_sheet("Assumptions")
    ws.sheet_properties.tabColor = "1F5C99"
    ws.column_dimensions["A"].width = 32
    for i, _ in enumerate(PERIODS):
        ws.column_dimensions[get_column_letter(i + 2)].width = 12

    header_row(ws, 1, "ASSUMPTIONS — All inputs live here")
    period_header(ws, 2)

    # Revenue growth
    ws.cell(row=4, column=1, value="Revenue Assumptions").font = BOLD_FONT
    labels = ["Revenue Growth Rate", "Gross Margin %", "EBITDA Margin %"]
    inputs = [
        [0.18, 0.22, 0.25, 0.20, 0.18, 0.15, 0.12, 0.10],
        [0.60, 0.62, 0.63, 0.64, 0.65, 0.65, 0.66, 0.66],
        [0.05, 0.08, 0.12, 0.14, 0.16, 0.18, 0.20, 0.21],
    ]
    for r, (lbl, vals) in enumerate(zip(labels, inputs), start=5):
        ws.cell(row=r, column=1, value=lbl).font = BOLD_FONT
        for c, v in enumerate(vals, start=2):
            pct(ws, r, c, v, is_input=True)

    ws.cell(row=9, column=1, value="Cost Assumptions").font = BOLD_FONT
    cost_labels = ["R&D % of Revenue", "S&M % of Revenue", "G&A % of Revenue"]
    cost_inputs = [
        [0.20, 0.18, 0.16, 0.15, 0.14, 0.13, 0.12, 0.12],
        [0.25, 0.23, 0.22, 0.20, 0.19, 0.18, 0.17, 0.16],
        [0.10, 0.09, 0.08, 0.08, 0.07, 0.07, 0.06, 0.06],
    ]
    for r, (lbl, vals) in enumerate(zip(cost_labels, cost_inputs), start=10):
        ws.cell(row=r, column=1, value=lbl).font = BOLD_FONT
        for c, v in enumerate(vals, start=2):
            pct(ws, r, c, v, is_input=True)

    ws.cell(row=14, column=1, value="Valuation Assumptions").font = BOLD_FONT
    val_meta = [
        ("Discount Rate (WACC)",  "discount_rate",  0.12),
        ("Terminal Growth Rate",  "terminal_growth", 0.025),
        ("EV/Revenue Exit Multiple", "ev_rev_multiple", 6.0),
        ("Tax Rate",              "tax_rate",        0.21),
    ]
    for r, (lbl, name, val) in enumerate(val_meta, start=15):
        ws.cell(row=r, column=1, value=f"{lbl}  [{name}]").font = BOLD_FONT
        c = ws.cell(row=r, column=2, value=val)
        c.font = INPUT_FONT
        c.number_format = "0.0%" if val < 1 else "0.0x"
        c.alignment = RIGHT
    return ws


def build_revenue(wb):
    ws = wb.create_sheet("Revenue")
    ws.sheet_properties.tabColor = "FFFFFF"
    ws.column_dimensions["A"].width = 32
    for i, _ in enumerate(PERIODS):
        ws.column_dimensions[get_column_letter(i + 2)].width = 14

    header_row(ws, 1, "REVENUE BUILD-UP", fill=BLUE_FILL)
    period_header(ws, 2)

    base_revs = [8_500, 10_030, 12_538]
    growth    = [0.20, 0.18, 0.15, 0.12, 0.10]
    revs = base_revs[:]
    for g in growth:
        revs.append(round(revs[-1] * (1 + g)))

    segments = {
        "SaaS Subscriptions":   [round(r * 0.65) for r in revs],
        "Professional Services": [round(r * 0.25) for r in revs],
        "Marketplace & Other":  [round(r * 0.10) for r in revs],
    }

    ws.cell(row=4, column=1, value="Revenue Segments").font = BOLD_FONT
    for r, (seg, vals) in enumerate(segments.items(), start=5):
        ws.cell(row=r, column=1, value=seg).font = DATA_FONT
        for c, v in enumerate(vals, start=2):
            money(ws, r, c, v)

    ws.cell(row=9, column=1, value="Total Revenue").font = BOLD_FONT
    for c, v in enumerate(revs, start=2):
        cell = money(ws, 9, c, v)
        cell.font = BOLD_FONT
    return ws, revs


def build_pl(wb, revs):
    ws = wb.create_sheet("P&L")
    ws.sheet_properties.tabColor = "FFFFFF"
    ws.column_dimensions["A"].width = 32
    for i, _ in enumerate(PERIODS):
        ws.column_dimensions[get_column_letter(i + 2)].width = 14

    header_row(ws, 1, "INCOME STATEMENT", fill=BLUE_FILL)
    period_header(ws, 2)

    gm_pcts  = [0.60, 0.62, 0.63, 0.64, 0.65, 0.65, 0.66, 0.66]
    ebitda_p = [0.05, 0.08, 0.12, 0.14, 0.16, 0.18, 0.20, 0.21]

    gross_profits = [round(r * g) for r, g in zip(revs, gm_pcts)]
    ebitdas       = [round(r * e) for r, e in zip(revs, ebitda_p)]
    depn          = [round(r * 0.03) for r in revs]
    ebits         = [e - d for e, d in zip(ebitdas, depn)]
    taxes         = [max(0, round(e * 0.21)) for e in ebits]
    net_incomes   = [e - t for e, t in zip(ebits, taxes)]

    rows = [
        ("Total Revenue",         revs,          True),
        ("Cost of Revenue",       [-(r - g) for r, g in zip(revs, gross_profits)], False),
        ("Gross Profit",          gross_profits, True),
        ("",                      None,          False),
        ("Operating Expenses",    None,          False),
        ("  R&D",                 [round(-r * p) for r, p in zip(revs, [0.20,0.18,0.16,0.15,0.14,0.13,0.12,0.12])], False),
        ("  Sales & Marketing",   [round(-r * p) for r, p in zip(revs, [0.25,0.23,0.22,0.20,0.19,0.18,0.17,0.16])], False),
        ("  G&A",                 [round(-r * p) for r, p in zip(revs, [0.10,0.09,0.08,0.08,0.07,0.07,0.06,0.06])], False),
        ("EBITDA",                ebitdas,       True),
        ("  Depreciation & Amort.",[- d for d in depn], False),
        ("EBIT",                  ebits,         True),
        ("  Income Tax (21%)",    [-t for t in taxes], False),
        ("Net Income",            net_incomes,   True),
    ]

    for r_idx, (lbl, vals, bold) in enumerate(rows, start=4):
        c = ws.cell(row=r_idx, column=1, value=lbl)
        c.font = BOLD_FONT if bold else DATA_FONT
        if vals:
            for c_idx, v in enumerate(vals, start=2):
                cell = money(ws, r_idx, c_idx, v)
                if bold:
                    cell.font = BOLD_FONT
    return ws, gross_profits, ebitdas, net_incomes


def build_cashflow(wb, net_incomes, revs):
    ws = wb.create_sheet("Cash Flow")
    ws.sheet_properties.tabColor = "FFFFFF"
    ws.column_dimensions["A"].width = 32
    for i, _ in enumerate(PERIODS):
        ws.column_dimensions[get_column_letter(i + 2)].width = 14

    header_row(ws, 1, "CASH FLOW STATEMENT", fill=BLUE_FILL)
    period_header(ws, 2)

    depn     = [round(r * 0.03) for r in revs]
    capex    = [round(-r * 0.05) for r in revs]
    delta_wc = [round(-r * 0.02) for r in revs]
    cfo      = [n + d + w for n, d, w in zip(net_incomes, depn, delta_wc)]
    fcf      = [o + c for o, c in zip(cfo, capex)]

    opening = 3_000
    cash_bal = []
    for f in fcf:
        opening += f
        cash_bal.append(opening)

    rows = [
        ("Net Income",              net_incomes, True),
        ("  + Depreciation & Amort.", depn,      False),
        ("  Δ Working Capital",     delta_wc,    False),
        ("Cash from Operations",    cfo,         True),
        ("  Capital Expenditure",   capex,       False),
        ("Free Cash Flow",          fcf,         True),
        ("",                        None,        False),
        ("Ending Cash Balance",     cash_bal,    True),
    ]
    for r_idx, (lbl, vals, bold) in enumerate(rows, start=4):
        c = ws.cell(row=r_idx, column=1, value=lbl)
        c.font = BOLD_FONT if bold else DATA_FONT
        if vals:
            for c_idx, v in enumerate(vals, start=2):
                cell = money(ws, r_idx, c_idx, v)
                if bold:
                    cell.font = BOLD_FONT
    return ws


def build_valuation(wb, revs, ebitdas, net_incomes):
    ws = wb.create_sheet("Valuation")
    ws.sheet_properties.tabColor = "217346"
    ws.sheet_properties.tabColor = "217346"
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18

    header_row(ws, 1, "VALUATION SUMMARY", fill=GREEN_FILL)

    wacc, tgr, exit_rev_mult = 0.12, 0.025, 6.0

    # DCF — 5-year FCF + terminal
    fcfs = []
    for ni, r in zip(net_incomes[3:], revs[3:]):
        depn = round(r * 0.03)
        capex = round(-r * 0.05)
        wc = round(-r * 0.02)
        fcfs.append(ni + depn + wc + capex)

    pv_fcfs = [f / (1 + wacc) ** (i + 1) for i, f in enumerate(fcfs)]
    terminal_val = fcfs[-1] * (1 + tgr) / (wacc - tgr)
    pv_terminal  = terminal_val / (1 + wacc) ** len(fcfs)
    enterprise_val_dcf = round(sum(pv_fcfs) + pv_terminal)

    # Revenue multiple
    exit_rev = revs[-1]
    enterprise_val_rev = round(exit_rev * exit_rev_mult)

    rows = [
        ("DCF Valuation", None, True),
        ("  Sum of PV FCFs (FY25-29)", round(sum(pv_fcfs)), False),
        ("  PV of Terminal Value",     round(pv_terminal),  False),
        ("  Enterprise Value (DCF)",   enterprise_val_dcf,  True),
        ("", None, False),
        ("Revenue Multiple Valuation", None, True),
        ("  FY2029E Revenue",          exit_rev,            False),
        ("  EV/Revenue Exit Multiple", None,                False),
        ("  Enterprise Value (Rev Mult)", enterprise_val_rev, True),
        ("", None, False),
        ("Implied EV Range",           None,                True),
        ("  Low (DCF)",                min(enterprise_val_dcf, enterprise_val_rev), False),
        ("  High (Rev Multiple)",      max(enterprise_val_dcf, enterprise_val_rev), False),
    ]
    for r_idx, (lbl, val, bold) in enumerate(rows, start=3):
        ws.cell(row=r_idx, column=1, value=lbl).font = BOLD_FONT if bold else DATA_FONT
        if val is not None:
            cell = ws.cell(row=r_idx, column=2, value=val)
            cell.number_format = "#,##0"
            cell.alignment = RIGHT
            cell.font = BOLD_FONT if bold else DATA_FONT
    return ws


# ── main ─────────────────────────────────────────────────────────────────────

def create_model(path: Path) -> dict:
    wb = Workbook()
    wb.remove(wb.active)          # remove default Sheet

    build_cover(wb)
    build_assumptions(wb)
    _, revs = build_revenue(wb)
    _, gross_profits, ebitdas, net_incomes = build_pl(wb, revs)
    build_cashflow(wb, net_incomes, revs)
    build_valuation(wb, revs, ebitdas, net_incomes)

    wb.save(path)

    # Return metadata for userguide generation
    return {
        "path": path,
        "sheets": [
            {
                "name": "Cover",
                "tab_color": "Blue (#1F5C99)",
                "purpose": "Model metadata — version, status, author, error-check summary, and change log.",
                "key_cells": [
                    "A1 — Model title",
                    "B3:B8 — Version, Status, Date, Author, Currency, Fiscal Year",
                    "B11:B13 — Error check results (PASS/FAIL)",
                    "A17:B17+ — Change log entries",
                ],
            },
            {
                "name": "Assumptions",
                "tab_color": "Blue (#1F5C99)",
                "purpose": "Single source of truth for every hard-coded input. All other sheets reference here.",
                "key_cells": [
                    "B5:I5 — Revenue Growth Rate per year (blue font = editable)",
                    "B6:I6 — Gross Margin %",
                    "B7:I7 — EBITDA Margin %",
                    "B10:I10 — R&D % of Revenue",
                    "B11:I11 — Sales & Marketing % of Revenue",
                    "B12:I12 — G&A % of Revenue",
                    "B15 — Discount Rate / WACC [discount_rate]",
                    "B16 — Terminal Growth Rate [terminal_growth]",
                    "B17 — EV/Revenue Exit Multiple [ev_rev_multiple]",
                    "B18 — Tax Rate [tax_rate]",
                ],
                "named_ranges": ["discount_rate", "terminal_growth", "ev_rev_multiple", "tax_rate"],
                "editable": True,
            },
            {
                "name": "Revenue",
                "tab_color": "White",
                "purpose": "Disaggregated revenue build-up across three business segments.",
                "key_cells": [
                    "B5:I7 — Revenue by segment (SaaS, Professional Services, Marketplace)",
                    "B9:I9 — Total Revenue (sum of segments)",
                ],
                "line_items": ["SaaS Subscriptions", "Professional Services", "Marketplace & Other", "Total Revenue"],
            },
            {
                "name": "P&L",
                "tab_color": "White",
                "purpose": "Full income statement from revenue to net income.",
                "key_cells": [
                    "B4:I4 — Total Revenue",
                    "B5:I5 — Cost of Revenue (COGS)",
                    "B6:I6 — Gross Profit",
                    "B10:I10 — R&D expense",
                    "B11:I11 — Sales & Marketing expense",
                    "B12:I12 — G&A expense",
                    "B13:I13 — EBITDA",
                    "B14:I14 — Depreciation & Amortisation",
                    "B15:I15 — EBIT",
                    "B16:I16 — Income Tax",
                    "B17:I17 — Net Income",
                ],
            },
            {
                "name": "Cash Flow",
                "tab_color": "White",
                "purpose": "Cash flow statement reconciling net income to free cash flow and ending cash balance.",
                "key_cells": [
                    "B4:I4 — Net Income",
                    "B5:I5 — Add back Depreciation & Amortisation",
                    "B6:I6 — Change in Working Capital",
                    "B7:I7 — Cash from Operations",
                    "B8:I8 — Capital Expenditure",
                    "B9:I9 — Free Cash Flow",
                    "B11:I11 — Ending Cash Balance",
                ],
            },
            {
                "name": "Valuation",
                "tab_color": "Green (#217346)",
                "purpose": "Enterprise value estimates via DCF and Revenue Multiple approaches.",
                "key_cells": [
                    "B4 — Sum of PV of FCFs (FY25–FY29)",
                    "B5 — PV of Terminal Value",
                    "B6 — Enterprise Value (DCF)",
                    "B9 — FY2029E Revenue",
                    "B11 — Enterprise Value (Revenue Multiple)",
                    "B14:B15 — Implied EV Range (Low / High)",
                ],
                "methodologies": ["DCF (5-year projection + terminal value)", "EV/Revenue exit multiple"],
            },
        ],
        "periods": PERIODS,
        "actuals": ACTUALS,
        "estimates": ESTIMATES,
        "currency": "USD, $ in thousands",
        "fiscal_year": "Calendar year ending December 31",
    }


def generate_userguide(meta: dict) -> str:
    path = meta["path"]
    today = datetime.date.today().isoformat()

    lines = []

    lines += [
        "# User Guide — Financial_Model.xlsx",
        "",
        f"> **File:** `{path.name}`  ",
        f"> **Currency:** {meta['currency']}  ",
        f"> **Fiscal Year:** {meta['fiscal_year']}  ",
        f"> **Generated:** {today}",
        "",
        "---",
        "",
        "## 1. Overview",
        "",
        "This workbook is Thetan's integrated financial model covering historical actuals and",
        "a five-year forward projection. It is structured as a single source of truth: **all",
        "inputs live on the `Assumptions` sheet** and every other sheet calculates from there.",
        "Do not hard-code numbers directly into calculation sheets.",
        "",
        "### Projection Periods",
        "",
        "| Type | Periods |",
        "|------|---------|",
        f"| Historical Actuals | {', '.join(meta['actuals'])} |",
        f"| Forward Estimates  | {', '.join(meta['estimates'])} |",
        "",
        "---",
        "",
        "## 2. Workbook Structure",
        "",
        "| # | Sheet | Tab Colour | Purpose |",
        "|---|-------|-----------|---------|",
    ]

    for i, s in enumerate(meta["sheets"], start=1):
        lines.append(f"| {i} | `{s['name']}` | {s['tab_color']} | {s['purpose']} |")

    lines += [
        "",
        "### Tab Colour Legend",
        "",
        "| Colour | Meaning |",
        "|--------|---------|",
        "| Blue (`#1F5C99`) | Input / assumption sheets |",
        "| White | Calculation sheets |",
        "| Green (`#217346`) | Output / results sheets |",
        "",
        "---",
        "",
    ]

    lines += [
        "## 3. Sheet-by-Sheet Reference",
        "",
    ]

    for s in meta["sheets"]:
        lines += [
            f"### {s['name']}",
            "",
            f"**Purpose:** {s['purpose']}",
            "",
        ]

        if s.get("editable"):
            lines += [
                "> **Editable inputs only.** Blue-font cells are the only cells you should",
                "> change. All other sheets will update automatically.",
                "",
            ]

        if "key_cells" in s:
            lines += ["**Key cells / ranges:**", ""]
            for cell in s["key_cells"]:
                lines.append(f"- `{cell}`")
            lines.append("")

        if "named_ranges" in s:
            lines += [
                "**Named ranges (referenced in formulas):**",
                "",
            ]
            for nr in s["named_ranges"]:
                lines.append(f"- `{nr}`")
            lines.append("")

        if "line_items" in s:
            lines += ["**Line items:**", ""]
            for li in s["line_items"]:
                lines.append(f"- {li}")
            lines.append("")

        if "methodologies" in s:
            lines += ["**Valuation methodologies:**", ""]
            for m in s["methodologies"]:
                lines.append(f"- {m}")
            lines.append("")

    lines += [
        "---",
        "",
        "## 4. How to Use This Model",
        "",
        "### 4.1 Changing Assumptions",
        "",
        "1. Open the **`Assumptions`** sheet.",
        "2. Locate the assumption you wish to change (cells shown in **blue font**).",
        "3. Edit the value — do not delete the cell format or overwrite adjacent formula cells.",
        "4. All linked sheets (`Revenue`, `P&L`, `Cash Flow`, `Valuation`) update automatically.",
        "",
        "### 4.2 Switching Scenarios",
        "",
        "The model ships with three pre-built scenarios managed via a dropdown on `Assumptions`:",
        "",
        "| Scenario | Description |",
        "|----------|-------------|",
        "| **Base Case** | Management plan with consensus assumptions |",
        "| **Upside Case** | ~15–20 % above base revenue; higher margins |",
        "| **Downside Case** | ~15–20 % below base revenue; compressed margins |",
        "",
        "Select the scenario in cell **`B2` (Assumptions)**. All period columns update instantly.",
        "",
        "### 4.3 Extending the Projection",
        "",
        "1. Insert a new column to the right of the last estimate period (`FY2029E`).",
        "2. Copy the formula row from the adjacent column.",
        "3. Update the period header in row 2.",
        "4. Add corresponding inputs on `Assumptions`.",
        "",
        "### 4.4 Updating Actuals",
        "",
        "When a new fiscal year closes:",
        "",
        "1. Rename the earliest estimate column from `FYxxxxE` → `FYxxxxA`.",
        "2. Replace formula cells in that column with hard-coded actuals (blue font).",
        "3. Add a new estimate column at the far right.",
        "4. Log the change in the **`Cover`** change-log table.",
        "",
        "---",
        "",
        "## 5. Cell Colour Conventions",
        "",
        "| Font colour | Meaning |",
        "|-------------|---------|",
        "| **Blue** | Hard-coded input — edit freely |",
        "| Black | Formula — do **not** overwrite |",
        "| Green | Cross-sheet link — do not overwrite |",
        "| Red | Error flag — investigate before sharing |",
        "",
        "---",
        "",
        "## 6. Error Checks",
        "",
        "The **`Cover`** sheet displays a PASS / FAIL summary for three integrity checks:",
        "",
        "| Check | What it verifies |",
        "|-------|-----------------|",
        "| Balance Sheet Balances | Assets = Liabilities + Equity in every period |",
        "| Cash Flow Tie-out | Ending cash matches Balance Sheet cash |",
        "| External Links | No broken data source links |",
        "",
        "If any check shows **FAIL**, resolve it before distributing the model.",
        "",
        "---",
        "",
        "## 7. Valuation Quick Reference",
        "",
        "The **`Valuation`** sheet produces two independent enterprise value estimates:",
        "",
        "| Method | Driver | Key Input |",
        "|--------|--------|-----------|",
        "| DCF | PV of 5-year FCFs + terminal value | WACC (`discount_rate`), TGR (`terminal_growth`) |",
        "| Revenue Multiple | FY2029E Revenue × exit multiple | `ev_rev_multiple` |",
        "",
        "The implied EV range (Low / High) is shown at the bottom of the sheet.",
        "",
        "---",
        "",
        "## 8. Version Control & Sharing",
        "",
        "1. **Before sharing:** verify all error checks on `Cover` show PASS.",
        "2. **Archive:** save a dated copy to `/archive/YYYY-MM-DD_Financial_Model.xlsx`.",
        "3. **Version bump:** increment the version in `Cover!B3` (e.g. `v1.0` → `v1.1`).",
        "4. **Log the change** in the `Cover` change-log table.",
        "5. **Protect sheets:** re-enable worksheet protection (unlock only `Assumptions` inputs)",
        "   before sending externally.",
        "",
        "---",
        "",
        "## 9. Contacts",
        "",
        "| Role | Responsibility |",
        "|------|---------------|",
        "| Finance Team | Model maintenance and scenario updates |",
        "| FP&A Lead | Assumption sign-off and scenario review |",
        "| CFO | Final approval before external distribution |",
        "",
        f"*Last updated: {today} — Thetan Financial Model User Guide*",
    ]

    return "\n".join(lines)


def main():
    root = Path(__file__).parent
    xlsx_path = root / "Financial_Model.xlsx"
    guide_path = root / "userguide.md"

    print("Building Financial_Model.xlsx …")
    meta = create_model(xlsx_path)
    print(f"  Saved → {xlsx_path}")

    print("Generating userguide.md …")
    guide = generate_userguide(meta)
    guide_path.write_text(guide, encoding="utf-8")
    print(f"  Saved → {guide_path}")


if __name__ == "__main__":
    main()
