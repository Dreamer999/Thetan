# User Guide — Financial_Model.xlsx

> **File:** `Financial_Model.xlsx`  
> **Currency:** USD, $ in thousands  
> **Fiscal Year:** Calendar year ending December 31  
> **Generated:** 2026-04-28

---

## 1. Overview

This workbook is Thetan's integrated financial model covering historical actuals and
a five-year forward projection. It is structured as a single source of truth: **all
inputs live on the `Assumptions` sheet** and every other sheet calculates from there.
Do not hard-code numbers directly into calculation sheets.

### Projection Periods

| Type | Periods |
|------|---------|
| Historical Actuals | FY2022A, FY2023A, FY2024A |
| Forward Estimates  | FY2025E, FY2026E, FY2027E, FY2028E, FY2029E |

---

## 2. Workbook Structure

| # | Sheet | Tab Colour | Purpose |
|---|-------|-----------|---------|
| 1 | `Cover` | Blue (#1F5C99) | Model metadata — version, status, author, error-check summary, and change log. |
| 2 | `Assumptions` | Blue (#1F5C99) | Single source of truth for every hard-coded input. All other sheets reference here. |
| 3 | `Revenue` | White | Disaggregated revenue build-up across three business segments. |
| 4 | `P&L` | White | Full income statement from revenue to net income. |
| 5 | `Cash Flow` | White | Cash flow statement reconciling net income to free cash flow and ending cash balance. |
| 6 | `Valuation` | Green (#217346) | Enterprise value estimates via DCF and Revenue Multiple approaches. |

### Tab Colour Legend

| Colour | Meaning |
|--------|---------|
| Blue (`#1F5C99`) | Input / assumption sheets |
| White | Calculation sheets |
| Green (`#217346`) | Output / results sheets |

---

## 3. Sheet-by-Sheet Reference

### Cover

**Purpose:** Model metadata — version, status, author, error-check summary, and change log.

**Key cells / ranges:**

- `A1 — Model title`
- `B3:B8 — Version, Status, Date, Author, Currency, Fiscal Year`
- `B11:B13 — Error check results (PASS/FAIL)`
- `A17:B17+ — Change log entries`

### Assumptions

**Purpose:** Single source of truth for every hard-coded input. All other sheets reference here.

> **Editable inputs only.** Blue-font cells are the only cells you should
> change. All other sheets will update automatically.

**Key cells / ranges:**

- `B5:I5 — Revenue Growth Rate per year (blue font = editable)`
- `B6:I6 — Gross Margin %`
- `B7:I7 — EBITDA Margin %`
- `B10:I10 — R&D % of Revenue`
- `B11:I11 — Sales & Marketing % of Revenue`
- `B12:I12 — G&A % of Revenue`
- `B15 — Discount Rate / WACC [discount_rate]`
- `B16 — Terminal Growth Rate [terminal_growth]`
- `B17 — EV/Revenue Exit Multiple [ev_rev_multiple]`
- `B18 — Tax Rate [tax_rate]`

**Named ranges (referenced in formulas):**

- `discount_rate`
- `terminal_growth`
- `ev_rev_multiple`
- `tax_rate`

### Revenue

**Purpose:** Disaggregated revenue build-up across three business segments.

**Key cells / ranges:**

- `B5:I7 — Revenue by segment (SaaS, Professional Services, Marketplace)`
- `B9:I9 — Total Revenue (sum of segments)`

**Line items:**

- SaaS Subscriptions
- Professional Services
- Marketplace & Other
- Total Revenue

### P&L

**Purpose:** Full income statement from revenue to net income.

**Key cells / ranges:**

- `B4:I4 — Total Revenue`
- `B5:I5 — Cost of Revenue (COGS)`
- `B6:I6 — Gross Profit`
- `B10:I10 — R&D expense`
- `B11:I11 — Sales & Marketing expense`
- `B12:I12 — G&A expense`
- `B13:I13 — EBITDA`
- `B14:I14 — Depreciation & Amortisation`
- `B15:I15 — EBIT`
- `B16:I16 — Income Tax`
- `B17:I17 — Net Income`

### Cash Flow

**Purpose:** Cash flow statement reconciling net income to free cash flow and ending cash balance.

**Key cells / ranges:**

- `B4:I4 — Net Income`
- `B5:I5 — Add back Depreciation & Amortisation`
- `B6:I6 — Change in Working Capital`
- `B7:I7 — Cash from Operations`
- `B8:I8 — Capital Expenditure`
- `B9:I9 — Free Cash Flow`
- `B11:I11 — Ending Cash Balance`

### Valuation

**Purpose:** Enterprise value estimates via DCF and Revenue Multiple approaches.

**Key cells / ranges:**

- `B4 — Sum of PV of FCFs (FY25–FY29)`
- `B5 — PV of Terminal Value`
- `B6 — Enterprise Value (DCF)`
- `B9 — FY2029E Revenue`
- `B11 — Enterprise Value (Revenue Multiple)`
- `B14:B15 — Implied EV Range (Low / High)`

**Valuation methodologies:**

- DCF (5-year projection + terminal value)
- EV/Revenue exit multiple

---

## 4. How to Use This Model

### 4.1 Changing Assumptions

1. Open the **`Assumptions`** sheet.
2. Locate the assumption you wish to change (cells shown in **blue font**).
3. Edit the value — do not delete the cell format or overwrite adjacent formula cells.
4. All linked sheets (`Revenue`, `P&L`, `Cash Flow`, `Valuation`) update automatically.

### 4.2 Switching Scenarios

The model ships with three pre-built scenarios managed via a dropdown on `Assumptions`:

| Scenario | Description |
|----------|-------------|
| **Base Case** | Management plan with consensus assumptions |
| **Upside Case** | ~15–20 % above base revenue; higher margins |
| **Downside Case** | ~15–20 % below base revenue; compressed margins |

Select the scenario in cell **`B2` (Assumptions)**. All period columns update instantly.

### 4.3 Extending the Projection

1. Insert a new column to the right of the last estimate period (`FY2029E`).
2. Copy the formula row from the adjacent column.
3. Update the period header in row 2.
4. Add corresponding inputs on `Assumptions`.

### 4.4 Updating Actuals

When a new fiscal year closes:

1. Rename the earliest estimate column from `FYxxxxE` → `FYxxxxA`.
2. Replace formula cells in that column with hard-coded actuals (blue font).
3. Add a new estimate column at the far right.
4. Log the change in the **`Cover`** change-log table.

---

## 5. Cell Colour Conventions

| Font colour | Meaning |
|-------------|---------|
| **Blue** | Hard-coded input — edit freely |
| Black | Formula — do **not** overwrite |
| Green | Cross-sheet link — do not overwrite |
| Red | Error flag — investigate before sharing |

---

## 6. Error Checks

The **`Cover`** sheet displays a PASS / FAIL summary for three integrity checks:

| Check | What it verifies |
|-------|-----------------|
| Balance Sheet Balances | Assets = Liabilities + Equity in every period |
| Cash Flow Tie-out | Ending cash matches Balance Sheet cash |
| External Links | No broken data source links |

If any check shows **FAIL**, resolve it before distributing the model.

---

## 7. Valuation Quick Reference

The **`Valuation`** sheet produces two independent enterprise value estimates:

| Method | Driver | Key Input |
|--------|--------|-----------|
| DCF | PV of 5-year FCFs + terminal value | WACC (`discount_rate`), TGR (`terminal_growth`) |
| Revenue Multiple | FY2029E Revenue × exit multiple | `ev_rev_multiple` |

The implied EV range (Low / High) is shown at the bottom of the sheet.

---

## 8. Version Control & Sharing

1. **Before sharing:** verify all error checks on `Cover` show PASS.
2. **Archive:** save a dated copy to `/archive/YYYY-MM-DD_Financial_Model.xlsx`.
3. **Version bump:** increment the version in `Cover!B3` (e.g. `v1.0` → `v1.1`).
4. **Log the change** in the `Cover` change-log table.
5. **Protect sheets:** re-enable worksheet protection (unlock only `Assumptions` inputs)
   before sending externally.

---

## 9. Contacts

| Role | Responsibility |
|------|---------------|
| Finance Team | Model maintenance and scenario updates |
| FP&A Lead | Assumption sign-off and scenario review |
| CFO | Final approval before external distribution |

*Last updated: 2026-04-28 — Thetan Financial Model User Guide*