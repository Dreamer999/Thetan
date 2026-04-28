# Financial Model Design Guidelines

## 1. Purpose

These guidelines establish standards for building, formatting, and maintaining financial models (e.g., `Financial_Model.xlsx`) to ensure consistency, auditability, and ease of collaboration across teams.

---

## 2. File & Folder Structure

```
/financial-models
  Financial_Model.xlsx          # Primary working model
  /inputs                       # Raw data files (never edit directly in model)
  /outputs                      # Exported reports and charts
  /archive                      # Versioned snapshots (YYYY-MM-DD_Financial_Model.xlsx)
```

- Keep one canonical model file. Archive before major changes.
- Name archived files with ISO date prefix: `2026-04-28_Financial_Model.xlsx`.
- Never store credentials, API keys, or confidential PII inside model files.

---

## 3. Workbook Architecture

### 3.1 Sheet Order

| Position | Tab Name       | Purpose                                   |
|----------|---------------|-------------------------------------------|
| 1        | `Cover`        | Title, version, date, author, status      |
| 2        | `TOC`          | Table of contents with hyperlinks         |
| 3        | `Assumptions`  | All hard-coded inputs in one place        |
| 4        | `Revenue`      | Revenue build-up                          |
| 5        | `COGS`         | Cost of goods sold                        |
| 6        | `OpEx`         | Operating expenses                        |
| 7        | `P&L`          | Income statement                          |
| 8        | `Balance Sheet`| Balance sheet                             |
| 9        | `Cash Flow`    | Cash flow statement                       |
| 10       | `Debt`         | Debt schedule                             |
| 11       | `Valuation`    | DCF / multiples / sensitivity tables      |
| 12       | `Charts`       | Visualisations only (no data here)        |
| 13+      | `Raw_*`        | Source data imports — never formula-edited|

### 3.2 Sheet Colour Coding

| Colour     | Hex       | Meaning                         |
|------------|-----------|---------------------------------|
| Blue       | `#1F5C99` | Input / assumption sheets       |
| White      | `#FFFFFF` | Calculation sheets              |
| Green      | `#217346` | Output / results sheets         |
| Grey       | `#808080` | Supporting / archived sheets    |

---

## 4. Cell Conventions

### 4.1 Colour Coding

| Font Colour | Meaning                                   |
|-------------|-------------------------------------------|
| Blue        | Hard-coded input (user-editable)          |
| Black       | Formula (never overwrite manually)        |
| Green       | Link from another sheet                   |
| Red         | Error check or flag                       |

### 4.2 Input vs. Formula Cells

- **All hard-coded values** must live on the `Assumptions` sheet only. Calculations reference them; they never embed literal numbers.
- **Formula cells** must be consistent across a row: the same formula copied across all period columns, no cell-specific overrides.
- Never mix inputs and formulas in the same cell.

### 4.3 Named Ranges

- Use named ranges for frequently referenced assumptions (e.g., `tax_rate`, `discount_rate`, `growth_rate_y1`).
- Names must be `snake_case`, descriptive, and scoped to the workbook.
- Document all named ranges on the `Assumptions` sheet.

---

## 5. Formatting Standards

### 5.1 Number Formats

| Data Type         | Format Code                  | Example          |
|-------------------|------------------------------|------------------|
| Currency (whole)  | `$#,##0`                     | $1,250,000       |
| Currency (decimal)| `$#,##0.00`                  | $1,250,000.00    |
| Percentage        | `0.0%`                       | 12.5%            |
| Multiples         | `0.0x`                       | 3.5x             |
| Basis points      | `0 bps`                      | 150 bps          |
| Whole number      | `#,##0`                      | 1,250,000        |
| Date              | `MMM-YYYY`                   | Apr-2026         |
| Year              | `0`                          | 2026             |

### 5.2 Units

- State the unit prominently in the section header (e.g., `$ in thousands`).
- Never switch units within a section. Choose thousands or millions and apply consistently.

### 5.3 Typography

- Font: Calibri 10pt for data cells; Calibri 11pt bold for headers.
- Row height: 15pt standard; 18pt for section headers.
- Freeze the first column and header row on every calculation sheet.

### 5.4 Negative Numbers

- Display negatives as `(1,250)` in P&L and cash flow (accounting convention).
- Use negative signs only in mathematical calculation sheets.

---

## 6. Projection Periods

- Label periods clearly: `FY2024A`, `FY2025E`, `FY2026E` (A = Actual, E = Estimate).
- The transition year between actuals and estimates must be clearly marked.
- Minimum projection horizon: 5 years. DCF models should include a terminal year.
- Time periods must flow left-to-right, earliest to latest.

---

## 7. Formula Best Practices

- Avoid `INDIRECT`, `OFFSET`, and volatile functions (`NOW()`, `TODAY()`, `RAND()`) in core calculation ranges — they recalculate on every keystroke and hide dependencies.
- Avoid nested `IF` chains deeper than 3 levels; use helper columns or `IFS`/`SWITCH` instead.
- Circular references are prohibited unless explicitly required for an interest schedule; if used, document the reason, enable iterative calculation, and set a visible flag on the `Cover` sheet.
- Every external data link must be documented on `Assumptions` with source, date retrieved, and refresh procedure.

---

## 8. Scenario & Sensitivity Analysis

### 8.1 Scenarios

Define at least three scenarios on the `Assumptions` sheet:

| Scenario    | Description                                          |
|-------------|------------------------------------------------------|
| Base Case   | Management plan with reasonable assumptions          |
| Upside Case | Optimistic — e.g., 15–20% above base revenue        |
| Downside Case | Conservative — e.g., 15–20% below base revenue   |

Use a scenario selector (dropdown or toggle cell) that drives all inputs from a scenario table. Never duplicate sheets per scenario.

### 8.2 Sensitivity Tables

- Build two-way data tables on the `Valuation` sheet for key value drivers (e.g., revenue growth vs. EBITDA margin; discount rate vs. exit multiple).
- Label axes clearly and colour-grade output cells from low (red) to high (green).

---

## 9. Error Checks & Audit Trails

### 9.1 Balance Sheet Check

Every period must include:
```
Assets = Liabilities + Equity   →  difference must equal 0
```
Display this prominently at the top of the `Balance Sheet` tab. Colour the cell red if non-zero.

### 9.2 Cash Flow Check

```
Ending Cash = Opening Cash + Net Cash Flow   →  must tie to Balance Sheet cash
```

### 9.3 Error Flag Summary

Consolidate all error checks into a summary row on the `Cover` sheet showing PASS / FAIL status for each check.

---

## 10. Version Control

- Increment version on every material change: `v1.0`, `v1.1`, `v2.0`.
- Log changes in the `Cover` sheet change log table: date, author, version, description.
- Before sharing externally, save a dated archive copy to `/archive`.
- For models tracked in Git, check in only the `.xlsx` file; add a companion `model_changelog.md` for human-readable history.

---

## 11. Documentation & Assumptions

- Every assumption must cite its source (management guidance, market research, historical average, analyst consensus).
- Undocumented assumptions are considered errors.
- Include a data dictionary on the `Assumptions` sheet defining each line item and its unit.

---

## 12. Review Checklist

Before sharing or publishing, confirm all items below:

- [ ] All hard-coded inputs are on `Assumptions` only
- [ ] Balance sheet balances in every period
- [ ] Cash flow ties to balance sheet cash
- [ ] No broken external links
- [ ] Scenario selector works correctly for all three scenarios
- [ ] Error check summary on `Cover` shows all PASS
- [ ] Version and date updated on `Cover`
- [ ] Archive copy saved to `/archive`
- [ ] Sensitive data (PII, credentials) removed
- [ ] Model prints cleanly to PDF (print areas set)

---

## 13. Security & Access Control

- Models containing MNPI (material non-public information) must be stored in access-controlled locations only.
- Do not email unencrypted models containing confidential financial projections; use secure file sharing.
- Enable worksheet protection on formula cells before sharing; leave only `Assumptions` input cells unlocked.
- Audit access logs quarterly for models stored in shared drives.

---

*Last updated: 2026-04-28 — Thetan Financial Modeling Standards*
