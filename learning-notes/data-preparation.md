# Data Preparation and Exploration

<sub>Forage Academy · Excel Skills Job Simulation</sub>

**Status:** Task 01 completed — conditional formatting applied to the Account Sales Data workbook.

## Task 01 — Account Sales Data

Detailed findings, account-level records, conditional-formatting logic, and validation guidance are documented in [Task 01 — Conditional Formatting Findings](../docs/01-conditional-formatting-findings.md).

### Workbook structure

- Worksheet: `Account Sales Data`
- Records: 40 account rows
- Headers: `A1:M1`
- Data: `A2:M41`
- Relevant columns: Product 2 (`F`), 2021 Unit Sales (`L`), and 5 YR CAGR (`M`).

### Confirmed data-quality findings

| Finding | Result | Details |
|---|---:|---|
| Missing cells | 5 | D6, C12, E21, F29, D35 |
| Text error markers | 2 | M9 and M27 contain `Formula Error` |
| Product 2 = Yes | 23 | 57.5% of accounts |
| Product 2 = No | 16 | 40.0% of accounts |
| Product 2 blank | 1 | F29; do not interpret as a confirmed non-purchase |
| 5 YR CAGR ≥ 100% | 3 | Orbit Commerce, Juniper Digital, Cloudline Commerce |
| Negative 5 YR CAGR | 7 | See the detailed findings document |
| Top 10% of 2021 Unit Sales | 4 | Cloudline, Juniper, Orbit, and MetroWorks |

### Conditional-formatting rules applied

| Rule | Range | Result |
|---|---|---|
| Formula errors / text error markers | `A2:M41` | Purple fill, white font |
| Missing values | `A2:M41` | Yellow fill |
| Product 2 not cross-sold | `F2:F41` | Orange fill |
| 5 YR CAGR ≥ 100% (numeric only) | `M2:M41` | Green fill |
| Negative 5 YR CAGR | `M2:M41` | Red fill, white font |
| Top 10% of 2021 Unit Sales | `L2:L41` | Blue fill |

### Learning notes

- Missing data, a valid `No` response, and numeric zero represent different meanings and should not be combined.
- The dataset's `Formula Error` values are text markers, not native Excel error values. A custom conditional-formatting rule is needed to identify them reliably.
- For the green CAGR rule, `ISNUMBER` prevents text markers from being misclassified as percentages.
- Conditional-formatting rules should be reviewed in Manage Rules, including their priority and Applies To ranges.

## Related materials

- [Original workbook](../material/01-conditional-formatting/Account%20Sales%20Data%20for%20Analysis.xlsx)
- [Completed workbook](../material/01-conditional-formatting/Account_Sales_Data_Analysis_Completed.xlsx)
- [Full findings and account-level detail](../docs/01-conditional-formatting-findings.md)
