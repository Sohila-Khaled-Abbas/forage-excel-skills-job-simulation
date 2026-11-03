# Formulas and Analysis

<sub>Forage Academy · Excel Skills Job Simulation</sub>

**Status:** Task 02 — Put Formulas Into Practice completed by the learner.

## Task 02 — Account Sales Formula Practice

The exercise calculates account-level sales metrics for 2016–2021, classifies accounts against a threshold, calculates percentage change, and produces two summary metrics.

### Account-level formulas

Using the worksheet layout from the exercise walkthrough (annual sales in columns G:L, first account on row 2):

| Metric | Formula | Notes |
|---|---|---|
| Total Unit Sales | `=SUM(G2:L2)` | Sum of annual unit sales across 2016–2021. |
| Average Annual Unit Sales | `=AVERAGE(G2:L2)` | Average annual unit sales across the six years. |
| Highest Annual Unit Sales | `=MAX(G2:L2)` | Maximum annual unit sales. |
| Lowest Annual Unit Sales | `=MIN(G2:L2)` | Minimum annual unit sales. |
| Sales Level | `=IF(L2>=$U$2,"High","Standard")` | Compares 2021 sales with the threshold in U2 (500). |
| Sales Change | `=IFERROR((L2-G2)/G2,"Check data")` | Percentage change from 2016 to 2021; returns a text fallback on error. |

### Key formula concepts

- **Absolute reference:** `$U$2` locks the threshold cell when Sales Level is copied down. The account reference `L2` remains relative.
- **Threshold behavior:** 2021 Unit Sales greater than or equal to 500 returns `High`; values below 500 return `Standard`.
- **Percentage change:** `(2021 - 2016) / 2016` expresses the change relative to the 2016 baseline. Format numeric outputs as Percentage.
- **Error handling:** If the Sales Change calculation errors (for example, because the 2016 baseline is zero), `IFERROR` returns `Check data`.

### Summary formulas

| Metric | Formula | Notes |
|---|---|---|
| Accounts with Product 2 | `=COUNTIF(F2:F41,"Yes")` | Counts cells marked Yes in Product 2. |
| Retail 2021 Unit Sales | `=SUMIF(C2:C41,"Retail",L2:L41)` | Sums 2021 Unit Sales where Account Type is Retail. |

### Validation considerations

- Check that the formula ranges cover every account row and exclude headers.
- Confirm the Sales Level threshold remains `$U$2` after filling down.
- Check the boundary case: exactly 500 should return `High`.
- Confirm Sales Change is displayed as a percentage and returns `Check data` when the calculation errors.
- Confirm the COUNTIF criterion is exactly `Yes`.
- Ensure the SUMIF criteria and sum ranges align row by row.

**Evidence:** [Task 02 workbook](../material/02-formulas-and-functions/Excel_Formula_Practice.xlsx) · [Task 02 findings](../docs/02-formulas-and-functions-findings.md)

**Verification note:** Completion is recorded from the learner's confirmation and the workbook uploaded to the repository. The formulas and expected behavior are documented from the task guidance and walkthrough; the workbook's stored formulas and calculated totals have not been independently audited.
