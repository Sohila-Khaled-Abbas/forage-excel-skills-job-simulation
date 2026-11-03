# Task 02 — Formulas and Functions Findings

<sub>Forage Academy · Excel Skills Virtual Experience Program · Completed exercise</sub>

## Exercise outcome

**Status:** Completed by the learner. The formula-practice workbook has been added to the repository.

This exercise applies Excel aggregation, conditional logic, error handling, and summary functions to account-level sales data. The calculations cover annual sales performance from 2016 through 2021, sales classification against a workbook-provided threshold, percentage change over the period, and two summary metrics.

### Workbook artifacts

| Artifact | Repository file |
|---|---|
| Original practice workbook | [Excel_Formula_Practice.xlsx](../material/02-formulas-and-functions/Excel_Formula_Practice.xlsx) |
| Completed workbook | [Excel_Formula_Practice_Completed.xlsx](../material/02-formulas-and-functions/Excel_Formula_Practice_Completed.xlsx) |

The Task 02 folder contains both the original practice workbook and the separately named completed workbook. The completed copy was uploaded after the learner confirmed completion. The workbook has not been independently recalculated or audited in this documentation pass, so the formulas below document the task logic and the confirmed Sales Level expression rather than claiming an independent formula-by-formula verification.

## 1. Account-level calculations

The account data uses annual unit-sales columns from 2016 to 2021. With the worksheet layout used during the exercise, the annual values are in columns G:L and the first account is on row 2.

| Calculation | Formula for first account | Purpose |
|---|---|---|
| Total Unit Sales | `=SUM(G2:L2)` | Adds annual unit sales from 2016 through 2021. |
| Average Annual Unit Sales | `=AVERAGE(G2:L2)` | Calculates average annual unit sales across the six years. |
| Highest Annual Unit Sales | `=MAX(G2:L2)` | Returns the highest annual unit-sales value. |
| Lowest Annual Unit Sales | `=MIN(G2:L2)` | Returns the lowest annual unit-sales value. |
| Sales Level | `=IF(L2>=$U$2,"High","Standard")` | Classifies the account against the threshold. |
| Sales Change | `=IFERROR((L2-G2)/G2,"Check data")` | Calculates percentage change from 2016 to 2021 and handles calculation errors. |

### Sales Level: absolute threshold reference

The workbook's Sales Level threshold is **500**, stored in cell `U2`.

```excel
=IF(L2>=$U$2,"High","Standard")
```

- `L2` is the account's 2021 Unit Sales.
- `$U$2` is an absolute reference, keeping the threshold fixed when the formula is filled down.
- A value of 500 or higher returns `High`; a value below 500 returns `Standard`.

### Sales Change: percentage change and error handling

The percentage-change calculation compares the 2021 value with the 2016 baseline:

```excel
=IFERROR((L2-G2)/G2,"Check data")
```

The calculation is `(2021 Unit Sales - 2016 Unit Sales) / 2016 Unit Sales`. Format the result column as Percentage. If the calculation produces an error, such as dividing by a zero 2016 baseline, Excel returns the text `Check data` instead.

## 2. Summary calculations

| Summary metric | Formula for the account range | Purpose |
|---|---|---|
| Accounts with Product 2 | `=COUNTIF(F2:F41,"Yes")` | Counts cells marked `Yes` in the Product 2 column. |
| Retail 2021 Unit Sales | `=SUMIF(C2:C41,"Retail",L2:L41)` | Adds 2021 Unit Sales for accounts whose Account Type is `Retail`. |

The `SUMIF` criteria range and sum range cover the same account rows: Account Type is checked in column C and the corresponding 2021 Unit Sales values are added from column L.

## 3. Validation checklist

- [ ] Confirm all six account-level formulas are filled down for every account.
- [ ] Check that Total Unit Sales, Average Annual Unit Sales, Highest Annual Unit Sales, and Lowest Annual Unit Sales use the six annual values from 2016–2021.
- [ ] Confirm Sales Level uses the absolute reference `$U$2` and the threshold value is 500.
- [ ] Check boundary behavior: exactly 500 should return `High`; values below 500 should return `Standard`.
- [ ] Format Sales Change as Percentage and check that a zero or invalid 2016 baseline returns `Check data`.
- [ ] Confirm `COUNTIF` counts only Product 2 cells containing `Yes`.
- [ ] Confirm `SUMIF` sums 2021 Unit Sales only for `Retail` accounts.
- [ ] Save the completed workbook in the Task 02 material folder.

## 4. Skills practiced

- Row-level aggregation with `SUM`, `AVERAGE`, `MAX`, and `MIN`.
- Conditional classification with `IF`.
- Absolute cell references for fixed business thresholds.
- Percentage-change calculations across time.
- Error handling with `IFERROR`.
- Criteria-based counting and summation with `COUNTIF` and `SUMIF`.
- Filling formulas down and checking calculation behavior.

---

**Source:** Excel Formula Practice workbook supplied for the [Forage Academy Excel Skills simulation](https://www.theforage.com/simulations/forage/sponsored-content-excel-skills-jf8r).

**Evidence note:** The learner confirmed completing and saving the exercise, and the workbook is present in the repository. Formula syntax and column references are documented from the exercise guidance and workbook layout used during the walkthrough; a separate independent audit of the workbook's stored formulas and calculated values has not been performed.
