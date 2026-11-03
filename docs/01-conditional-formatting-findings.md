# Task 01 — Conditional Formatting Findings

<sub>Forage Academy · Excel Skills Virtual Experience Program · Completed exercise</sub>

## Exercise outcome

**Status:** Completed by the learner. The finished workbook has been added to the repository.

The exercise applies six conditional-formatting rules to the account-sales dataset. The purpose is to make data-quality issues visible, distinguish accounts not cross-sold with Product 2, and surface high- and negative-growth accounts alongside the strongest 2021 unit-sales performers.

### Workbook artifacts

| Artifact | Repository file |
|---|---|
| Original practice workbook | [Account Sales Data for Analysis.xlsx](../material/01-conditional-formatting/Account%20Sales%20Data%20for%20Analysis.xlsx) |
| Completed workbook | [Account_Sales_Data_Analysis_Completed.xlsx.xlsx](../material/01-conditional-formatting/Account_Sales_Data_Analysis_Completed.xlsx.xlsx) |

The completed workbook currently appears in the repository with a duplicate `.xlsx` suffix in its filename. The file itself is stored as an Excel workbook.

## 1. Dataset structure

- **Worksheet:** `Account Sales Data`
- **Header row:** `A1:M1`
- **Account records:** 40
- **Data range:** `A2:M41`
- **Fields:** 13

| Column | Field | Description |
|---|---|---|
| A | Account ID | Unique account identifier |
| B | Account Name | Account/customer name |
| C | Account Type | Business segment |
| D | Region | Geographic region |
| E | Product 1 | Product 1 cross-sell status |
| F | Product 2 | Product 2 cross-sell status |
| G:K | 2016–2020 Unit Sales | Historical annual unit sales |
| L | 2021 Unit Sales | Most recent annual unit sales in this dataset |
| M | 5 YR CAGR | Five-year compound annual growth rate |

The analysis below was calculated from the original workbook. Cell references use the original worksheet layout.

## 2. Data-quality findings

### Missing values

There are **5 blank cells** across the 40 account records.

| Cell | Account ID | Account | Missing field |
|---|---|---|---|
| D6 | AC-1005 | Summit Ridge Retail | Region |
| C12 | AC-1011 | Pioneer Business Group | Account Type |
| E21 | AC-1020 | Oak & Pine Wholesale | Product 1 |
| F29 | AC-1028 | Momentum Commerce | Product 2 |
| D35 | AC-1034 | Horizon Goods | Region |

These blanks are not interchangeable with zero or a negative response. In particular, the blank Product 2 value for Momentum Commerce is unknown and should not be classified as a confirmed non-purchase.

### CAGR error markers

Two cells in the 5 YR CAGR column contain the literal text **`Formula Error`**:

| Cell | Account ID | Account |
|---|---|---|
| M9 | AC-1008 | Evergreen Wholesale |
| M27 | AC-1026 | Clearview Products |

These values are text labels, not native Excel error values such as `#DIV/0!` or `#VALUE!`. A built-in “Errors” rule alone would not reliably identify them. The custom error-formatting logic therefore needs to recognize both native Excel errors and the dataset's text marker.

### Product 2 cross-sell status

| Product 2 value | Accounts | Share of 40 records |
|---|---:|---:|
| Yes | 23 | 57.5% |
| No | 16 | 40.0% |
| Blank | 1 | 2.5% |

The 16 accounts with Product 2 marked `No` are:

- AC-1002 — Cedar Lane Market
- AC-1004 — Harborview Goods
- AC-1006 — BluePeak Solutions
- AC-1009 — Golden Field Retail
- AC-1012 — Redwood Home Market
- AC-1015 — Stonebridge Supply
- AC-1017 — Crescent Operations
- AC-1020 — Oak & Pine Wholesale
- AC-1022 — Riverbend Market
- AC-1025 — Aspen Retail Group
- AC-1027 — Beacon Hill Market
- AC-1029 — Parkside Supply
- AC-1032 — Summit Digital Market
- AC-1034 — Horizon Goods
- AC-1037 — Granite Bay Wholesale
- AC-1039 — Keystone Business Group

The blank Product 2 value for AC-1028 is intentionally excluded from this list.

## 3. Growth and sales findings

### 5-year CAGR at or above 100%

There are **3 accounts** with numeric CAGR values of 100% or higher.

| Account ID | Account | 5 YR CAGR | 2021 Unit Sales |
|---|---|---:|---:|
| AC-1010 | Orbit Commerce | 100.83% | 980 |
| AC-1014 | Juniper Digital | 121.76% | 1,180 |
| AC-1019 | Cloudline Commerce | 140.56% | 1,450 |

These accounts also occupy the top three positions in 2021 unit sales. Their historical unit sales started from relatively low bases, so the CAGR values should be read together with absolute sales volumes rather than in isolation.

### Negative 5-year CAGR

There are **7 accounts** with negative CAGR values.

| Account ID | Account | 5 YR CAGR |
|---|---|---:|
| AC-1002 | Cedar Lane Market | -4.63% |
| AC-1009 | Golden Field Retail | -7.96% |
| AC-1015 | Stonebridge Supply | -4.75% |
| AC-1017 | Crescent Operations | -3.73% |
| AC-1022 | Riverbend Market | -3.88% |
| AC-1027 | Beacon Hill Market | -3.77% |
| AC-1034 | Horizon Goods | -4.56% |

### Top 10% of 2021 unit sales

With 40 numeric account records, the top 10% represents the four highest 2021 unit-sales values. There are no ties at the fourth-place cutoff.

| Rank | Account ID | Account | 2021 Unit Sales |
|---:|---|---|---:|
| 1 | AC-1019 | Cloudline Commerce | 1,450 |
| 2 | AC-1014 | Juniper Digital | 1,180 |
| 3 | AC-1010 | Orbit Commerce | 980 |
| 4 | AC-1021 | MetroWorks Supply | 820 |

## 4. Conditional-formatting implementation

The intended ranges exclude the header row and cover the 40 account records.

| Requirement | Applies to | Condition | Format |
|---|---|---|---|
| Formula errors / text error markers | `$A$2:$M$41` | Native error or text equals `Formula Error` | Purple fill, white font |
| Missing values | `$A$2:$M$41` | Blank cells | Yellow fill |
| Product 2 not cross-sold | `$F$2:$F$41` | Product 2 equals `No` | Orange fill |
| High CAGR | `$M$2:$M$41` | Numeric value greater than or equal to 1 (100%) | Green fill |
| Negative CAGR | `$M$2:$M$41` | Numeric value below 0 | Red fill, white font |
| Top 10% of 2021 unit sales | `$L$2:$L$41` | Top 10 percent | Blue fill |

### Formula-based rule examples

**Formula error rule** — use the top-left cell of the Applies To range as the relative reference:

```excel
=IF(ISERROR(A2),TRUE,A2="Formula Error")
```

**Product 2 not cross-sold**:

```excel
=$F2="No"
```

**CAGR at or above 100% — numeric values only**:

```excel
=IFERROR(AND(ISNUMBER($M2),$M2>=1),FALSE)
```

The numeric check is important. Text values can behave unexpectedly in comparison rules; without `ISNUMBER`, the literal `Formula Error` markers may be treated as qualifying values by a basic greater-than-or-equal rule.

A practical rule-priority safeguard is to place the error rule above overlapping numeric rules and enable **Stop If True** for that rule. The numeric-only CAGR condition remains the primary protection against text markers being classified as high growth.

## 5. Quality checks

Before saving the completed workbook, review the Conditional Formatting Rules Manager and confirm:

- All six rules are enabled.
- The Applies To ranges match the ranges in the table above.
- Row 1 is not included in the rules.
- The Product 2 rule highlights `No` but not the blank cell F29.
- The CAGR green rule only evaluates numeric values.
- The two text error markers remain purple rather than being classified as high growth.
- The top-10% rule is restricted to 2021 Unit Sales, not the entire table.
- The workbook is saved as a separate completed copy so the source remains unchanged.

## 6. Analytical takeaways

1. **Data completeness matters:** five missing fields occur across account type, region, and product cross-sell information. These should remain distinct from valid negative answers.
2. **Cross-sell opportunity:** 16 of 40 accounts (40%) are explicitly marked as not cross-sold with Product 2. The separate blank Product 2 value should be investigated rather than counted as a non-purchase.
3. **Growth concentration:** the three accounts with CAGR above 100% are also the three largest 2021 unit-sales accounts. Their strong growth should be interpreted alongside their low historical starting volumes.
4. **Declining accounts:** seven accounts show negative five-year CAGR and may warrant further commercial investigation.
5. **Error handling:** the text-based `Formula Error` markers demonstrate why data-quality rules must account for the actual representation of errors in a workbook.

## 7. Skills practiced

- Workbook inspection and range selection using Excel GUI and keyboard shortcuts.
- Built-in conditional formatting for blanks, thresholds, and top-percentile values.
- Formula-based conditional formatting with relative and absolute references.
- Numeric type checks to prevent text from being classified as numeric growth.
- Data-quality interpretation and cross-sell segmentation.
- Validation of rule scope, priority, and display behavior.

---

**Source:** Account Sales Data for Analysis workbook supplied for the [Forage Academy Excel Skills simulation](https://www.theforage.com/simulations/forage/sponsored-content-excel-skills-jf8r).

**Evidence note:** Dataset counts and account-level findings were independently recalculated from the original workbook. Completion of the formatting exercise is recorded based on the learner's confirmation and the completed workbook uploaded to this repository.
