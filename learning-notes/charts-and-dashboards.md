# Charts and Dashboards

**Status:** Task 03 — Build Charts and Dashboards is in progress.

## Task 03 — Prepare Your Data for Visualization

### Source workbook

[Account Sales Data for Dashboard Practice](../material/03-build-charts-and-dashboards/Account_Sales_Data_for_Dashboard_Practice.xlsx)

[Task 03 preparation, chart selection, and dashboard guide](../docs/03-chart-and-dashboard-preparation.md)

### Checkpoint 1 — Data preparation results

The learner reported the following results after preparing the workbook in Excel:

| Item | Result |
|---|---|
| Worksheet | `AccountSales` |
| Dataset range | `A1:M41` |
| Account records | 40 |
| Excel Table | `AccountSales` — created |

Checkpoint 1 is complete. PivotTables, charts, and dashboard assembly remain in progress.

## PivotTable plan

### 1. 2021 Unit Sales by Account Type

| Area | Field |
|---|---|
| Rows | Account Type |
| Values | Sum of 2021 Unit Sales |

Purpose: compare 2021 sales across account types and provide the source for the required column chart.

Suggested worksheet: `PT_Sales_by_Type`.

### 2. Top 10 Accounts by 2021 Unit Sales

| Area | Field |
|---|---|
| Rows | Account Name |
| Values | Sum of 2021 Unit Sales |

Sort descending and apply a Top 10 value filter based on Sum of 2021 Unit Sales.

Suggested worksheet: `PT_Top10_2021`.

### 3. Annual sales summary

The source has separate annual sales columns for 2016–2021. Build a two-column summary on `Annual_Sales_Summary`, with Year and Total Unit Sales. Structured-reference examples:

- 2016: `=SUM(AccountSales[2016 Unit Sales])`
- 2017: `=SUM(AccountSales[2017 Unit Sales])`
- 2018: `=SUM(AccountSales[2018 Unit Sales])`
- 2019: `=SUM(AccountSales[2019 Unit Sales])`
- 2020: `=SUM(AccountSales[2020 Unit Sales])`
- 2021: `=SUM(AccountSales[2021 Unit Sales])`

Use the exact Table headers in Excel if they differ. Check the totals against the corresponding source columns.

## Chart selection and dashboard specification

| Required visual | Business question | Chart type | Supporting data |
|---|---|---|---|
| Unit Sales by Year | How did total unit sales change from 2016 through 2021? | Line | Annual_Sales_Summary |
| Top 10 Accounts by 2021 Unit Sales | Which 10 accounts had the highest 2021 unit sales? | Bar | PT_Top10_2021 |
| 2021 Unit Sales by Account Type | How do 2021 unit sales compare across account types? | Column | PT_Sales_by_Type |

Chart selection principles:
- Use a line chart to show change over time.
- Use a column chart to compare a small number of categories.
- Use a horizontal bar chart to rank several items, especially when category names are long.
- Use stacked charts only when showing meaningful component contributions to a total.

## Chart formatting checklist

- [ ] Add a descriptive title to every chart.
- [ ] Use readable axis labels and number formats.
- [ ] Keep a legend only when it helps explain the visual.
- [ ] Remove unnecessary clutter.
- [ ] Ensure the top 10 bar chart is sorted from highest to lowest.
- [ ] Check that every chart uses the correct summary data.

## Dashboard layout

Create a worksheet named `Dashboard` and place all three charts on it. A practical layout is the annual trend across the top, with the top-10 ranking and account-type comparison side by side below. Keep charts aligned, readable, and non-overlapping.

## Evidence and business takeaways

| Item | Record after completing it |
|---|---|
| PivotTable worksheet names | Actual worksheet names |
| PivotTable field placements | Rows, Columns, Values, Filters |
| Aggregations | Sum, Count, Average, etc. |
| Annual totals | Actual 2016–2021 values |
| Chart sources and types | Confirm each chart uses its specified source and type |
| Dashboard review | Confirm titles, labels, spacing, and reconciliation |

Record only insights supported by the finished PivotTables and charts. Explain why each finding matters to the intended audience and note any limitations.

## Progress checklist

- [x] Inspect the `AccountSales` worksheet and identify range `A1:M41` (40 account records; learner-reported).
- [x] Convert the dataset into an Excel Table named `AccountSales` (learner-reported).
- [ ] Create and validate the account-type PivotTable.
- [ ] Create and validate the top-10 accounts PivotTable.
- [ ] Prepare and validate annual sales totals for 2016–2021.
- [ ] Complete the chart-type activity.
- [ ] Create the required line, bar, and column charts.
- [ ] Assemble all three charts on the `Dashboard` worksheet.
- [ ] Validate chart values, labels, layout, and totals.
- [ ] Save and submit the completed workbook.

**Documentation rule:** Do not mark PivotTables, charts, or the dashboard complete until they have been created and checked in Excel.
