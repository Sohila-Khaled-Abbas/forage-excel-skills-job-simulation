# Task 03 — Build Charts and Dashboards

**Status:** In progress. Checkpoint 1 — data preparation — is complete based on the learner's reported Excel results.

Workbook: [Account Sales Data for Dashboard Practice](../material/03-build-charts-and-dashboards/Account_Sales_Data_for_Dashboard_Practice.xlsx)

## Objective

Prepare the account sales dataset, summarize it with PivotTables, choose charts that answer the required questions, and bring the three visuals together on a single Dashboard worksheet.

## 1. Prepare a working copy

1. Open the supplied workbook in Excel.
2. Press `F12` and save a working copy in the Task 03 folder, for example `Account_Sales_Data_for_Dashboard_Practice_Working.xlsx`.
3. Keep the supplied source workbook unchanged.
4. Confirm the source worksheet, headers, data types, record count, and full range.

## 2. Convert the dataset into an Excel Table

1. Select any cell inside the dataset.
2. Press `Ctrl + T`.
3. Confirm the selected range includes all intended data rows and columns.
4. Ensure **My table has headers** is checked.
5. Select **OK**.
6. Open **Table Design** and set the Table Name to `AccountSales`.
7. Check that filter arrows appear and the Table contains all records.

### Checkpoint 1 — Data preparation results

**Learner-reported results:**

| Item | Result |
|---|---|
| Worksheet | `AccountSales` |
| Dataset range | `A1:M41` |
| Account records | 40 |
| Excel Table name | `AccountSales` |
| Table created | Yes |

Checkpoint 1 is complete. The worksheet and Table share the reported name `AccountSales`; Excel treats worksheet names and Table names as separate objects.

## 3. Create and validate PivotTables

For each PivotTable, select a cell in the `AccountSales` Table, choose **Insert → PivotTable**, confirm the source is `AccountSales`, and choose **New Worksheet**.

### PivotTable 1 — 2021 Unit Sales by Account Type

This summary will feed the required column chart.

| PivotTable area | Field |
|---|---|
| Rows | Account Type |
| Values | 2021 Unit Sales |

- Confirm the Values field is **Sum of 2021 Unit Sales**, not Count.
- If Excel chooses Count: open **Value Field Settings → Sum**.
- Format values as Number with a thousands separator and zero decimal places.
- Rename the worksheet `PT_Sales_by_Type`.

### PivotTable 2 — Top 10 Accounts by 2021 Unit Sales

| PivotTable area | Field |
|---|---|
| Rows | Account Name |
| Values | 2021 Unit Sales |

- Confirm the Values field uses **Sum**.
- Sort the values from largest to smallest.
- Apply **Value Filters → Top 10** to Account Name, using the Sum of 2021 Unit Sales field.
- Keep the top 10 accounts only.
- Rename the worksheet `PT_Top10_2021`.

### Annual sales summary — 2016 through 2021

The source stores each year in a separate column. A compact two-column summary is useful for the line chart:

| Year | Total Unit Sales |
|---|---|
| 2016 | `=SUM(AccountSales[2016 Unit Sales])` |
| 2017 | `=SUM(AccountSales[2017 Unit Sales])` |
| 2018 | `=SUM(AccountSales[2018 Unit Sales])` |
| 2019 | `=SUM(AccountSales[2019 Unit Sales])` |
| 2020 | `=SUM(AccountSales[2020 Unit Sales])` |
| 2021 | `=SUM(AccountSales[2021 Unit Sales])` |

Create this summary on a separate worksheet named `Annual_Sales_Summary`. Use the actual Table column headers if they differ. Format totals as whole numbers. This is a simple summary table using structured-reference SUM formulas; it avoids treating the six annual columns as six unrelated chart series.

### PivotTable validation

- [ ] Both PivotTables use `AccountSales` as their source.
- [ ] Sales measures use **Sum**, not Count.
- [ ] The account-type totals reconcile to the source 2021 Unit Sales total.
- [ ] The top 10 PivotTable contains exactly 10 accounts, sorted descending.
- [ ] Annual totals recalculate correctly and match the source columns.

## 4. Choose the right chart

Choose the chart based on the analytical question rather than decoration.

| Chart type | Best for | Example |
|---|---|---|
| Line chart | Showing how values change over time | Total unit sales from 2016 through 2021 |
| Column chart | Comparing values across categories | 2021 unit sales by account type |
| Bar chart | Comparing or ranking several items | Top 10 accounts by 2021 unit sales |
| Stacked column or bar | Comparing totals while showing component contributions | Category totals split into multiple components |

Do not add chart types that do not help answer a business question. The required dashboard uses one line chart, one bar chart, and one column chart.

## 5. Create and format the charts

For a standard chart:

1. Select the summary data, including its headers.
2. Choose **Insert** and select the suitable chart type.
3. Add a descriptive chart title.
4. Make axis labels clear and readable.
5. Keep the legend only when it adds meaning; remove it for a single-series chart if it is redundant.
6. Remove unnecessary visual elements and resize the chart so labels and values are easy to read.

For a PivotChart, select a cell inside the relevant PivotTable, then choose **Insert → PivotChart** and select the chart type.

### Required chart specifications

| Chart title | Data source | Chart type |
|---|---|---|
| Unit Sales by Year | `Annual_Sales_Summary` (2016–2021 totals) | Line |
| Top 10 Accounts by 2021 Unit Sales | `PT_Top10_2021` | Bar |
| 2021 Unit Sales by Account Type | `PT_Sales_by_Type` | Column |

For the top 10 bar chart, keep the categories in descending sales order. Make account labels readable; use a horizontal bar chart rather than squeezing long account names into columns.

## 6. Build the Dashboard worksheet

1. Insert a new worksheet and name it `Dashboard`.
2. Move or copy all three charts onto the Dashboard worksheet. The PivotTables and annual summary can remain on their own worksheets.
3. Arrange the visuals so they do not overlap:
   - Place **Unit Sales by Year** across the top to show the overall trend.
   - Place **Top 10 Accounts by 2021 Unit Sales** below or to the left.
   - Place **2021 Unit Sales by Account Type** beside the top-10 chart.
4. Use consistent titles, readable labels, and restrained formatting.
5. Check that all chart data is current and that the layout is understandable at a glance.

## 7. Final review and submission

- [ ] Dashboard worksheet exists and is named `Dashboard`.
- [ ] All three required charts are on the Dashboard worksheet.
- [ ] Each chart uses the specified chart type and correct source data.
- [ ] Titles, labels, and legends are clear and useful.
- [ ] Charts do not overlap and are large enough to read.
- [ ] Chart totals reconcile with the source data or supporting summaries.
- [ ] Save the completed working workbook and submit it through the Forage activity.

## Progress record

| Stage | Status |
|---|---|
| Inspect worksheet and confirm range / record count | Complete — learner-reported |
| Create Excel Table `AccountSales` | Complete — learner-reported |
| Create and validate PivotTables | In progress |
| Complete chart-type activity | Pending |
| Create and format three charts | Pending |
| Assemble Dashboard worksheet | Pending |
| Final validation and submission | Pending |

**Documentation rule:** Only mark PivotTables, charts, or the dashboard complete after they have been created and checked in Excel.
