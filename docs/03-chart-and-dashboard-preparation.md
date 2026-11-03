# Task 03 — Build Charts and Dashboards: Preparation

**Status:** In progress.

Workbook: [Account Sales Data for Dashboard Practice](../material/03-build-charts-and-dashboards/Account_Sales_Data_for_Dashboard_Practice.xlsx)

## Objective

Prepare the supplied account sales dataset for visualization. Preserve the source data, convert the dataset into an Excel Table, and create PivotTables on separate worksheets.

## 1. Prepare a working copy

1. Open the workbook in Excel.
2. Press `F12` and save a working copy in this folder, for example `Account_Sales_Data_for_Dashboard_Practice_Working.xlsx`.
3. Review worksheet names, headers, record count, data types, blank rows, and the actual dataset range.
4. Keep the supplied source workbook unchanged.

## 2. Convert the dataset into an Excel Table

1. Select any cell inside the dataset.
2. Press `Ctrl + T`.
3. Confirm the selected range includes all intended data rows and columns.
4. Ensure **My table has headers** is checked.
5. Select **OK**.
6. Open **Table Design** and rename the table `AccountSales`.
7. Check that filter arrows appear and the table includes all records.

## 3. Create PivotTables

1. Select a cell inside `AccountSales`.
2. Choose **Insert → PivotTable**.
3. Confirm the source is `AccountSales`.
4. Choose **New Worksheet**, then select **OK**.
5. Add fields to Rows, Columns, Values, and Filters according to a business question.
6. Check each Values field uses the intended aggregation (usually Sum for unit sales) and an appropriate number format.
7. Keep PivotTables separate from the source data.

Suggested exploratory questions:

| Question | Possible PivotTable layout |
|---|---|
| How do sales vary by account type? | Account Type in Rows; sales measure in Values (Sum). |
| Which regions contribute the most sales? | Region in Rows; sales measure in Values (Sum). |
| Which accounts have the highest sales? | Account Name in Rows; sales measure in Values (Sum), sorted descending. |
| How do annual sales compare? | Add annual sales fields to Values, or reshape wide annual columns into Year and Unit Sales using Power Query if a single Year field is needed. |

Use the actual field names available in the workbook. These are exploratory suggestions, not a prescribed dashboard design.

## 4. Validate

- [ ] Save a working copy separately from the supplied source.
- [ ] Confirm the Excel Table includes all intended rows and columns.
- [ ] Confirm `AccountSales` is the table name.
- [ ] Confirm PivotTables use the table as their source.
- [ ] Check aggregations and number formats.
- [ ] Reconcile PivotTable totals against the source data.
- [ ] Keep the original data intact.

## 5. Record your results

After completing the exercise, document the actual table name, worksheet, source range, record count, PivotTable names, field placements, aggregation choices, and any data-quality issues. Add your own screenshots if useful and safe to share.

**Scope note:** This is a preparation guide. It does not claim the Table or PivotTables have already been created or validated. Continue to the chart-type guidance after this preparation stage.
