# Task 03 — Build Charts and Dashboards

**Status:** In progress. Checkpoints 1–6 are complete based on learner-reported Excel results, including the annual summary and all three required charts. Dashboard KPI planning has started: the Sales Growth KPI is calculated at 76.8% using the reported 2016 and 2021 totals. The learner has not yet confirmed implementation of the KPI cards, chart placement, or dashboard formatting.

Updated working workbook: [Account Sales Data for Dashboard Practice — Working](../material/03-build-charts-and-dashboards/Account_Sales_Data_for_Dashboard_Practice_Working.xlsx)  
Original practice workbook: [Account Sales Data for Dashboard Practice](../material/03-build-charts-and-dashboards/Account_Sales_Data_for_Dashboard_Practice.xlsx)

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

**Learner-reported PivotTable results:**

| Account Type | Sum of 2021 Unit Sales |
|---|---:|
| Distributor | 3,861 |
| Enterprise | 4,051 |
| Online | 5,140 |
| Retail | 3,299 |
| Small Business | 927 |
| **Grand Total** | **17,278** |

The five category totals add to 17,278, matching the PivotTable Grand Total.

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

**Learner-reported PivotTable results:**

| Account Name | Sum of 2021 Unit Sales |
|---|---:|
| Cloudline Commerce | 1,450 |
| Juniper Digital | 1,180 |
| Orbit Commerce | 980 |
| MetroWorks Supply | 820 |
| Keystone Business Group | 748 |
| Ironwood Distribution | 705 |
| Silverline Distribution | 640 |
| Westbrook Distribution | 590 |
| Peakline Commerce | 560 |
| Granite Bay Wholesale | 548 |
| **Grand Total** | **8,221** |

The ten displayed values add to 8,221, matching the filtered PivotTable Grand Total. This is the combined sales of the Top 10 accounts, not the full-dataset total.

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

**Learner-reported annual summary results:**

| Year | Total Unit Sales |
|---|---:|
| 2016 | 9,773 |
| 2017 | 10,380 |
| 2018 | 11,115 |
| 2019 | 12,283 |
| 2020 | 14,193 |
| 2021 | 17,278 |

The reported totals increase in each successive year. The 2021 figure matches the reported Account Type PivotTable Grand Total. From 2016 to 2021, total unit sales increase by 7,505 units, approximately 76.8%. The learner-reported values are documented here; a separate independent workbook audit has not been performed.

### PivotTable validation

- [x] Both PivotTables have been created; the learner confirmed Sum aggregation and the reported results.
- [x] The account-type totals add to the reported Grand Total of 17,278.
- [x] The Top 10 PivotTable contains 10 accounts in descending order; its values add to the reported Grand Total of 8,221.
- [ ] Independently reconcile both PivotTables against the working workbook and source data.
- [x] Record the learner-reported annual summary values; the 2021 total matches the reported PivotTable Grand Total.
- [ ] Independently reconcile all six annual totals against the source columns.

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

**Learner-reported chart completion:** The learner confirms that all three charts have been created and saved in the updated working workbook:

- **Unit Sales by Year** — line chart, based on `Annual_Sales_Summary`.
- **Top 10 Accounts by 2021 Unit Sales** — horizontal bar chart; Cloudline Commerce is at the top of the ranking.
- **2021 Unit Sales by Account Type** — column chart; Online is the highest category.

The chart selections and key visual checks are learner-confirmed. The workbook's chart objects and plotted ranges have not been independently inspected.

## 6. Build the Dashboard worksheet

1. Insert a new worksheet and name it `Dashboard`.
2. Move or copy all three charts onto the Dashboard worksheet. The PivotTables and annual summary can remain on their own worksheets.
3. Arrange the visuals so they do not overlap:
   - Place **Unit Sales by Year** across the top to show the overall trend.
   - Place **Top 10 Accounts by 2021 Unit Sales** below or to the left.
   - Place **2021 Unit Sales by Account Type** beside the top-10 chart.
4. Use consistent titles, readable labels, and restrained formatting.
5. Check that all chart data is current and that the layout is understandable at a glance.


### KPI cards and dashboard formatting

The learner-provided annual summary is:

| Year | Total Unit Sales |
|---|---:|
| 2016 | 9,773 |
| 2017 | 10,380 |
| 2018 | 11,115 |
| 2019 | 12,283 |
| 2020 | 14,193 |
| 2021 | 17,278 |

Assuming the summary headers are in row 1, with 2016 in row 2 and 2021 in row 7, use this formula for the **Sales Growth (2016–2021)** KPI:

```excel
=('Annual_Sales_Summary'!$B$7-'Annual_Sales_Summary'!$B$2)/'Annual_Sales_Summary'!$B$2
```

Format the result as **Percentage** with one decimal place. The expected result is **76.8%**, representing an increase of 7,505 units from 2016 to 2021. The absolute references keep the source cells fixed if the formula is copied.

Suggested KPI cards:

| KPI card | Value | Source / calculation |
|---|---:|---|
| Total 2021 Sales | 17,278 units | Link to the Grand Total in `PT_Sales_by_Type`; use the actual Grand Total cell. |
| Top Account | Cloudline Commerce | Link to the first account in `PT_Top10_2021`; use the actual account-name cell. |
| Sales Growth (2016–2021) | 76.8% | Formula above, formatted as a percentage with one decimal place. |

Suggested worksheet layout:

| Dashboard element | Suggested cell area |
|---|---|
| Main title | `B2:N3` |
| Subtitle | `B4:N4` |
| Total 2021 Sales card | `B6:D8` |
| Top Account card | `F6:H8` |
| Sales Growth card | `J6:L8` |
| Annual sales line chart | `B10:N23` |
| Top 10 bar chart | `B25:H42` |
| Account type column chart | `J25:N42` |

Use consistent fills, borders, font sizes, and alignment across the three KPI cards. Keep the dashboard background uncluttered, hide worksheet gridlines, use a clear title and subtitle, and ensure the charts are readable and non-overlapping. Keep the PivotTables and annual summary on their own worksheets; the Dashboard should present linked KPI values and charts rather than duplicate source tables.

**Progress boundary:** The annual summary values and the Sales Growth calculation above are documented from learner-provided information. The formula is based on the stated row layout and has not been confirmed as entered into the workbook. Do not mark the KPI cards or dashboard formatting complete until the learner confirms the changes in Excel.

## 7. Final review and submission

- [ ] Dashboard worksheet exists and is named `Dashboard`.
- [ ] All three completed charts are on the Dashboard worksheet.
- [ ] Each chart uses the specified chart type and correct source data.
- [ ] Titles, labels, and legends are clear and useful.
- [ ] Charts do not overlap and are large enough to read.
- [ ] Chart totals reconcile with the source data or supporting summaries.
- [x] Save the updated working workbook in the repository (learner-confirmed).
- [ ] Complete the final review and submit the workbook through the Forage activity.

## Progress record

| Stage | Status |
|---|---|
| Inspect worksheet and confirm range / record count | Complete — learner-reported |
| Create Excel Table `AccountSales` | Complete — learner-reported |
| Create and validate the two 2021 sales PivotTables | Complete — learner-reported results and totals checked |
| Prepare annual sales summary | Complete — learner-reported values; 2021 cross-check matches the reported PivotTable total |
| Independently reconcile all six annual totals against source columns | Pending |
| Complete chart-type activity | Complete — learner-reported |
| Create and format the Unit Sales by Year line chart | Complete — learner-reported |
| Create and format the Top 10 bar chart | Complete — learner-reported; Cloudline Commerce is at the top |
| Create and format the account-type column chart | Complete — learner-reported; Online is the highest category |
| Save updated working workbook in the repository | Complete — repository file confirmed |
| Document Sales Growth KPI calculation (76.8%) from the reported annual summary | Complete — calculation documented; workbook entry not yet confirmed |
| Implement and format the three KPI cards | Pending — learner confirmation required |
| Assemble Dashboard worksheet and position all three charts | Pending — learner confirmation required |
| Format title, subtitle, chart layout, and card styling | Pending — learner confirmation required |
| Final validation and submission | Pending |

**Documentation rule:** Chart completion and visual checks in this guide are learner-reported; the updated workbook's chart objects and plotted ranges have not been independently inspected. The Sales Growth KPI calculation is documented from the learner-provided annual summary and assumes headers in row 1 with data in rows 2–7; its entry in Excel, the other KPI cards, the Dashboard worksheet, chart placement, and formatting remain unconfirmed until the learner verifies them.
