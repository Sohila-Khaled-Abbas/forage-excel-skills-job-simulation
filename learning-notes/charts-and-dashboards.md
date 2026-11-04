# Charts and Dashboards

**Status:** Task 03 — Build Charts and Dashboards is in progress. The annual sales summary and all three required charts are complete based on learner-reported workbook updates. The learner confirms that the Dashboard now has three consistently formatted KPI cards (17,278; Cloudline Commerce; +76.8%), an aligned title and subtitle, and all three charts visible and readable. The remaining layout check is to confirm that no charts or KPI cards overlap; final validation remains pending.

## Task 03 — Prepare Your Data for Visualization

### Source workbook

[Updated working workbook — Account Sales Data for Dashboard Practice](../material/03-build-charts-and-dashboards/Account_Sales_Data_for_Dashboard_Practice_Working.xlsx)  
[Original practice workbook](../material/03-build-charts-and-dashboards/Account_Sales_Data_for_Dashboard_Practice.xlsx)

[Task 03 preparation, chart selection, and dashboard guide](../docs/03-chart-and-dashboard-preparation.md)

### Checkpoint 1 — Data preparation results

The learner reported the following results after preparing the workbook in Excel:

| Item | Result |
|---|---|
| Worksheet | `AccountSales` |
| Dataset range | `A1:M41` |
| Account records | 40 |
| Excel Table | `AccountSales` — created |

Checkpoints 1–6 are complete based on learner-reported Excel results. Both 2021 sales PivotTables have been created, with their displayed totals adding up correctly. The annual sales summary and all three required charts are complete. The learner confirms that Cloudline Commerce appears at the top of the Top 10 bar chart and Online is the highest account-type column. The updated working workbook is saved in the repository. The Dashboard is assembled with all three charts and formatted KPI cards; the no-overlap check and final review remain.

## PivotTable plan

### 1. 2021 Unit Sales by Account Type

| Area | Field |
|---|---|
| Rows | Account Type |
| Values | Sum of 2021 Unit Sales |

Purpose: compare 2021 sales across account types and provide the source for the required column chart.

Suggested worksheet: `PT_Sales_by_Type`.

**Learner-reported results:**

| Account Type | 2021 Unit Sales |
|---|---:|
| Distributor | 3,861 |
| Enterprise | 4,051 |
| Online | 5,140 |
| Retail | 3,299 |
| Small Business | 927 |
| **Grand Total** | **17,278** |

The category totals sum to the displayed Grand Total.

### 2. Top 10 Accounts by 2021 Unit Sales

| Area | Field |
|---|---|
| Rows | Account Name |
| Values | Sum of 2021 Unit Sales |

Sort descending and apply a Top 10 value filter based on Sum of 2021 Unit Sales.

Suggested worksheet: `PT_Top10_2021`.

**Learner-reported results:**

| Account Name | 2021 Unit Sales |
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

The ten selected account values add to the filtered Grand Total of 8,221. This Top 10 subtotal should not be confused with the full-dataset 2021 total.

### 3. Annual sales summary

The source has separate annual sales columns for 2016–2021. Build a two-column summary on `Annual_Sales_Summary`, with Year and Total Unit Sales. Structured-reference examples:

- 2016: `=SUM(AccountSales[2016 Unit Sales])`
- 2017: `=SUM(AccountSales[2017 Unit Sales])`
- 2018: `=SUM(AccountSales[2018 Unit Sales])`
- 2019: `=SUM(AccountSales[2019 Unit Sales])`
- 2020: `=SUM(AccountSales[2020 Unit Sales])`
- 2021: `=SUM(AccountSales[2021 Unit Sales])`

Use the exact Table headers in Excel if they differ. Check the totals against the corresponding source columns.

**Learner-reported annual summary results:**

| Year | Total Unit Sales |
|---|---:|
| 2016 | 9,773 |
| 2017 | 10,380 |
| 2018 | 11,115 |
| 2019 | 12,283 |
| 2020 | 14,193 |
| 2021 | 17,278 |

The reported series increases in every year. The 2021 value matches the reported account-type PivotTable total. Across the period, unit sales increased by 7,505, or approximately 76.8%. The six values are learner-reported; independent reconciliation to the source workbook remains outstanding.

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

- [x] Add a descriptive title to every chart (learner-confirmed).
- [x] Use readable axis labels and number formats (learner-confirmed).
- [x] Keep a legend only when it helps explain the visual (learner-confirmed).
- [x] Remove unnecessary clutter (learner-confirmed).
- [x] Ensure the top 10 bar chart is sorted from highest to lowest; Cloudline Commerce is at the top (learner-confirmed).
- [x] Check that every chart uses the correct summary data (learner-confirmed).

### Chart completion

The learner confirms that all three required charts have been created and saved in the updated working workbook:

- **Unit Sales by Year** — line chart using the six-year annual summary. The reported trend rises from 9,773 units in 2016 to 17,278 in 2021, an increase of 7,505 units (approximately 76.8%).
- **Top 10 Accounts by 2021 Unit Sales** — horizontal bar chart, with Cloudline Commerce (1,450 units) at the top.
- **2021 Unit Sales by Account Type** — column chart, with Online (5,140 units) as the highest category.

These chart details and visual checks are learner-reported. The workbook's chart objects and plotted ranges have not been independently inspected.


## Dashboard KPI cards and layout

The learner-provided annual summary is:

| Year | Total Unit Sales |
|---|---:|
| 2016 | 9,773 |
| 2017 | 10,380 |
| 2018 | 11,115 |
| 2019 | 12,283 |
| 2020 | 14,193 |
| 2021 | 17,278 |

Assuming `Annual_Sales_Summary` has headers in row 1, the 2016 value in `B2`, and the 2021 value in `B7`, the Dashboard's Sales Growth KPI formula is:

```excel
=('Annual_Sales_Summary'!$B$7-'Annual_Sales_Summary'!$B$2)/'Annual_Sales_Summary'!$B$2
```

Format the result as a percentage with one decimal place. The expected result is **76.8%**: sales increased by **7,505 units** between 2016 and 2021.

Recommended KPI cards:

| KPI | Expected display | Source |
|---|---:|---|
| Total 2021 Sales | 17,278 units | Link to the Grand Total cell in `PT_Sales_by_Type`. |
| Top Account | Cloudline Commerce | Link to the first account name in `PT_Top10_2021`. |
| Sales Growth (2016–2021) | 76.8% | Formula above. |

Suggested dashboard placement: title in `B2:N3`, subtitle in `B4:N4`, KPI cards in `B6:D8`, `F6:H8`, and `J6:L8`, the annual trend chart in `B10:N23`, and the two comparison charts side by side in `B25:H42` and `J25:N42`.

Style all cards consistently, use a restrained color palette, hide worksheet gridlines, and keep chart titles and labels readable. The PivotTables and annual summary should remain on their source worksheets; the dashboard should reference them.

**Progress note:** The learner confirms that the KPI cards are implemented and formatted, with values 17,278, Cloudline Commerce, and 76.8%. The dashboard title and subtitle are aligned, and all three charts are visible and readable. The formula assumes the annual summary headers are in row 1 with 2016 and 2021 in B2 and B7 respectively. The no-overlap inspection and final validation remain outstanding.

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
- [x] Create the account-type PivotTable and check that the displayed category values add to its Grand Total (learner-reported).
- [x] Create the descending Top 10 accounts PivotTable and check that the ten displayed values add to its Grand Total (learner-reported).
- [x] Prepare the annual sales summary for 2016–2021 and record the learner-reported values; 2021 matches the reported PivotTable total.
- [ ] Independently reconcile all six annual totals against the source columns.
- [x] Complete the chart-type activity (learner-reported).
- [x] Create the Unit Sales by Year line chart (learner-reported).
- [x] Create the Top 10 bar chart; Cloudline Commerce is at the top (learner-reported).
- [x] Create the account-type column chart; Online is the highest category (learner-reported).
- [x] Save the updated working workbook in the repository.
- [x] Document the Sales Growth KPI formula and expected result (76.8%) from the learner-reported annual summary; workbook entry not yet confirmed.
- [x] Implement and consistently format the three KPI cards (learner-confirmed).
- [x] Assemble the three charts on the `Dashboard` worksheet; all are visible and readable (learner-confirmed).
- [x] Align the dashboard title and subtitle (learner-confirmed).
- [ ] Verify that no KPI card or chart overlaps another object.
- [ ] Validate chart values, labels, layout, and totals.
- [ ] Save and submit the completed workbook.

**Documentation rule:** PivotTable, annual-summary, and chart details above are learner-reported. The annual summary's 2021 value matches the PivotTable total; all six source-column totals have not been independently audited. KPI cards, title/subtitle alignment, and chart visibility/readability are learner-confirmed. Chart objects and plotted ranges have not been independently inspected. The remaining dashboard layout check is to verify that the KPI cards and charts do not overlap.
