# Dashboard App-Style Formatting and VBA Enhancement

**Purpose:** Extend the finalized Task 03 dashboard with an application-inspired visual design and optional desktop-Excel navigation automation.

**Current status:** The finalized XLSX dashboard is learner-reported complete. This guide and the importable VBA module are prepared as an optional enhancement. The VBA has not yet been imported or executed.

## 1. Keep the finalized workbook safe

Do not overwrite the finalized macro-free workbook.

1. Open the current [Task 03 working workbook](../material/03-build-charts-and-dashboards/Account_Sales_Data_for_Dashboard_Practice_Working.xlsx).
2. Save a local duplicate, for example Account_Sales_Dashboard_AppStyle.xlsm.
3. Keep the repository's finalized XLSX unchanged as the clean deliverable.
4. Import and test the VBA module only in the duplicate.

The VBA module is formatting/navigation code, not a data transformation. It does not rewrite source values, formulas, PivotTable definitions, or chart data sources.

## 2. Target dashboard design

Use a restrained, product-dashboard visual language rather than making every cell decorative.

| Area | Range / position | Purpose |
|---|---|---|
| Header band | B2:N3 | Dashboard identity |
| Subtitle | B4:N4 | Scope and context |
| Total 2021 Sales card | B6:D8 | Headline volume |
| Top Account card | F6:H8 | Leading account |
| Sales Growth card | J6:L8 | Change from 2016 to 2021 |
| Annual trend chart | B10:N23 | Six-year sales movement |
| Top 10 ranking chart | B25:H42 | Leading accounts |
| Account type chart | J25:N42 | Sales comparison by segment |

The range map is based on the dashboard layout used in the Task 03 guidance. If your finalized workbook uses different cell positions, adjust the ranges in the VBA module before running it. The macro intentionally does not reposition charts.

## 3. Visual system

Use the same color and typography rules across the full dashboard.

| Token | Color | Application |
|---|---|---|
| Canvas | #F8FAFC | Dashboard background |
| Header / primary text | #0F172A | Title and KPI values |
| Accent teal | #0F766E | Buttons and key data emphasis |
| Secondary accent | #0891B2 | Optional highlights |
| Card surface | #FFFFFF | KPI cards and chart panels |
| Border | #E2E8F0 | Subtle card and chart boundaries |
| Secondary text | #64748B | Labels and explanatory text |

Recommended typography:
- Aptos or Aptos Display for a modern Office-native appearance.
- 20–22 pt bold for the main title.
- 9–10 pt for card labels and subtitles.
- 16–20 pt for KPI values; reduce the Top Account value slightly if it wraps.
- 11–13 pt for chart titles and 9–10 pt for axis labels.

## 4. Manual app-style polish

You can apply the visual system without macros:

1. Select the Dashboard worksheet and hide Gridlines (View → Gridlines). For a cleaner presentation, hide Headings as well; turn them back on when editing.
2. Set the zoom to 85–90% so the dashboard reads as a single canvas.
3. Apply a very light surface fill to the dashboard canvas, keeping the title band dark navy.
4. Give all KPI cards the same white fill, subtle outside border, centered label/value, and equal height.
5. Give charts a white background, consistent title font, and no unnecessary outlines or legends.
6. Use Shape Format → Align and Distribute Horizontally/Vertically to line up objects precisely.
7. Add rounded-rectangle navigation buttons beneath the chart area if there is room. Suggested labels: SOURCE DATA, YEARLY SUMMARY, TOP 10 ACCOUNTS, and ACCOUNT TYPES.
8. Keep whitespace around every object. Avoid shadows on every element, rainbow chart colors, 3D charts, and heavy cell borders.

## 5. Optional VBA automation

The repository includes the importable VBA module:

[DashboardAppStyle.bas](../material/03-build-charts-and-dashboards/DashboardAppStyle.bas)

The module provides:

- ApplyDashboardAppStyle — formats the title, subtitle, KPI cards, chart panels, and dashboard view.
- GoToAccountSales — opens the source data worksheet.
- GoToAnnualSummary — opens the annual sales summary.
- GoToTopAccounts — opens the Top 10 PivotTable.
- GoToAccountTypes — opens the account-type PivotTable.
- ReturnToDashboard — returns to the Dashboard worksheet.
- RestoreDashboardEditView — restores row/column headings, gridlines, and 100% zoom for editing.

The main macro adds or updates four named navigation buttons beneath the lowest chart. It finds a suitable row below the existing charts rather than moving or resizing them.

### Import and run in Excel desktop

1. Open the duplicate workbook in the Excel desktop app.
2. If the Developer tab is hidden, enable it through File → Options → Customize Ribbon → Developer.
3. Press Alt + F11.
4. In the Visual Basic Editor, select File → Import File....
5. Select DashboardAppStyle.bas from the repository.
6. Review the module and confirm the worksheet names and KPI card ranges match your workbook.
7. Press Alt + F8, select ApplyDashboardAppStyle, then select Run.
8. Confirm the colors, text, chart panels, and navigation buttons. Use RestoreDashboardEditView from Alt + F8 when you want to return to normal worksheet editing view.
9. Save the duplicate as Excel Macro-Enabled Workbook (*.xlsm).

If the card or title cell ranges differ from the target layout above, edit the corresponding range strings inside ApplyDashboardAppStyle and FormatKpiCard before running.

### Macro safety

- The module contains no Workbook_Open or auto-run procedure.
- It does not read or write files, call external services, send email, or invoke shell commands.
- It formats the Dashboard sheet and creates worksheet-navigation buttons only.
- It temporarily disables screen updating/events while running and restores their previous state.
- Keep Excel's notification-based macro security. Do not enable all macros globally; review and run only code you understand in a copy of your workbook.

## 6. Excel for the web: choose Office Scripts, not VBA

Excel for the web can display worksheet formatting and charts, but it cannot create, edit, or run VBA macros. The VBA navigation buttons therefore work in desktop Excel only. Microsoft documents that limitation here: [Work with VBA macros in Excel for the web](https://support.microsoft.com/en-us/excel/work-with-vba-macros-in-excel-for-the-web).

For browser-based automation, use Office Scripts when it is available in your Microsoft 365 environment. Office Scripts uses TypeScript and supports recording or authoring scripts, replaying them from Automate → View Scripts, and associating a script button with a workbook:
- [Introduction to Office Scripts](https://support.microsoft.com/en-us/office/introduction-to-office-scripts-in-excel-9fbe283d-adb8-4f13-a75b-a81c6baf163a)
- [Record actions as Office Scripts](https://support.microsoft.com/en-us/excel/office-scripts/record-your-actions-as-office-scripts)
- [Create a button to run an Office Script](https://support.microsoft.com/en-us/excel/office-scripts/create-a-button-to-run-an-office-script)

Office Scripts availability can depend on the Microsoft 365 license and tenant/admin settings.

## 7. Validation checklist

- [ ] Original finalized XLSX remains unchanged.
- [ ] VBA module imported into a duplicate workbook.
- [ ] Main formatting macro completes without an error.
- [ ] Title and subtitle remain legible.
- [ ] All three KPI values and formulas remain unchanged.
- [ ] All three charts retain their original data and positions.
- [ ] Four navigation buttons open the intended worksheets.
- [ ] No buttons cover charts, KPI cards, or labels.
- [ ] Enhanced copy saved as XLSM.
- [ ] Original macro-free XLSX opens normally.

**Evidence rule:** Until the macro has been run and tested, describe it as a prepared enhancement module, not as a feature already embedded in the finalized workbook.
