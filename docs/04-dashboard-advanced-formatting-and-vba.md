# Dashboard App-Style Formatting and VBA Enhancement

**Purpose:** Extend the finalized Task 03 dashboard with an application-inspired visual design and optional desktop-Excel navigation automation.

**Current status:** The finalized XLSX dashboard is learner-reported complete. The learner imported the original DashboardAppStyle.bas into Excel. A compatibility-hardened revision is now in the repository, but it must be re-imported into the learner's local XLSM before compilation and execution can be confirmed. Office Scripts source files are committed, but have not yet been pasted into or run from Excel's Scripts editor.

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
5. Back up the current local XLSM. Remove the older `DashboardAppStyle` module from Project Explorer (export it first if you want a backup), then import the latest repository version to avoid duplicate public macro names.
6. Confirm the worksheet names and KPI card ranges match your workbook.
7. In the VBE, run **Debug → Compile VBAProject**. If compilation succeeds, press Alt + F8, select ApplyDashboardAppStyle, then select Run.
8. Confirm the colors, text, chart panels, and navigation buttons. If runtime execution stops, note the stage, error number, full description, and highlighted line from the updated diagnostic message. Use RestoreDashboardEditView from Alt + F8 when you want to return to normal worksheet editing view.
9. Save the duplicate as Excel Macro-Enabled Workbook (*.xlsm).

If the card or title cell ranges differ from the target layout above, edit the corresponding range strings inside ApplyDashboardAppStyle and FormatKpiCard before running.

### Macro safety

- The module contains no Workbook_Open or auto-run procedure.
- It does not read or write files, call external services, send email, or invoke shell commands.
- It formats the Dashboard sheet and creates worksheet-navigation buttons only.
- It temporarily disables screen updating/events while running and restores their previous state.
- Keep Excel's notification-based macro security. Do not enable all macros globally; review and run only code you understand in a copy of your workbook.

## 6. Office Scripts version for Excel on the web

The repository includes two TypeScript scripts:

- [DashboardAppStyle_Web.ts](../material/03-build-charts-and-dashboards/DashboardAppStyle_Web.ts) — formats the title, subtitle, three KPI cards, and existing chart panels. It also adds internal worksheet hyperlinks in four navigation ranges if those ranges are blank.
- [RestoreDashboardEditView_Web.ts](../material/03-build-charts-and-dashboards/RestoreDashboardEditView_Web.ts) — restores gridlines and row/column headings for normal editing.

The Office Script navigation uses merged cell hyperlinks rather than VBA shape actions. It does not reposition charts or change their data sources. It also does not set the Excel window zoom.

### Add and run the formatting script

1. Open a copy of the finalized workbook in Excel for the web.
2. Select **Automate → New Script** (or open **View Scripts** and create a new script).
3. Paste the full contents of `DashboardAppStyle_Web.ts` into the code editor.
4. Save the script as **Dashboard App Style — Web**.
5. Confirm that the navigation ranges `B45:D46`, `E45:G46`, `H45:J46`, and `K45:M46` are available. The script leaves any range containing content unchanged.
6. Select **Run**, then inspect the dashboard and the script's log output.

Create a second script from `RestoreDashboardEditView_Web.ts`, name it **Restore Dashboard Editing View**, and run it whenever you want headings and gridlines visible again.

### Add a run button to the worksheet

1. Open the saved dashboard script from **Automate**.
2. In **Share this script**, enable **Associate with workbook**.
3. Select **Add button to worksheet**.
4. Place the generated button where it does not cover an existing chart or KPI card.

The run button triggers the script; the four navigation hyperlinks created by the script navigate to source worksheets. Microsoft's instructions are here: [Create a button to run an Office Script](https://support.microsoft.com/en-us/excel/office-scripts/create-a-button-to-run-an-office-script).

Office Scripts availability depends on the signed-in account, platform, license, and tenant settings. See the [Office Scripts introduction](https://support.microsoft.com/en-us/office/introduction-to-office-scripts-in-excel-9fbe283d-adb8-4f13-a75b-a81c6baf163a).

## 7. Validation checklist

- [ ] Original finalized XLSX remains unchanged.
- [x] VBA module imported into Excel (learner-confirmed).
- [ ] VBA project compiles without an error.
- [ ] Main formatting macro completes without an error.
- [ ] Office Scripts source saved and run in a supported Excel environment.
- [ ] Title and subtitle remain legible.
- [ ] All three KPI values and formulas remain unchanged.
- [ ] All three charts retain their original data and positions.
- [ ] Four navigation buttons open the intended worksheets.
- [ ] No buttons cover charts, KPI cards, or labels.
- [ ] Enhanced copy saved as XLSM.
- [ ] Original macro-free XLSX opens normally.

**Evidence rule:** VBA import is learner-confirmed; compilation and execution are pending. The Office Scripts are committed source code only until saved and run in Excel. Do not describe either automation route as tested until the learner verifies the results.


## Latest VBA compatibility update — 2026-11-04

The repository module was hardened after the learner reported a VBA runtime/compile error, but the exact Excel error message and highlighted line were not included. The update therefore addresses likely compatibility problems without claiming a confirmed root cause:

- Replaces the module's `mso...` shape/text constants with local numeric constants to reduce dependence on the Office object-library reference.
- Sets the shape object variable to `Nothing` before looking up an existing navigation shape.
- Adds a stage label and reports the error number and description if the main macro fails.

**Verification status:** the updated BAS source is committed and checked for the intended replacements; it has not been compiled or executed in desktop Excel. Re-import it into the local workbook copy, then compile and run. The exact highlighted line is still needed if an error persists.
