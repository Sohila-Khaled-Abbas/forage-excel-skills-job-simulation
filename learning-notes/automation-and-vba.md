# Automation and VBA

**Status:** Task 04's macro-recording and sorting/button exercises are complete based on learner confirmation. The initial macro was recorded, tested, and inspected; a Form Control button was added for it; `SortBy5YrCAGR` and `SortBy2021UnitSales` were recorded and assigned to separate buttons; both sorting buttons passed their tests; and the requested screenshots were assembled into a readable PDF. The macro-enabled workbook and PDF are present in the repository. The exact name/action of the first recorded macro remains undocumented, and Forage upload/submission is pending. Separately, the optional Task 03 dashboard VBA module still needs local compile/run verification after the compatibility-hardened version is re-imported. Office Scripts sources remain untested in Excel.

## Current Forage exercise — Task 04: Record and Run a Macro

The learner reports completing the initial macro recording, test, and VBA inspection, adding a Form Control button for that first macro, recording the two descending-sort macros, assigning each to its own button, testing both sorting buttons successfully, and creating/reopening a readable evidence PDF. The exact first macro name/action was not provided. The [initial Task 04 guide](../docs/05-record-and-run-a-macro.md) and [sorting macros and buttons guide](../docs/06-sort-macros-and-buttons.md) now record the completed work. The Task 04 automation workbook remains separate from the Task 03 dashboard workbook. Only upload and submission on Forage remain.

## Selected process

The dashboard already contains its KPI cards and three charts. The optional macro automates repetitive presentation work:

- Apply a consistent application-inspired visual theme.
- Format the dashboard title and subtitle.
- Apply a reusable style to the three KPI cards.
- Apply consistent chart-area styling and readable chart title/axis fonts.
- Add worksheet-navigation buttons for the source data, annual summary, Top 10 accounts, and account-type PivotTable.

## Source code

- [DashboardAppStyle.bas](../material/03-build-charts-and-dashboards/DashboardAppStyle.bas)
- [DashboardAppStyle_Web.ts](../material/03-build-charts-and-dashboards/DashboardAppStyle_Web.ts)
- [RestoreDashboardEditView_Web.ts](../material/03-build-charts-and-dashboards/RestoreDashboardEditView_Web.ts)
- [Dashboard design and VBA setup guide](../docs/04-dashboard-advanced-formatting-and-vba.md)

The module is written from scratch for this learning project. The latest repository revision replaces the used `mso...` drawing constants with local numeric constants to reduce Office object-library reference sensitivity, initializes the navigation shape reference before lookup, and reports the execution stage and error number/description. This is defensive compatibility hardening, not a confirmed root-cause fix: the exact Excel error and highlighted line were not supplied, and the revised module has not yet been compiled or run in the learner's local workbook. It uses Excel's worksheet, range, chart, and shape object models. The code is intentionally limited to formatting and navigation; it does not alter source data, KPI formulas, PivotTable definitions, chart ranges, or workbook connections.

## Import and run

1. Duplicate the finalized XLSX workbook first.
2. Open the duplicate in desktop Excel.
3. Press Alt + F11 to open the Visual Basic Editor.
4. Choose File → Import File... and select DashboardAppStyle.bas.
5. Back up the current local XLSM. In Project Explorer, right-click the older `DashboardAppStyle` module and choose **Remove** (export it first if you want a backup). Then import the updated repository version so that the VBA project does not contain two modules with the same public macro names.
6. In the Visual Basic Editor, select **Debug → Compile VBAProject**. If compilation succeeds, return to Excel, press Alt + F8, select ApplyDashboardAppStyle, and choose Run.
7. Check the KPI cards, charts, and newly added navigation buttons. If the macro fails at runtime, note the exact error number, full message, and highlighted VBA line. The updated module now identifies the stage where it stopped. Run RestoreDashboardEditView from Alt + F8 when you need normal row/column headings and gridlines again.
8. Save the enhanced copy as Excel Macro-Enabled Workbook (.xlsm). Keep the original XLSX unchanged.

## Macro behavior and safety

- No auto-run procedures such as Workbook_Open or Auto_Open are included.
- The macro only formats the Dashboard worksheet, styles chart objects on that sheet, and creates or updates clearly named navigation buttons.
- Navigation buttons switch to existing worksheets by name.
- RestoreDashboardEditView returns the Dashboard window to visible headings/gridlines and 100% zoom for editing.
- The module does not access the network, files, email, shell commands, or external processes.
- The code uses error handling and restores the prior values of ScreenUpdating and EnableEvents after execution.
- Do not enable all macros globally. Review this module and enable it only for the local workbook copy you control.

## Excel desktop versus Excel for the web

VBA macros can be opened and retained in a workbook in Excel for the web, but they cannot be created, edited, or run there. Use desktop Excel for this BAS module. If the workflow must run in a browser, consider Office Scripts, which uses TypeScript and supports Excel for the web where the Microsoft 365 tenant/license allows it.

Official references:
- [Work with VBA macros in Excel for the web](https://support.microsoft.com/en-us/excel/work-with-vba-macros-in-excel-for-the-web)
- [Save a macro-enabled workbook](https://support.microsoft.com/en-us/excel/save-a-macro)
- [Excel macro security settings](https://support.microsoft.com/en-us/excel/change-macro-security-settings-in-excel)
- [Introduction to Office Scripts](https://support.microsoft.com/en-us/office/introduction-to-office-scripts-in-excel-9fbe283d-adb8-4f13-a75b-a81c6baf163a)

## Office Scripts alternative

The web-compatible implementation is authored in TypeScript:

- `DashboardAppStyle_Web.ts` formats existing dashboard cells and charts, and adds internal worksheet hyperlinks only when the designated navigation ranges are empty.
- `RestoreDashboardEditView_Web.ts` restores gridlines and headings.

Paste each file into **Automate → New Script** in Excel for the web. Save and run it. To expose a script as a worksheet button, open the saved script, enable **Associate with workbook**, and select **Add button to worksheet**. The worksheet hyperlinks are independent of the script-run button.

Office Scripts do not run VBA `OnAction` macros; these are separate automation paths for similar presentation/navigation goals.

Official references:
- [ExcelScript Range API — hyperlinks](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts)
- [ExcelScript Chart API — formatting](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.chart?view=office-scripts)
- [Create a button to run an Office Script](https://support.microsoft.com/en-us/excel/office-scripts/create-a-button-to-run-an-office-script)

## Test plan

| Test | Expected result | Status |
|---|---|---|
| Import module in a duplicate workbook | Module appears in the VBA project without replacing source sheets | Complete — learner-confirmed |
| Compile VBAProject | VBA editor reports no compile errors | Pending |
| Run ApplyDashboardAppStyle | Title, KPI cards, and chart panels receive the app-style formatting | Not run |
| Inspect navigation buttons | Each button opens its intended worksheets | Not run |
| Reopen the XLSM copy | Workbook opens with macros disabled until explicitly trusted; visual formatting remains | Not run |
| Reopen original XLSX | Original macro-free workbook and its data remain unchanged | Not run |

Do not mark the VBA enhancement complete until it has been run and the tests above have been recorded.

| Run DashboardAppStyle_Web | Dashboard styling and blank-area navigation hyperlinks are applied | Not run |
| Run RestoreDashboardEditView_Web | Worksheet headings and gridlines return | Not run |
| Add Office Script run button | Workbook-associated script button appears and runs the selected script | Not run |


## Task 04 progress record

| Milestone | Status |
|---|---|
| Automation workbook downloaded | Complete — learner-confirmed |
| Macro-enabled working copy saved and committed to the repository | Complete — XLSM file verified in GitHub |
| Developer tab enabled for recording and Form Control insertion | Complete — implied by the learner-confirmed workflow |
| First macro recorded in This Workbook | Complete — learner-confirmed |
| Macro tested from Developer → Macros | Complete — learner-confirmed |
| Generated VBA reviewed | Complete — learner-confirmed |
| Form Control button added for the first recorded macro | Complete — learner-confirmed |
| `SortBy5YrCAGR` recorded and assigned its own button | Complete — learner-confirmed |
| `SortBy2021UnitSales` recorded and assigned its own button | Complete — learner-confirmed |
| Both sorting buttons tested successfully | Complete — learner-confirmed |
| Workbook and VBA screenshots captured; PDF reopened and checked | Complete — learner-confirmed |
| XLSM and PDF uploaded to Forage; task submitted | Pending confirmation |

See [Task 04 — Record and Run a Macro](../docs/05-record-and-run-a-macro.md) for the step-by-step workflow.


## Task 04 — Sorting macro implementation

The learner confirms that the two recorded macros were created and tested in the same Task 04 XLSM workbook:

| Macro | Sort key | Direction | Button label |
|---|---|---|---|
| `SortBy5YrCAGR` | `5 YR CAGR` | Largest to Smallest | Sort by 5 YR CAGR |
| `SortBy2021UnitSales` | `2021 Unit Sales` | Largest to Smallest | Sort by 2021 Unit Sales |

The entire contiguous data range must be included in each sort and the header row must remain fixed. The macros should be recorded through **Developer → Record Macro** and **Data → Sort**, not manually typed, because the exercise is teaching the Macro Recorder. Each macro gets its own **Form Control → Button**, assigned through the Assign Macro dialog. The learner confirms both buttons passed their tests and captured a screenshot of both buttons plus a screenshot of one macro procedure in the Visual Basic Editor. The screenshots were assembled into `Task_04_Macro_Buttons_Evidence.pdf`, which was reopened and checked for readability.

The earlier account-data workbook used range `A1:M41`, with `2021 Unit Sales` in column L and `5 YR CAGR` in column M. The Task 04 copy should be checked before recording; use its actual header labels and actual last row/column. If CAGR contains text placeholders or non-numeric entries, inspect the resulting descending order rather than changing the source markers. Microsoft documents that mixed text and numeric values have distinct sort order and blanks remain last: [Sort data in a range or table in Excel](https://support.microsoft.com/en-us/excel/sort-data-in-a-range-or-table-in-excel).

Official button guidance: [Assign a macro to a Form or a Control button](https://support.microsoft.com/en-us/excel/assign-a-macro-to-a-form-or-a-control-button).
