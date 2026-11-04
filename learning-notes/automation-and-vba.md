# Automation and VBA

**Status:** A reusable optional VBA enhancement has been prepared for the Task 03 dashboard. The module is committed as source code; it has not yet been imported into Excel or executed against the finalized workbook.

## Selected process

The dashboard already contains its KPI cards and three charts. The optional macro automates repetitive presentation work:

- Apply a consistent application-inspired visual theme.
- Format the dashboard title and subtitle.
- Apply a reusable style to the three KPI cards.
- Apply consistent chart-area styling and readable chart title/axis fonts.
- Add worksheet-navigation buttons for the source data, annual summary, Top 10 accounts, and account-type PivotTable.

## Source code

- [DashboardAppStyle.bas](../material/03-build-charts-and-dashboards/DashboardAppStyle.bas)
- [Dashboard design and VBA setup guide](../docs/04-dashboard-advanced-formatting-and-vba.md)

The module is written from scratch for this learning project. It uses Excel's worksheet, range, chart, and shape object models. The code is intentionally limited to formatting and navigation; it does not alter source data, KPI formulas, PivotTable definitions, chart ranges, or workbook connections.

## Import and run

1. Duplicate the finalized XLSX workbook first.
2. Open the duplicate in desktop Excel.
3. Press Alt + F11 to open the Visual Basic Editor.
4. Choose File → Import File... and select DashboardAppStyle.bas.
5. Review the module procedures and workbook/sheet/range names.
6. Return to Excel, press Alt + F8, select ApplyDashboardAppStyle, and choose Run.
7. Check the KPI cards, charts, and newly added navigation buttons.
8. Save the enhanced copy as Excel Macro-Enabled Workbook (.xlsm). Keep the original XLSX unchanged.

## Macro behavior and safety

- No auto-run procedures such as Workbook_Open or Auto_Open are included.
- The macro only formats the Dashboard worksheet, styles chart objects on that sheet, and creates or updates clearly named navigation buttons.
- Navigation buttons switch to existing worksheets by name.
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

## Test plan

| Test | Expected result | Status |
|---|---|---|
| Import module in a duplicate workbook | Module appears in the VBA project without replacing source sheets | Not run |
| Run ApplyDashboardAppStyle | Title, KPI cards, and chart panels receive the app-style formatting | Not run |
| Inspect navigation buttons | Each button opens its intended worksheets | Not run |
| Reopen the XLSM copy | Workbook opens with macros disabled until explicitly trusted; visual formatting remains | Not run |
| Reopen original XLSX | Original macro-free workbook and its data remain unchanged | Not run |

Do not mark the VBA enhancement complete until it has been run and the tests above have been recorded.
