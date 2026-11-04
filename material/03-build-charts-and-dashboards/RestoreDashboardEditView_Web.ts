/**
 * RestoreDashboardEditView_Web
 * Restores row/column headings and gridlines after DashboardAppStyle_Web.
 */
function main(workbook: ExcelScript.Workbook) {
  const dashboard = workbook.getWorksheet("Dashboard");
  if (!dashboard) {
    throw new Error("Worksheet 'Dashboard' was not found.");
  }

  dashboard.setShowGridlines(true);
  dashboard.setShowHeadings(true);
  dashboard.activate();

  console.log("Dashboard is ready for normal editing.");
}
