/**
 * DashboardAppStyle_Web
 * App-inspired formatting for the Task 03 Dashboard in Excel environments
 * that support Office Scripts.
 *
 * The script formats existing cells/charts, and adds internal worksheet
 * hyperlinks only where its navigation ranges are blank.
 * It does not change KPI values/formulas, source data, PivotTables,
 * chart data sources, or chart positions.
 *
 * Run against a copy of the finalized workbook first.
 */
function main(workbook: ExcelScript.Workbook) {
  const dashboard = workbook.getWorksheet("Dashboard");
  if (!dashboard) {
    throw new Error("Worksheet 'Dashboard' was not found.");
  }

  const colors = {
    canvas: "#F8FAFC",
    navy: "#0F172A",
    teal: "#0F766E",
    white: "#FFFFFF",
    border: "#E2E8F0",
    secondary: "#64748B"
  };

  dashboard.setShowGridlines(false);
  dashboard.setShowHeadings(false);

  const canvas = dashboard.getRange("A1:N50");
  canvas.getFormat().getFill().setColor(colors.canvas);
  canvas.getFormat().getFont().setName("Aptos");
  canvas.getFormat().getFont().setColor(colors.navy);

  // Style existing text only; never replace the title or subtitle content.
  const title = dashboard.getRange("B2:N3");
  title.getFormat().getFill().setColor(colors.navy);
  title.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  title.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  title.getFormat().setWrapText(true);
  title.getFormat().getFont().setName("Aptos Display");
  title.getFormat().getFont().setSize(20);
  title.getFormat().getFont().setBold(true);
  title.getFormat().getFont().setColor(colors.white);

  const subtitle = dashboard.getRange("B4:N4");
  subtitle.getFormat().getFill().setColor(colors.canvas);
  subtitle.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  subtitle.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  subtitle.getFormat().getFont().setName("Aptos");
  subtitle.getFormat().getFont().setSize(10);
  subtitle.getFormat().getFont().setColor(colors.secondary);

  styleKpiCard(dashboard, "B6:D8", "B6:D6", "B7:D8", 20, colors);
  styleKpiCard(dashboard, "F6:H8", "F6:H6", "F7:H8", 15, colors);
  styleKpiCard(dashboard, "J6:L8", "J6:L6", "J7:L8", 20, colors);

  // Keep existing chart placement and data sources intact.
  const charts = dashboard.getCharts();
  for (const chart of charts) {
    try {
      chart.getFormat().getFill().setSolidColor(colors.white);
      chart.getFormat().setRoundedCorners(true);

      const chartTitle = chart.getTitle();
      if (chartTitle.getVisible()) {
        const font = chartTitle.getFormat().getFont();
        font.setName("Aptos");
        font.setSize(12);
        font.setBold(true);
        font.setColor(colors.navy);
      }

      try {
        const axes = chart.getAxes();
        axes.getCategoryAxis().getFormat().getFont().setName("Aptos");
        axes.getCategoryAxis().getFormat().getFont().setSize(9);
        axes.getCategoryAxis().getFormat().getFont().setColor(colors.secondary);
        axes.getValueAxis().getFormat().getFont().setName("Aptos");
        axes.getValueAxis().getFormat().getFont().setSize(9);
        axes.getValueAxis().getFormat().getFont().setColor(colors.secondary);
      } catch (axisError) {
        console.log("Axis styling skipped for " + chart.getName() + ": " + String(axisError));
      }
    } catch (chartError) {
      console.log("Part of chart styling skipped for " + chart.getName() + ": " + String(chartError));
    }
  }

  // The navigation strip is placed below the documented chart layout.
  // A range is left untouched if it already contains any value.
  const navigationItems: { address: string; sheetName: string; label: string }[] = [
    { address: "B45:D46", sheetName: "AccountSales", label: "SOURCE DATA" },
    { address: "E45:G46", sheetName: "Annual_Sales_Summary", label: "YEARLY SUMMARY" },
    { address: "H45:J46", sheetName: "PT_Top10_2021", label: "TOP 10 ACCOUNTS" },
    { address: "K45:M46", sheetName: "PT_Sales_by_Type", label: "ACCOUNT TYPES" }
  ];

  let linksCreated = 0;
  let linksSkipped = 0;

  for (const item of navigationItems) {
    const destination = workbook.getWorksheet(item.sheetName);
    if (!destination) {
      console.log("Missing destination worksheet: " + item.sheetName);
      linksSkipped++;
      continue;
    }

    const buttonRange = dashboard.getRange(item.address);
    if (!isBlank(buttonRange)) {
      console.log("Navigation area already has content; left unchanged: " + item.address);
      linksSkipped++;
      continue;
    }

    buttonRange.merge(false);
    buttonRange.getFormat().getFill().setColor(colors.teal);
    buttonRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    buttonRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    buttonRange.getFormat().getFont().setName("Aptos");
    buttonRange.getFormat().getFont().setSize(9);
    buttonRange.getFormat().getFont().setBold(true);
    buttonRange.getFormat().getFont().setColor(colors.white);

    dashboard.getRange(item.address.split(":")[0]).setHyperlink({
      documentReference: item.sheetName + "!A1",
      screenTip: "Open the " + item.sheetName + " worksheet",
      textToDisplay: item.label
    });

    linksCreated++;
  }

  console.log(
    "Dashboard styling complete. Charts: " + charts.length +
    ". Navigation links created: " + linksCreated +
    ". Areas skipped: " + linksSkipped + "."
  );
}

function styleKpiCard(
  worksheet: ExcelScript.Worksheet,
  cardAddress: string,
  labelAddress: string,
  valueAddress: string,
  valueFontSize: number,
  colors: { canvas: string; navy: string; teal: string; white: string; border: string; secondary: string }
) {
  const card = worksheet.getRange(cardAddress);
  const format = card.getFormat();

  format.getFill().setColor(colors.white);
  format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  format.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  format.setWrapText(true);

  const edges: ExcelScript.BorderIndex[] = [
    ExcelScript.BorderIndex.edgeTop,
    ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight
  ];

  for (const edge of edges) {
    const border = format.getRangeBorder(edge);
    border.setStyle(ExcelScript.BorderLineStyle.continuous);
    border.setColor(colors.border);
    border.setWeight(ExcelScript.BorderWeight.thin);
  }

  const label = worksheet.getRange(labelAddress);
  label.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  label.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  label.getFormat().setWrapText(true);
  label.getFormat().getFont().setName("Aptos");
  label.getFormat().getFont().setSize(9);
  label.getFormat().getFont().setBold(true);
  label.getFormat().getFont().setColor(colors.secondary);

  const value = worksheet.getRange(valueAddress);
  value.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  value.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  value.getFormat().setWrapText(true);
  value.getFormat().getFont().setName("Aptos Display");
  value.getFormat().getFont().setSize(valueFontSize);
  value.getFormat().getFont().setBold(true);
  value.getFormat().getFont().setColor(colors.navy);
}

function isBlank(range: ExcelScript.Range): boolean {
  const values = range.getValues();
  for (let row = 0; row < values.length; row++) {
    for (let column = 0; column < values[row].length; column++) {
      const value = values[row][column];
      if (value !== "" && value !== null) {
        return false;
      }
    }
  }
  return true;
}
