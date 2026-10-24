using System.Globalization;
using FinanceDashboard.Core.Calculation;
using FinanceDashboard.Core.Model;
using Microsoft.Office.Interop.Excel;

namespace FinanceDashboard.AddIn;

/// <summary>
/// Draws a <see cref="SeriesResult"/> on the Dashboard sheet: a small table of the numbers and one chart built from it.
/// Only the Dashboard sheet is written to; Financials and Metrics are never touched.
/// </summary>
internal static class DashboardRenderer
{
    private const string SheetName = "Dashboard";
    private const string ChartName = "FinanceChart";
    private const int HeaderRow = 4;

    // Excel colours are BGR: 0xBBGGRR.
    private const int Orange = 0x317DED;  // RGB(237, 125, 49): periods with missing data
    private const int Blue = 0xBD7244;    // RGB(68, 114, 189)
    private const int Green = 0x3F8A2E;   // RGB(46, 138, 63): good news
    private const int Red = 0x2C2CC0;     // RGB(192, 44, 44): bad news
    private const int Grey = 0x707070;

    public static void Render(Workbook workbook, SeriesResult result, KpiSummary kpis)
    {
        var app = workbook.Application;
        var screenUpdating = app.ScreenUpdating;
        app.ScreenUpdating = false;
        try
        {
            var sheet = GetOrCreateSheet(workbook);
            var valueFormat = result.Metric.IsPercent ? "0.0\"%\"" : "#,##0";
            var lastRow = WriteTable(sheet, result, valueFormat);
            WriteKpis(sheet, kpis);
            DrawChart(sheet, result, valueFormat, lastRow);
        }
        finally
        {
            app.ScreenUpdating = screenUpdating;
        }
    }

    internal static string Title(SeriesResult result)
    {
        var s = result.Selection;
        var segment = s.Segment == Selection.All ? "All segments" : s.Segment;
        var country = s.Country == Selection.All ? "All countries" : s.Country;
        return $"{result.Metric.DisplayName} · {segment} · {country} · {s.Frequency}";
    }

    private static Worksheet GetOrCreateSheet(Workbook workbook)
    {
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.Name == SheetName)
                return sheet;
        }

        var created = (Worksheet)workbook.Worksheets.Add(Before: workbook.Worksheets[1]);
        created.Name = SheetName;
        return created;
    }

    /// <summary>Writes title, unit line and the Period / Value / Note table. Returns the last table row.</summary>
    private static int WriteTable(Worksheet sheet, SeriesResult result, string valueFormat)
    {
        var points = result.Points;
        var unit = result.Metric.IsPercent ? "Percent" : "USD thousands";
        var warnings = points.Count(p => !p.IsComplete);

        sheet.Range["A:C"].Clear();

        sheet.Range["A1"].Value2 = Title(result);
        sheet.Range["A1"].Font.Bold = true;
        sheet.Range["A1"].Font.Size = 14;
        sheet.Range["A2"].Value2 = warnings == 0
            ? unit
            : $"{unit} · ⚠ {warnings} period(s) have missing data, see Note";

        // Labels such as "Jan 25" would be turned into dates by Excel, so the column is formatted as text first.
        var lastRow = HeaderRow + points.Count;
        sheet.Range[$"A{HeaderRow}:A{lastRow}"].NumberFormat = "@";
        sheet.Range[$"B{HeaderRow + 1}:B{lastRow}"].NumberFormat = valueFormat;

        var cells = new object?[points.Count + 1, 3];
        cells[0, 0] = "Period";
        cells[0, 1] = "Value";
        cells[0, 2] = "Note";
        for (var i = 0; i < points.Count; i++)
        {
            cells[i + 1, 0] = points[i].Label;
            cells[i + 1, 1] = points[i].Value is decimal v ? (double)v : null;
            cells[i + 1, 2] = points[i].IsComplete ? null : string.Join(" ", points[i].Notes);
        }
        sheet.Range[$"A{HeaderRow}:C{lastRow}"].Value2 = cells;

        var header = sheet.Range[$"A{HeaderRow}:C{HeaderRow}"];
        header.Font.Bold = true;
        header.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

        for (var i = 0; i < points.Count; i++)
        {
            if (!points[i].IsComplete)
                sheet.Range[$"A{HeaderRow + 1 + i}:C{HeaderRow + 1 + i}"].Font.Color = Orange;
        }

        ((Range)sheet.Columns["A"]).ColumnWidth = 10;
        ((Range)sheet.Columns["B"]).ColumnWidth = 14;
        ((Range)sheet.Columns["C"]).ColumnWidth = 12;
        return lastRow;
    }

    /// <summary>
    /// Three headline tiles above the chart: latest value, change on the previous period, change on last year.
    /// A change is green when it is good news for that metric (a fall in expenses is green).
    /// </summary>
    private static void WriteKpis(Worksheet sheet, KpiSummary kpis)
    {
        var area = sheet.Range["E1:M3"];
        area.Clear();
        // Values such as "+3.4%" would be turned into numbers by Excel, so the tiles are text.
        area.NumberFormat = "@";

        Tile(sheet, KpiColumns[0], $"Latest · {kpis.LatestLabel}", DashboardPane.Format(kpis.Latest, kpis.IsPercent), null);
        ComparisonTile(sheet, KpiColumns[1], "vs", kpis.PreviousPeriod, kpis);
        ComparisonTile(sheet, KpiColumns[2], "vs", kpis.SamePeriodLastYear, kpis);

        if (kpis.HasMissingData)
        {
            var warning = sheet.Range[$"{KpiColumns[0]}3"];
            warning.Value2 = "⚠ Includes periods with missing data";
            warning.Font.Color = Orange;
        }
    }

    private static readonly string[] KpiColumns = ["E", "H", "K"];

    private static void ComparisonTile(Worksheet sheet, string column, string prefix, Comparison? comparison, KpiSummary kpis)
    {
        if (comparison?.Change is not decimal change)
        {
            Tile(sheet, column, comparison is null ? "" : $"{prefix} {comparison.Label}", comparison is null ? "" : "n/a", null);
            return;
        }

        var text = kpis.IsPercent
            ? change.ToString("+0.0;-0.0;0.0", CultureInfo.InvariantCulture) + " pp"
            : change.ToString("+0.0;-0.0;0.0", CultureInfo.InvariantCulture) + "%";

        int? colour = change == 0m ? null : (change > 0m) == kpis.HigherIsBetter ? Green : Red;
        Tile(sheet, column, $"{prefix} {comparison.Label}", text, colour);
    }

    private static void Tile(Worksheet sheet, string column, string label, string value, int? valueColour)
    {
        var labelCell = sheet.Range[$"{column}1"];
        labelCell.Value2 = label;
        labelCell.Font.Size = 9;
        labelCell.Font.Color = Grey;

        var valueCell = sheet.Range[$"{column}2"];
        valueCell.Value2 = value;
        valueCell.Font.Size = 16;
        valueCell.Font.Bold = true;
        if (valueColour is int c)
            valueCell.Font.Color = c;
    }

    private static void DrawChart(Worksheet sheet, SeriesResult result, string valueFormat, int lastRow)
    {
        var chart = GetOrCreateChart(sheet).Chart;
        var isPercent = result.Metric.IsPercent;

        chart.ChartType = isPercent ? XlChartType.xlLineMarkers : XlChartType.xlColumnClustered;
        chart.DisplayBlanksAs = XlDisplayBlanksAs.xlNotPlotted;
        chart.HasLegend = false;
        chart.HasTitle = true;
        chart.ChartTitle.Text = Title(result);

        var seriesCollection = (SeriesCollection)chart.SeriesCollection();
        while (seriesCollection.Count > 0)
            seriesCollection.Item(1).Delete();

        var series = seriesCollection.NewSeries();
        series.Name = result.Metric.DisplayName;
        series.XValues = sheet.Range[$"A{HeaderRow + 1}:A{lastRow}"];
        series.Values = sheet.Range[$"B{HeaderRow + 1}:B{lastRow}"];

        if (isPercent)
        {
            series.Format.Line.ForeColor.RGB = Blue;
            series.MarkerBackgroundColor = Blue;
            series.MarkerForegroundColor = Blue;
        }
        else
        {
            series.Format.Fill.ForeColor.RGB = Blue;
        }

        // Numbers on the bars only while they still fit.
        series.HasDataLabels = result.Points.Count <= 12;
        if (series.HasDataLabels)
            ((DataLabels)series.DataLabels()).NumberFormat = valueFormat;

        // Periods with missing data are drawn in orange, matching the table.
        for (var i = 0; i < result.Points.Count; i++)
        {
            var point = result.Points[i];
            if (point.IsComplete || point.Value is null) continue;

            var chartPoint = (Point)series.Points(i + 1);
            if (isPercent)
            {
                chartPoint.MarkerBackgroundColor = Orange;
                chartPoint.MarkerForegroundColor = Orange;
            }
            else
            {
                chartPoint.Format.Fill.ForeColor.RGB = Orange;
            }
        }

        var valueAxis = (Axis)chart.Axes(XlAxisType.xlValue);
        valueAxis.TickLabels.NumberFormat = valueFormat;
        valueAxis.HasTitle = true;
        valueAxis.AxisTitle.Text = isPercent ? "%" : "USD thousands";
    }

    private static ChartObject GetOrCreateChart(Worksheet sheet)
    {
        // The chart sits below the headline tiles in rows 1-3.
        var anchor = sheet.Range["E5"];
        var chartObjects = (ChartObjects)sheet.ChartObjects();
        foreach (ChartObject existing in chartObjects)
        {
            if (existing.Name == ChartName)
            {
                existing.Left = (double)anchor.Left;
                existing.Top = (double)anchor.Top;
                return existing;
            }
        }

        var created = chartObjects.Add((double)anchor.Left, (double)anchor.Top, 640, 360);
        created.Name = ChartName;
        return created;
    }
}
