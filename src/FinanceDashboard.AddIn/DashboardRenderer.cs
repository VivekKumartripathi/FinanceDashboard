using System.Globalization;
using FinanceDashboard.Core.Calculation;
using FinanceDashboard.Core.Model;
using Microsoft.Office.Interop.Excel;

namespace FinanceDashboard.AddIn;

/// <summary>
/// Draws a <see cref="DashboardView"/> on the Dashboard sheet: a table of the numbers, headline tiles and one chart
/// built from the table. Only the Dashboard sheet is written to; Financials and Metrics are never touched.
/// </summary>
/// <remarks>
/// Layout: table from column A (Period, Value, one column per breakdown part, Note); tiles and chart to its right.
/// </remarks>
internal static class DashboardRenderer
{
    private const string SheetName = "Dashboard";
    private const string ChartName = "FinanceChart";
    private const string CommentaryName = "FinanceCommentary";
    private const float ChartWidth = 640;
    private const int HeaderRow = 4;

    // Excel colours are BGR: 0xBBGGRR.
    // Brand colours from the brief's footer: #240141 and #E6E2E8.
    private const int BrandDark = 0x410124;  // #240141: the latest period, titles
    private const int BrandLight = 0xE8E2E6; // #E6E2E8: earlier periods, header fills
    private const int Orange = 0x317DED;  // RGB(237, 125, 49): periods with missing data
    private const int Green = 0x3F8A2E;   // RGB(46, 138, 63): good news
    private const int Red = 0x2C2CC0;     // RGB(192, 44, 44): bad news
    private const int Grey = 0x707070;

    public static void Render(Workbook workbook, DashboardView view)
    {
        var app = workbook.Application;
        var screenUpdating = app.ScreenUpdating;
        app.ScreenUpdating = false;
        try
        {
            var sheet = GetOrCreateSheet(workbook);
            HideCfoNote(sheet);

            var layout = new Layout(view.Parts.Count, view.Series.Points.Count);
            var valueFormat = view.Series.Metric.IsPercent ? "0.0\"%\"" : "#,##0";

            sheet.Cells.Clear();
            WriteTable(sheet, view, layout, valueFormat);
            WriteKpis(sheet, view.Kpis, layout);
            var chartTop = WriteCommentary(sheet, view.Commentary, layout);
            DrawChart(sheet, view, layout, valueFormat, chartTop);
        }
        finally
        {
            app.ScreenUpdating = screenUpdating;
        }
    }

    /// <summary>Where everything goes; the table grows one column per breakdown part.</summary>
    private sealed class Layout(int partCount, int pointCount)
    {
        public const int PeriodColumn = 1;
        public const int ValueColumn = 2;
        public int PartColumn(int part) => 3 + part;
        public int NoteColumn => 3 + partCount;
        public int TileColumn => NoteColumn + 2;
        public int FirstDataRow => HeaderRow + 1;
        public int LastRow => HeaderRow + pointCount;

        public string Column(int column, int fromRow, int toRow) =>
            $"{Letter(column)}{fromRow}:{Letter(column)}{toRow}";

        public string DataColumn(int column) => Column(column, FirstDataRow, LastRow);
    }

    internal static string Title(DashboardView view)
    {
        var s = view.Series.Selection;
        var segment = s.Segment == Selection.All ? "All segments" : s.Segment;
        var country = s.Country == Selection.All ? "All countries" : s.Country;
        var title = $"{view.Series.Metric.DisplayName} · {segment} · {country} · {s.Frequency}";
        return view.BreakdownBy switch
        {
            Breakdown.Segment => title + " · by segment",
            Breakdown.Country => title + " · by country",
            _ => title
        };
    }

    private static string PartName(DashboardView view, SeriesResult part) =>
        view.BreakdownBy == Breakdown.Country ? part.Selection.Country : part.Selection.Segment;

    private static string TotalName(DashboardView view) => view.BreakdownBy switch
    {
        Breakdown.None => "Value",
        _ when !view.Series.Metric.IsPercent => "Total",
        Breakdown.Segment => "All segments",
        _ => "All countries"
    };

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

    /// <summary>
    /// The sample workbook has a "CFO note" text box where the table and chart go. Its request (make the latest
    /// period stand out) is built into the chart, so the box is hidden, not deleted, to keep the dashboard readable.
    /// </summary>
    private static void HideCfoNote(Worksheet sheet)
    {
        foreach (Shape shape in sheet.Shapes)
        {
            if (shape.Name == "CFO note")
                shape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
        }
    }

    /// <summary>Writes title, unit line and the table.</summary>
    private static void WriteTable(Worksheet sheet, DashboardView view, Layout layout, string valueFormat)
    {
        var points = view.Series.Points;
        var unit = view.Series.Metric.IsPercent ? "Percent" : "USD thousands";
        var warnings = points.Count(p => !p.IsComplete);

        var title = sheet.Range["A1"];
        title.Value2 = Title(view);
        title.Font.Bold = true;
        title.Font.Size = 14;
        title.Font.Color = BrandDark;
        sheet.Range["A2"].Value2 = warnings == 0
            ? unit
            : $"{unit} · ⚠ {warnings} period(s) have missing data, see Note";

        // Labels such as "Jan 25" would be turned into dates by Excel, so the column is formatted as text first.
        sheet.Range[layout.Column(Layout.PeriodColumn, HeaderRow, layout.LastRow)].NumberFormat = "@";
        sheet.Range[$"{Letter(Layout.ValueColumn)}{layout.FirstDataRow}:{Letter(layout.NoteColumn - 1)}{layout.LastRow}"]
            .NumberFormat = valueFormat;

        var cells = new object?[points.Count + 1, layout.NoteColumn];
        cells[0, Layout.PeriodColumn - 1] = "Period";
        cells[0, Layout.ValueColumn - 1] = TotalName(view);
        cells[0, layout.NoteColumn - 1] = "Note";
        for (var p = 0; p < view.Parts.Count; p++)
            cells[0, layout.PartColumn(p) - 1] = PartName(view, view.Parts[p]);

        for (var i = 0; i < points.Count; i++)
        {
            cells[i + 1, Layout.PeriodColumn - 1] = points[i].Label;
            cells[i + 1, Layout.ValueColumn - 1] = ToCell(points[i].Value);
            cells[i + 1, layout.NoteColumn - 1] = points[i].IsComplete ? null : string.Join(" ", points[i].Notes);
            for (var p = 0; p < view.Parts.Count; p++)
                cells[i + 1, layout.PartColumn(p) - 1] = ToCell(view.Parts[p].Points[i].Value);
        }
        sheet.Range[$"A{HeaderRow}:{Letter(layout.NoteColumn)}{layout.LastRow}"].Value2 = cells;

        var header = sheet.Range[$"A{HeaderRow}:{Letter(layout.NoteColumn)}{HeaderRow}"];
        header.Font.Bold = true;
        header.Font.Color = BrandDark;
        header.Interior.Color = BrandLight;
        header.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
        header.Borders[XlBordersIndex.xlEdgeBottom].Color = BrandDark;

        for (var i = 0; i < points.Count; i++)
        {
            if (!points[i].IsComplete)
                sheet.Range[$"A{layout.FirstDataRow + i}:{Letter(layout.NoteColumn)}{layout.FirstDataRow + i}"].Font.Color = Orange;
        }

        // The latest period is the one the CFO's office discusses first.
        var latestRow = sheet.Range[$"A{layout.LastRow}:{Letter(layout.NoteColumn - 1)}{layout.LastRow}"];
        latestRow.Font.Bold = true;
        latestRow.Interior.Color = BrandLight;

        ((Range)sheet.Columns[Layout.PeriodColumn]).ColumnWidth = 10;
        ((Range)sheet.Columns[Layout.ValueColumn]).ColumnWidth = 14;
        for (var p = 0; p < view.Parts.Count; p++)
            ((Range)sheet.Columns[layout.PartColumn(p)]).ColumnWidth = 12;
        ((Range)sheet.Columns[layout.NoteColumn]).ColumnWidth = 12;
        sheet.Range[$"{Letter(layout.NoteColumn + 1)}:{Letter(layout.TileColumn + 12)}"].ColumnWidth = sheet.StandardWidth;
    }

    private static object? ToCell(decimal? value) => value is decimal v ? (double)v : null;

    /// <summary>
    /// Three headline tiles above the chart: latest value, change on the previous period, change on last year.
    /// A change is green when it is good news for that metric (a fall in expenses is green).
    /// </summary>
    private static void WriteKpis(Worksheet sheet, KpiSummary kpis, Layout layout)
    {
        var first = layout.TileColumn;
        // Values such as "+3.4%" would be turned into numbers by Excel, so the tiles are text.
        sheet.Range[$"{Letter(first)}1:{Letter(first + 8)}3"].NumberFormat = "@";

        Tile(sheet, first, $"Latest · {kpis.LatestLabel}", DashboardPane.Format(kpis.Latest, kpis.IsPercent), BrandDark);
        ComparisonTile(sheet, first + 3, kpis.PreviousPeriod, kpis);
        ComparisonTile(sheet, first + 6, kpis.SamePeriodLastYear, kpis);

        if (kpis.HasMissingData)
        {
            var warning = sheet.Range[$"{Letter(first)}3"];
            warning.Value2 = "⚠ Includes periods with missing data";
            warning.Font.Color = Orange;
        }
    }

    private static void ComparisonTile(Worksheet sheet, int column, Comparison? comparison, KpiSummary kpis)
    {
        if (comparison is null) return;

        if (comparison.Change is not decimal change)
        {
            Tile(sheet, column, $"vs {comparison.Label}", "n/a", null);
            return;
        }

        var text = change.ToString("+0.0;-0.0;0.0", CultureInfo.InvariantCulture) + (kpis.IsPercent ? " pp" : "%");
        int? colour = change == 0m ? null : (change > 0m) == kpis.HigherIsBetter ? Green : Red;
        Tile(sheet, column, $"vs {comparison.Label}", text, colour);
    }

    private static void Tile(Worksheet sheet, int column, string label, string value, int? valueColour)
    {
        var labelCell = sheet.Range[$"{Letter(column)}1"];
        labelCell.Value2 = label;
        labelCell.Font.Size = 9;
        labelCell.Font.Color = Grey;

        var valueCell = sheet.Range[$"{Letter(column)}2"];
        valueCell.Value2 = value;
        valueCell.Font.Size = 16;
        valueCell.Font.Bold = true;
        if (valueColour is int c)
            valueCell.Font.Color = c;
    }

    /// <summary>
    /// The plain-English summary, in a text box above the chart so it reads first and is copied with it.
    /// Returns where the chart should start.
    /// </summary>
    private static double WriteCommentary(Worksheet sheet, string commentary, Layout layout)
    {
        var anchor = sheet.Range[$"{Letter(layout.TileColumn)}5"];
        Shape? box = null;
        foreach (Shape shape in sheet.Shapes)
        {
            if (shape.Name == CommentaryName)
                box = shape;
        }

        if (box is null)
        {
            box = sheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                (float)(double)anchor.Left, (float)(double)anchor.Top, ChartWidth, 40);
            box.Name = CommentaryName;
            box.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            box.Fill.ForeColor.RGB = BrandLight;
            box.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoTrue;
            box.TextFrame2.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeShapeToFitText;
            box.TextFrame2.MarginLeft = 8;
            box.TextFrame2.MarginRight = 8;
        }

        box.Left = (float)(double)anchor.Left;
        box.Top = (float)(double)anchor.Top;
        box.Width = ChartWidth;

        var text = box.TextFrame2.TextRange;
        text.Text = commentary;
        text.Font.Size = 10;
        text.Font.Fill.ForeColor.RGB = BrandDark;

        return box.Top + box.Height + 6;
    }

    private static void DrawChart(Worksheet sheet, DashboardView view, Layout layout, string valueFormat, double top)
    {
        var chart = GetOrCreateChart(sheet, layout, top).Chart;
        var isPercent = view.Series.Metric.IsPercent;
        var brokenDown = view.Parts.Count > 0;

        // Ratios are trends, so they are lines; amounts split into parts stack up to their total.
        chart.ChartType = isPercent ? XlChartType.xlLineMarkers
            : brokenDown ? XlChartType.xlColumnStacked
            : XlChartType.xlColumnClustered;
        chart.DisplayBlanksAs = XlDisplayBlanksAs.xlNotPlotted;
        chart.HasTitle = true;
        chart.ChartTitle.Text = Title(view);
        chart.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = BrandDark;
        chart.HasLegend = brokenDown;
        if (brokenDown)
            chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;

        var seriesCollection = (SeriesCollection)chart.SeriesCollection();
        while (seriesCollection.Count > 0)
            seriesCollection.Item(1).Delete();

        if (!brokenDown)
        {
            var series = AddSeries(seriesCollection, sheet, layout, Layout.ValueColumn, view.Series.Metric.DisplayName);
            StyleSingle(series, view.Series, isPercent, valueFormat);
        }
        else
        {
            for (var p = 0; p < view.Parts.Count; p++)
            {
                var part = view.Parts[p];
                var series = AddSeries(seriesCollection, sheet, layout, layout.PartColumn(p), PartName(view, part));
                StylePart(series, part, Tint(p, view.Parts.Count), isPercent, valueFormat);
            }

            // For a ratio the total cannot be read off stacked parts, so it is drawn as its own bold line.
            if (isPercent)
            {
                var total = AddSeries(seriesCollection, sheet, layout, Layout.ValueColumn, TotalName(view));
                StyleTotalLine(total, view.Series, valueFormat);
            }
        }

        var valueAxis = (Axis)chart.Axes(XlAxisType.xlValue);
        valueAxis.TickLabels.NumberFormat = valueFormat;
        valueAxis.HasTitle = true;
        valueAxis.AxisTitle.Text = isPercent ? "%" : "USD thousands";
        valueAxis.HasMajorGridlines = true;
        valueAxis.MajorGridlines.Format.Line.ForeColor.RGB = BrandLight;
    }

    private static Series AddSeries(SeriesCollection collection, Worksheet sheet, Layout layout, int column, string name)
    {
        var series = collection.NewSeries();
        series.Name = name;
        series.XValues = sheet.Range[layout.DataColumn(Layout.PeriodColumn)];
        series.Values = sheet.Range[layout.DataColumn(column)];
        return series;
    }

    /// <summary>One series: earlier periods in the light brand colour, the latest in the dark one.</summary>
    private static void StyleSingle(Series series, SeriesResult result, bool isPercent, string valueFormat)
    {
        if (isPercent)
        {
            series.Format.Line.ForeColor.RGB = BrandDark;
            series.Format.Line.Weight = 1.5f;
            series.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
            series.MarkerSize = 6;
            series.MarkerBackgroundColor = BrandLight;
            series.MarkerForegroundColor = BrandDark;
        }
        else
        {
            series.Format.Fill.ForeColor.RGB = BrandLight;
            series.Format.Line.ForeColor.RGB = BrandDark;
            series.Format.Line.Weight = 0.75f;
        }

        // Numbers on every point only while they still fit; the latest period is always labelled.
        series.HasDataLabels = result.Points.Count <= 12;
        if (series.HasDataLabels)
            ((DataLabels)series.DataLabels()).NumberFormat = valueFormat;

        ForEachPoint(series, result, (point, isLatest) =>
        {
            if (isLatest)
            {
                if (isPercent)
                {
                    point.MarkerSize = 10;
                    point.MarkerBackgroundColor = BrandDark;
                    point.MarkerForegroundColor = BrandDark;
                }
                else
                {
                    point.Format.Fill.ForeColor.RGB = BrandDark;
                }
                LabelLatest(point, valueFormat, BrandDark);
            }
        }, isPercent);
    }

    /// <summary>One part of a breakdown: its own shade of the brand colour; earlier periods faded so the latest stands out.</summary>
    private static void StylePart(Series series, SeriesResult part, int colour, bool isPercent, string valueFormat)
    {
        if (isPercent)
        {
            series.Format.Line.ForeColor.RGB = colour;
            series.Format.Line.Weight = 1.5f;
            series.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
            series.MarkerSize = 5;
            series.MarkerBackgroundColor = colour;
            series.MarkerForegroundColor = colour;
        }
        else
        {
            series.Format.Fill.ForeColor.RGB = colour;
        }

        ForEachPoint(series, part, (point, isLatest) =>
        {
            if (isPercent)
            {
                if (isLatest)
                {
                    point.MarkerSize = 9;
                    LabelLatest(point, valueFormat, colour);
                }
            }
            else if (!isLatest)
            {
                point.Format.Fill.Transparency = 0.45f;
            }
        }, isPercent);
    }

    private static void StyleTotalLine(Series series, SeriesResult total, string valueFormat)
    {
        series.Format.Line.ForeColor.RGB = BrandDark;
        series.Format.Line.Weight = 3f;
        series.Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;
        series.MarkerStyle = XlMarkerStyle.xlMarkerStyleNone;

        var last = total.Points.Count;
        if (total.Points[last - 1].Value is not null)
            LabelLatest((Point)series.Points(last), valueFormat, BrandDark);
    }

    /// <summary>Applies <paramref name="style"/> to each plotted point, then outlines periods with missing data in orange.</summary>
    private static void ForEachPoint(Series series, SeriesResult result, Action<Point, bool> style, bool isPercent)
    {
        for (var i = 0; i < result.Points.Count; i++)
        {
            var point = result.Points[i];
            if (point.Value is null) continue;

            var chartPoint = (Point)series.Points(i + 1);
            style(chartPoint, i == result.Points.Count - 1);

            if (!point.IsComplete)
            {
                if (isPercent)
                    chartPoint.MarkerForegroundColor = Orange;
                else
                {
                    chartPoint.Format.Line.ForeColor.RGB = Orange;
                    chartPoint.Format.Line.Weight = 2.25f;
                }
            }
        }
    }

    private static void LabelLatest(Point point, string valueFormat, int colour)
    {
        point.HasDataLabel = true;
        point.DataLabel.NumberFormat = valueFormat;
        point.DataLabel.Font.Bold = true;
        point.DataLabel.Font.Color = colour;
    }

    /// <summary>Shades from the dark brand colour towards the light one, so every part stays on-brand but distinct.</summary>
    internal static int Tint(int index, int count)
    {
        var t = count <= 1 ? 0.0 : 0.75 * index / (count - 1);
        int Channel(int shift) =>
            (int)Math.Round(((BrandDark >> shift) & 0xFF) + t * (((BrandLight >> shift) & 0xFF) - ((BrandDark >> shift) & 0xFF)));
        return Channel(0) | Channel(8) << 8 | Channel(16) << 16;
    }

    private static ChartObject GetOrCreateChart(Worksheet sheet, Layout layout, double top)
    {
        // The chart sits below the headline tiles and the commentary.
        var left = (double)sheet.Range[$"{Letter(layout.TileColumn)}1"].Left;
        var chartObjects = (ChartObjects)sheet.ChartObjects();
        foreach (ChartObject existing in chartObjects)
        {
            if (existing.Name == ChartName)
            {
                existing.Left = left;
                existing.Top = top;
                return existing;
            }
        }

        var created = chartObjects.Add(left, top, ChartWidth, 360);
        created.Name = ChartName;
        return created;
    }

    /// <summary>1 -> "A", 27 -> "AA".</summary>
    private static string Letter(int column)
    {
        var letters = "";
        for (; column > 0; column = (column - 1) / 26)
            letters = (char)('A' + (column - 1) % 26) + letters;
        return letters;
    }
}
