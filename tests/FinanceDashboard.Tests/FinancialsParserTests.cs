using FinanceDashboard.Core.Model;
using FinanceDashboard.Core.Parsing;

namespace FinanceDashboard.Tests;

public class FinancialsParserTests
{
    /// <summary>Builds a 1-based array, the same shape Excel's Range.Value2 returns.</summary>
    private static object?[,] Cells(params object?[][] rows)
    {
        var cols = rows.Max(r => r.Length);
        var cells = (object?[,])Array.CreateInstance(typeof(object), [rows.Length, cols], [1, 1]);
        for (var r = 0; r < rows.Length; r++)
            for (var c = 0; c < rows[r].Length; c++)
                cells[r + 1, c + 1] = rows[r][c];
        return cells;
    }

    private static readonly object?[,] SampleFinancials = Cells(
        ["RowId", "Segment", "Country", "Metric", "2023M01", "2023M02", "2023M03", "Notes"],
        ["FIN-001", "Retail", "UK", "Revenue", 100.0, 110.0, 120.0, null],
        ["FIN-002", "Wealth", "India", "Revenue", 50.0, "Feb awaiting restatement", 70.0, "Feb awaiting restatement"],
        ["FIN-003", "Retail", "UK", "Loans", 1000.0, null, 1200.0, null],
        [null, null, null, null, null, null, null, null]);

    private static readonly object?[,] SampleMetrics = Cells(
        ["Metric", "Display Name", "Shown as", "Calculated from"],
        ["Revenue", "Revenue", "USD thousands", null],
        ["Loans", "Loans outstanding", "USD thousands", null],
        ["CostToIncome", "Cost-to-income %", "Percent", "Expenses and Revenue"]);

    [Fact]
    public void Reads_month_columns_and_ignores_notes_column()
    {
        var (periods, _) = FinancialsParser.ParseFinancials(SampleFinancials);

        Assert.Equal([new Period(2023, 1), new Period(2023, 2), new Period(2023, 3)], periods);
    }

    [Fact]
    public void Reads_rows_and_skips_trailing_empty_rows()
    {
        var (_, records) = FinancialsParser.ParseFinancials(SampleFinancials);

        Assert.Equal(3, records.Count);
        var first = records[0];
        Assert.Equal(("FIN-001", "Retail", "UK", "Revenue"), (first.RowId, first.Segment, first.Country, first.Metric));
        Assert.Equal([100m, 110m, 120m], first.Values);
        Assert.Null(first.Note);
    }

    [Fact]
    public void Blank_and_text_cells_become_null_not_zero()
    {
        var (_, records) = FinancialsParser.ParseFinancials(SampleFinancials);

        Assert.Null(records[1].Values[1]); // text in a month cell
        Assert.Null(records[2].Values[1]); // empty cell
        Assert.Equal("Feb awaiting restatement", records[1].Note);
    }

    [Fact]
    public void Month_columns_are_found_by_header_not_position()
    {
        var shuffled = Cells(
            ["2023M02", "Metric", "Segment", "2023M01", "Country"],
            [20.0, "Revenue", "Retail", 10.0, "UK"]);

        var (periods, records) = FinancialsParser.ParseFinancials(shuffled);

        Assert.Equal([new Period(2023, 1), new Period(2023, 2)], periods);
        Assert.Equal([10m, 20m], records[0].Values);
    }

    [Fact]
    public void Missing_required_column_is_reported()
    {
        var noCountry = Cells(["Segment", "Metric", "2023M01"], ["Retail", "Revenue", 1.0]);

        var ex = Assert.Throws<FormatException>(() => FinancialsParser.ParseFinancials(noCountry));
        Assert.Contains("Country", ex.Message);
    }

    [Fact]
    public void Duplicate_rows_are_rejected()
    {
        var duplicate = Cells(
            ["Segment", "Country", "Metric", "2023M01"],
            ["Retail", "UK", "Revenue", 1.0],
            ["Retail", "UK", "Revenue", 2.0]);

        Assert.Throws<FormatException>(() => FinancialsParser.ParseFinancials(duplicate));
    }

    [Fact]
    public void Reads_metrics_including_calculated_ones()
    {
        var metrics = FinancialsParser.ParseMetrics(SampleMetrics);

        Assert.Equal(3, metrics.Count);
        Assert.Equal("Loans outstanding", metrics[1].DisplayName);
        Assert.False(metrics[1].IsPercent);
        Assert.False(metrics[1].IsCalculated);

        var cti = metrics[2];
        Assert.True(cti.IsPercent);
        Assert.True(cti.IsCalculated);
        Assert.Equal(["Expenses", "Revenue"], cti.CalculatedFrom);
    }

    [Fact]
    public void Parse_builds_full_dataset()
    {
        var data = FinancialsParser.Parse(SampleFinancials, SampleMetrics);

        Assert.Equal(new Period(2023, 3), data.LatestPeriod);
        Assert.Equal(["Retail", "Wealth"], data.Segments);
        Assert.Equal(["India", "UK"], data.Countries);
        Assert.Equal("Cost-to-income %", data.GetMetric("costtoincome").DisplayName);
    }
}
