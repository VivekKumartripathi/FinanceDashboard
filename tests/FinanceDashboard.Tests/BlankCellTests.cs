using FinanceDashboard.Core.Calculation;
using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Tests;

/// <summary>
/// Blank cells are never zero: amounts use the cells that have values and are flagged,
/// ratios leave out months that are blank on either side.
/// </summary>
public class BlankCellTests
{
    private static readonly Period[] Months = [new(2025, 4), new(2025, 5), new(2025, 6)];

    private static readonly FinancialData Data = new(
        Months,
        [
            new("R1", "Retail", "UK", "Revenue", [100m, 100m, 100m], null),
            new("R2", "Retail", "UK", "Expenses", [50m, 50m, null], "Jun-25 not closed yet"),
            new("R3", "Retail", "India", "Expenses", [10m, 10m, 10m], null),
            new("R4", "Retail", "UK", "Loans", [500m, null, 700m], null),
            new("R5", "Retail", "India", "Loans", [50m, 60m, null], null),
        ],
        [
            new("Revenue", "Revenue", false, []),
            new("Expenses", "Operating expenses", false, []),
            new("Loans", "Loans outstanding", false, []),
            new("CostToIncome", "Cost-to-income %", true, ["Expenses", "Revenue"]),
        ]);

    private static SeriesPoint Last(string metric, string segment, string country, Frequency frequency) =>
        SeriesCalculator.Calculate(Data, new Selection(metric, segment, country, frequency, 1)).Points.Single();

    [Fact]
    public void A_blank_cell_on_its_own_gives_no_value_rather_than_zero()
    {
        var point = Last("Expenses", "Retail", "UK", Frequency.Monthly);

        Assert.Null(point.Value);
        Assert.False(point.IsComplete);
    }

    [Fact]
    public void Totals_use_the_cells_that_have_values_and_quote_the_sheet_note()
    {
        var point = Last("Expenses", "Retail", Selection.All, Frequency.Monthly);

        Assert.Equal(10m, point.Value);
        Assert.Equal("Retail / UK / Expenses is blank for Jun 25 (Jun-25 not closed yet).", Assert.Single(point.Notes));
    }

    [Fact]
    public void Quarterly_sum_flags_the_blank_month()
    {
        var point = Last("Expenses", "Retail", "UK", Frequency.Quarterly);

        Assert.Equal(100m, point.Value);
        Assert.Contains("Jun 25", Assert.Single(point.Notes));
    }

    [Fact]
    public void Balance_ignores_blanks_before_the_quarter_end()
    {
        var point = Last("Loans", "Retail", "UK", Frequency.Quarterly);

        Assert.Equal(700m, point.Value);
        Assert.True(point.IsComplete);
    }

    [Fact]
    public void Balance_flags_a_blank_at_the_quarter_end()
    {
        var point = Last("Loans", "Retail", Selection.All, Frequency.Quarterly);

        Assert.Equal(700m, point.Value);
        Assert.Contains("Retail / India / Loans is blank for Jun 25", Assert.Single(point.Notes));
    }

    [Fact]
    public void Ratio_leaves_out_months_that_are_blank_on_either_side()
    {
        // Apr + May only: 100 / 200 = 50%. Counting June's blank as zero would give 100 / 300 = 33.3%.
        var point = Last("CostToIncome", "Retail", "UK", Frequency.Quarterly);

        Assert.Equal(50m, point.Value);
        Assert.Contains("left out of the ratio", Assert.Single(point.Notes));
    }

    [Fact]
    public void Ratio_with_every_month_blank_has_no_value()
    {
        var point = Last("CostToIncome", "Retail", "UK", Frequency.Monthly);

        Assert.Null(point.Value);
        Assert.False(point.IsComplete);
    }

    [Fact]
    public void Real_workbook_flags_the_Treasury_UK_June_gap()
    {
        var point = SeriesCalculator.Calculate(SampleWorkbook.Data,
            new Selection("Expenses", "Treasury", "UK", Frequency.Monthly, 1)).Points.Single();

        Assert.Null(point.Value);
        Assert.Contains("Jun-25 not closed yet", Assert.Single(point.Notes));
    }

    [Fact]
    public void Real_workbook_flags_the_Wealth_India_February_gap()
    {
        var points = SeriesCalculator.Calculate(SampleWorkbook.Data,
            new Selection("Revenue", "Wealth", "India", Frequency.Monthly, 30)).Points;

        var feb = points.Single(p => p.Label == "Feb 24");
        Assert.Null(feb.Value);
        Assert.Contains("Feb-24 awaiting restatement", Assert.Single(feb.Notes));
        Assert.Equal(29, points.Count(p => p.IsComplete));
    }
}
