using FinanceDashboard.Core.Calculation;
using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Tests;

/// <summary>
/// Breakdowns, checked against the real workbook: the parts must add back up to the total.
/// </summary>
public class DashboardCalculatorTests
{
    private static DashboardView View(string metric, string segment, string country, Frequency frequency, Breakdown by) =>
        DashboardCalculator.Calculate(SampleWorkbook.Data, new Selection(metric, segment, country, frequency, 30), by);

    [Fact]
    public void No_breakdown_has_no_parts()
    {
        var view = View("Revenue", "Retail", "UK", Frequency.Monthly, Breakdown.None);

        Assert.Empty(view.Parts);
        Assert.Equal("Retail", view.Series.Selection.Segment);
    }

    [Fact]
    public void Breakdown_by_segment_has_one_part_per_segment_and_total_covers_all_segments()
    {
        var view = View("Revenue", "Retail", "India", Frequency.Monthly, Breakdown.Segment);

        Assert.Equal(["Corporate", "Retail", "Treasury", "Wealth"], view.Parts.Select(p => p.Selection.Segment));
        Assert.All(view.Parts, p => Assert.Equal("India", p.Selection.Country));
        Assert.Equal(Selection.All, view.Series.Selection.Segment);
        Assert.Equal(view.Series.Points[^1].Value, view.Kpis.Latest);   // headline numbers describe the total
    }

    [Theory]
    [InlineData("Revenue", Frequency.Monthly, Breakdown.Segment)]
    [InlineData("Expenses", Frequency.Quarterly, Breakdown.Country)]
    [InlineData("Loans", Frequency.Quarterly, Breakdown.Segment)]
    [InlineData("Loans", Frequency.Monthly, Breakdown.Country)]
    public void Amount_parts_add_up_to_the_total(string metric, Frequency frequency, Breakdown by)
    {
        var view = View(metric, Selection.All, Selection.All, frequency, by);

        for (var i = 0; i < view.Series.Points.Count; i++)
        {
            var sumOfParts = view.Parts.Sum(p => p.Points[i].Value ?? 0m);
            Assert.Equal(view.Series.Points[i].Value, sumOfParts);
        }
    }

    [Fact]
    public void Cost_to_income_by_segment_keeps_E4_as_the_total()
    {
        var view = DashboardCalculator.Calculate(SampleWorkbook.Data,
            new Selection("CostToIncome", Selection.All, "India", Frequency.Quarterly, 4), Breakdown.Segment);

        Assert.Equal([57.0m, 56.4m, 57.5m, 56.3m], view.Series.Points.Select(p => Math.Round(p.Value!.Value, 1)));
        Assert.Equal(4, view.Parts.Count);
        // E3 is one of the parts.
        var treasury = view.Parts.Single(p => p.Selection.Segment == "Treasury");
        Assert.Equal([39.0m, 31.2m, 36.5m, 37.4m], treasury.Points.Select(p => Math.Round(p.Value!.Value, 1)));
    }

    [Fact]
    public void Parts_without_data_are_left_out()
    {
        var data = new FinancialData(
            [new Period(2025, 1)],
            [
                new("R1", "Retail", "UK", "Revenue", [1m], null),
                new("R2", "Wealth", "India", "Revenue", [2m], null),
            ],
            [new("Revenue", "Revenue", false, [])]);

        var view = DashboardCalculator.Calculate(data, new Selection("Revenue", Selection.All, "UK", Frequency.Monthly, 1), Breakdown.Segment);

        Assert.Equal("Retail", Assert.Single(view.Parts).Selection.Segment);
    }
}
