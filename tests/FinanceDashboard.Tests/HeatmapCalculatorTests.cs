using FinanceDashboard.Core.Calculation;
using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Tests;

public class HeatmapCalculatorTests
{
    private static Heatmap Map(string metric, Frequency frequency) =>
        HeatmapCalculator.Calculate(SampleWorkbook.Data, metric, frequency);

    [Fact]
    public void Has_every_segment_and_country_plus_totals()
    {
        var map = Map("Revenue", Frequency.Monthly);

        Assert.Equal("Jun 25", map.PeriodLabel);
        Assert.Equal(["Corporate", "Retail", "Treasury", "Wealth"], map.Segments);
        Assert.Equal(["India", "UK", "USA"], map.Countries);
        Assert.Equal(5 * 4, map.Cells.Count);
    }

    [Fact]
    public void Amount_totals_add_up_across_rows_and_columns()
    {
        var map = Map("Revenue", Frequency.Quarterly);

        foreach (var segment in map.Segments)
            Assert.Equal(map.Cell(segment, Selection.All).Value, map.Countries.Sum(c => map.Cell(segment, c).Value));

        foreach (var country in map.Countries)
            Assert.Equal(map.Cell(Selection.All, country).Value, map.Segments.Sum(s => map.Cell(s, country).Value));
    }

    [Fact]
    public void Cost_to_income_cells_follow_E3_and_E4()
    {
        var map = Map("CostToIncome", Frequency.Quarterly);

        Assert.Equal("Q2 25", map.PeriodLabel);
        Assert.Equal(37.4m, Math.Round(map.Cell("Treasury", "India").Value!.Value, 1));      // E3, Q2 25
        Assert.Equal(56.3m, Math.Round(map.Cell(Selection.All, "India").Value!.Value, 1));   // E4, Q2 25
    }

    [Fact]
    public void Missing_data_is_flagged_in_its_cell()
    {
        var map = Map("Expenses", Frequency.Monthly);

        var cell = map.Cell("Treasury", "UK");
        Assert.Null(cell.Value);
        Assert.False(cell.IsComplete);
        Assert.True(map.Cell("Retail", "UK").IsComplete);
    }

    [Fact]
    public void Combinations_without_data_are_empty_not_errors()
    {
        var data = new FinancialData(
            [new Period(2025, 1)],
            [
                new("R1", "Retail", "UK", "Revenue", [1m], null),
                new("R2", "Wealth", "India", "Revenue", [2m], null),
            ],
            [new("Revenue", "Revenue", false, [])]);

        var map = HeatmapCalculator.Calculate(data, "Revenue", Frequency.Monthly);

        Assert.Null(map.Cell("Retail", "India").Value);
        Assert.True(map.Cell("Retail", "India").IsComplete);
        Assert.Equal(3m, map.Cell(Selection.All, Selection.All).Value);
    }

    [Fact]
    public void Dashboard_view_includes_the_heatmap_for_the_chosen_metric()
    {
        var view = DashboardCalculator.Calculate(SampleWorkbook.Data,
            new Selection("Loans", "Retail", "UK", Frequency.Quarterly, 4));

        Assert.NotNull(view.Heatmap);
        Assert.Equal(453_406m, view.Heatmap!.Cell("Retail", Selection.All).Value);   // E1, Q2 25
    }
}
