using FinanceDashboard.Core.Calculation;
using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Tests;

/// <summary>
/// Calculator rules on a small hand-made dataset, so each expected number is easy to check by eye.
/// </summary>
public class SeriesCalculatorTests
{
    private static readonly Period[] Months =
        [new(2024, 10), new(2024, 11), new(2024, 12), new(2025, 1), new(2025, 2), new(2025, 3)];

    private static readonly FinancialData Data = new(
        Months,
        [
            new("R1", "Retail", "UK", "Revenue", [10m, 11m, 12m, 13m, 14m, 15m], null),
            new("R2", "Retail", "India", "Revenue", [1m, 2m, 3m, 4m, 5m, 6m], null),
            new("R3", "Wealth", "UK", "Revenue", [100m, 100m, 100m, 100m, null, 100m], "Feb blank"),
            new("R4", "Retail", "UK", "Loans", [500m, 510m, 520m, 530m, 540m, 550m], null),
        ],
        [new("Revenue", "Revenue", false, []), new("Loans", "Loans outstanding", false, [])]);

    private static SeriesResult Monthly(string metric, string segment, string country, int periods) =>
        SeriesCalculator.Calculate(Data, new Selection(metric, segment, country, Frequency.Monthly, periods));

    [Fact]
    public void Single_segment_and_country_uses_only_that_row()
    {
        var result = Monthly("Revenue", "Retail", "UK", 3);

        Assert.Equal(["Jan 25", "Feb 25", "Mar 25"], result.Points.Select(p => p.Label));
        Assert.Equal([13m, 14m, 15m], result.Points.Select(p => p.Value));
    }

    [Fact]
    public void All_countries_adds_every_country_for_the_segment()
    {
        var result = Monthly("Revenue", "Retail", Selection.All, 2);

        Assert.Equal([19m, 21m], result.Points.Select(p => p.Value));
    }

    [Fact]
    public void All_segments_and_all_countries_adds_every_row_of_the_metric()
    {
        var result = Monthly("Revenue", Selection.All, Selection.All, 1);

        Assert.Equal([121m], result.Points.Select(p => p.Value));
    }

    [Fact]
    public void Other_metrics_are_never_included()
    {
        var result = Monthly("Loans", Selection.All, "UK", 1);

        Assert.Equal([550m], result.Points.Select(p => p.Value));
    }

    [Fact]
    public void Series_always_ends_at_the_latest_month()
    {
        var result = Monthly("Revenue", "Retail", "UK", 1);

        Assert.Equal("Mar 25", Assert.Single(result.Points).Label);
    }

    [Fact]
    public void Asking_for_more_periods_than_exist_returns_all_of_them()
    {
        var result = Monthly("Revenue", "Retail", "UK", 99);

        Assert.Equal(6, result.Points.Count);
        Assert.Equal("Oct 24", result.Points[0].Label);
    }

    [Fact]
    public void Matching_ignores_case()
    {
        var result = Monthly("revenue", "retail", "uk", 1);

        Assert.Equal([15m], result.Points.Select(p => p.Value));
    }

    [Fact]
    public void Result_carries_the_metric_definition()
    {
        var result = Monthly("Loans", "Retail", "UK", 1);

        Assert.Equal("Loans outstanding", result.Metric.DisplayName);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-3)]
    public void Period_count_below_one_is_rejected(int periods)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => Monthly("Revenue", "Retail", "UK", periods));
    }

    [Fact]
    public void Unknown_segment_or_metric_is_reported()
    {
        Assert.Throws<ArgumentException>(() => Monthly("Revenue", "Corporate", "UK", 1));
        Assert.Throws<ArgumentException>(() => Monthly("Profit", "Retail", "UK", 1));
    }
}
