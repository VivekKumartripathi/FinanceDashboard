using FinanceDashboard.Core.Calculation;
using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Tests;

/// <summary>
/// Quarterly roll-up rules on a small dataset: Sep 24 to Apr 25, so it starts and ends mid-quarter.
/// </summary>
public class QuarterlyTests
{
    private static readonly Period[] Months =
    [
        new(2024, 9),
        new(2024, 10), new(2024, 11), new(2024, 12),
        new(2025, 1), new(2025, 2), new(2025, 3),
        new(2025, 4)
    ];

    private static readonly FinancialData Data = new(
        Months,
        [
            new("R1", "Retail", "UK", "Revenue", [9m, 1m, 2m, 3m, 4m, 5m, 6m, 7m], null),
            new("R2", "Retail", "UK", "Loans", [90m, 100m, 110m, 120m, 130m, 140m, 150m, 160m], null),
        ],
        [new("Revenue", "Revenue", false, []), new("Loans", "Loans outstanding", false, [])]);

    private static SeriesResult Quarterly(string metric, int periods) =>
        SeriesCalculator.Calculate(Data, new Selection(metric, "Retail", "UK", Frequency.Quarterly, periods));

    [Fact]
    public void Flows_sum_the_three_months_of_each_quarter()
    {
        var result = Quarterly("Revenue", 3);

        Assert.Equal(["Q4 24", "Q1 25", "Q2 25"], result.Points.Select(p => p.Label));
        Assert.Equal([6m, 15m, 7m], result.Points.Select(p => p.Value));
    }

    [Fact]
    public void Balances_take_the_last_month_of_each_quarter()
    {
        var result = Quarterly("Loans", 3);

        Assert.Equal([120m, 150m, 160m], result.Points.Select(p => p.Value));
    }

    [Fact]
    public void Series_ends_with_the_quarter_holding_the_latest_month()
    {
        var result = Quarterly("Revenue", 1);

        Assert.Equal("Q2 25", Assert.Single(result.Points).Label);
    }

    [Fact]
    public void Partial_quarters_are_kept_but_flagged()
    {
        var result = Quarterly("Revenue", 99);

        Assert.Equal(["Q3 24", "Q4 24", "Q1 25", "Q2 25"], result.Points.Select(p => p.Label));
        Assert.Contains("1 of 3 months", Assert.Single(result.Points[0].Notes));
        Assert.True(result.Points[1].IsComplete);
        Assert.True(result.Points[2].IsComplete);
        Assert.Contains("1 of 3 months", Assert.Single(result.Points[3].Notes));
    }

    [Theory]
    [InlineData("Loans", Aggregation.PeriodEnd)]
    [InlineData("loans", Aggregation.PeriodEnd)]
    [InlineData("Revenue", Aggregation.Sum)]
    [InlineData("Expenses", Aggregation.Sum)]
    public void Aggregation_rule_per_metric(string metric, Aggregation expected)
    {
        Assert.Equal(expected, AggregationRules.For(metric));
    }
}
