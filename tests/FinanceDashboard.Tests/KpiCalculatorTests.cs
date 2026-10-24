using FinanceDashboard.Core.Calculation;
using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Tests;

public class KpiCalculatorTests
{
    // Jan 24 to Mar 25: 15 months, so a monthly year-on-year comparison exists.
    private static readonly Period[] Months =
        Enumerable.Range(0, 15).Select(i => new Period(2024, 1).AddMonths(i)).ToArray();

    private static readonly FinancialData Data = new(
        Months,
        [
            // Revenue: 100, 110, ... 240.  Expenses: 50 every month, except a blank Feb 25.
            new("R1", "Retail", "UK", "Revenue", Months.Select((_, i) => (decimal?)(100 + 10 * i)).ToList(), null),
            new("R2", "Retail", "UK", "Expenses", Months.Select((_, i) => i == 13 ? null : (decimal?)50).ToList(), "Feb-25 late"),
        ],
        [
            new("Revenue", "Revenue", false, []),
            new("Expenses", "Operating expenses", false, []),
            new("CostToIncome", "Cost-to-income %", true, ["Expenses", "Revenue"]),
        ]);

    private static KpiSummary Kpis(string metric, Frequency frequency = Frequency.Monthly) =>
        KpiCalculator.Calculate(Data, new Selection(metric, "Retail", "UK", frequency, 3));

    [Fact]
    public void Latest_value_is_the_last_period()
    {
        var kpis = Kpis("Revenue");

        Assert.Equal("Mar 25", kpis.LatestLabel);
        Assert.Equal(240m, kpis.Latest);
    }

    [Fact]
    public void Compares_with_previous_period_and_same_period_last_year()
    {
        var kpis = Kpis("Revenue");

        Assert.Equal("Feb 25", kpis.PreviousPeriod!.Label);
        Assert.Equal(230m, kpis.PreviousPeriod.Value);
        Assert.Equal(Math.Round(10m / 230m * 100m, 6), Math.Round(kpis.PreviousPeriod.Change!.Value, 6));

        Assert.Equal("Mar 24", kpis.SamePeriodLastYear!.Label);
        Assert.Equal(120m, kpis.SamePeriodLastYear.Value);
        Assert.Equal(100m, kpis.SamePeriodLastYear.Change);
    }

    [Fact]
    public void Comparisons_reach_beyond_the_periods_shown_on_the_chart()
    {
        // The selection shows 3 months, but last year is 12 months back.
        Assert.NotNull(Kpis("Revenue").SamePeriodLastYear);
    }

    [Fact]
    public void Quarterly_last_year_is_four_quarters_back()
    {
        var kpis = Kpis("Revenue", Frequency.Quarterly);

        Assert.Equal("Q1 25", kpis.LatestLabel);
        Assert.Equal("Q4 24", kpis.PreviousPeriod!.Label);
        Assert.Equal("Q1 24", kpis.SamePeriodLastYear!.Label);
    }

    [Fact]
    public void No_year_on_year_when_history_is_too_short()
    {
        var shortData = new FinancialData(Months.Take(5).ToList(),
            [new("R1", "Retail", "UK", "Revenue", [1m, 2m, 3m, 4m, 5m], null)],
            [new("Revenue", "Revenue", false, [])]);

        var kpis = KpiCalculator.Calculate(shortData, new Selection("Revenue", "Retail", "UK", Frequency.Monthly, 3));

        Assert.NotNull(kpis.PreviousPeriod);
        Assert.Null(kpis.SamePeriodLastYear);
    }

    [Fact]
    public void Percent_metrics_change_in_percentage_points()
    {
        var kpis = Kpis("CostToIncome");

        // Mar 25: 50 / 240 = 20.83%.  Mar 24: 50 / 120 = 41.67%.  Difference is -20.83 pp, not -50%.
        Assert.Equal(-20.83m, Math.Round(kpis.SamePeriodLastYear!.Change!.Value, 2));
        Assert.True(kpis.IsPercent);
    }

    [Theory]
    [InlineData("Revenue", true)]
    [InlineData("Expenses", false)]
    [InlineData("CostToIncome", false)]
    public void Knows_whether_a_rise_is_good_news(string metric, bool higherIsBetter)
    {
        Assert.Equal(higherIsBetter, Kpis(metric).HigherIsBetter);
    }

    [Fact]
    public void Flags_missing_data_in_any_compared_period()
    {
        Assert.True(Kpis("Expenses").HasMissingData);   // Feb 25 is the previous period
        Assert.False(Kpis("Revenue").HasMissingData);
    }

    [Theory]
    [InlineData(110, 100, false, 10)]
    [InlineData(90, 100, false, -10)]
    [InlineData(37.4, 36.5, true, 0.9)]
    public void Change_rules(double current, double earlier, bool isPercent, double expected)
    {
        var change = KpiCalculator.Change((decimal)current, (decimal)earlier, isPercent);

        Assert.Equal((decimal)expected, Math.Round(change!.Value, 6));
    }

    [Fact]
    public void Change_is_unknown_without_both_values_or_from_zero()
    {
        Assert.Null(KpiCalculator.Change(null, 1m, false));
        Assert.Null(KpiCalculator.Change(1m, null, false));
        Assert.Null(KpiCalculator.Change(1m, 0m, false));
    }
}
