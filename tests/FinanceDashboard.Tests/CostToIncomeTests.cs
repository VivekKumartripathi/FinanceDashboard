using FinanceDashboard.Core.Calculation;
using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Tests;

/// <summary>
/// Cost-to-income % rules on a small dataset with round numbers.
/// </summary>
public class CostToIncomeTests
{
    private static readonly Period[] Months = [new(2025, 1), new(2025, 2), new(2025, 3)];

    private static readonly FinancialData Data = new(
        Months,
        [
            // Retail UK: 50% every month.  Retail India: 25% every month.
            new("R1", "Retail", "UK", "Revenue", [100m, 100m, 100m], null),
            new("R2", "Retail", "UK", "Expenses", [50m, 50m, 50m], null),
            new("R3", "Retail", "India", "Revenue", [400m, 400m, 400m], null),
            new("R4", "Retail", "India", "Expenses", [100m, 100m, 100m], null),
            // Wealth UK: 80% every month.
            new("R5", "Wealth", "UK", "Revenue", [10m, 10m, 10m], null),
            new("R6", "Wealth", "UK", "Expenses", [8m, 8m, 8m], null),
            // Corporate UK: revenue only, so it has no ratio.
            new("R7", "Corporate", "UK", "Revenue", [1000m, 1000m, 1000m], null),
        ],
        [
            new("Revenue", "Revenue", false, []),
            new("Expenses", "Operating expenses", false, []),
            new("CostToIncome", "Cost-to-income %", true, ["Expenses", "Revenue"]),
        ]);

    private static decimal Value(string segment, string country, Frequency frequency = Frequency.Monthly) =>
        SeriesCalculator.Calculate(Data, new Selection("CostToIncome", segment, country, frequency, 1))
            .Points.Single().Value!.Value;

    [Fact]
    public void One_segment_one_country_is_expenses_over_revenue()
    {
        Assert.Equal(50m, Value("Retail", "UK"));
    }

    [Fact]
    public void Quarterly_uses_the_quarter_totals()
    {
        Assert.Equal(50m, Value("Retail", "UK", Frequency.Quarterly));
    }

    [Fact]
    public void All_countries_adds_countries_within_the_segment_first()
    {
        // (50 + 100) / (100 + 400) = 30%, not the average of 50% and 25%.
        Assert.Equal(30m, Value("Retail", Selection.All));
    }

    [Fact]
    public void All_segments_averages_each_segment_ratio()
    {
        // Retail UK 50% and Wealth UK 80% average to 65%; combined totals would give 58 / 110 = 52.7%.
        // Corporate has no expenses, so it has no ratio and is left out.
        Assert.Equal(65m, Value(Selection.All, "UK"));
    }

    [Fact]
    public void Segment_without_both_metrics_is_reported()
    {
        Assert.Throws<ArgumentException>(() => Value("Corporate", "UK"));
    }
}
