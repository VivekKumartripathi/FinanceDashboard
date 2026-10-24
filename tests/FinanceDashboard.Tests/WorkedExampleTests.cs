using FinanceDashboard.Core.Calculation;
using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Tests;

/// <summary>
/// The four worked examples from the brief, run against the real Financials.xlsx.
/// Amounts are compared to the nearest USD thousand and percents to one decimal place,
/// which is the precision the brief shows.
/// </summary>
public class WorkedExampleTests
{
    private static SeriesResult Run(string metric, string segment, string country, Frequency frequency, int periods) =>
        SeriesCalculator.Calculate(SampleWorkbook.Data, new Selection(metric, segment, country, frequency, periods));

    private static void AssertSeries(SeriesResult result, string[] labels, decimal[] values, int decimals)
    {
        Assert.Equal(labels, result.Points.Select(p => p.Label));
        Assert.Equal(values, result.Points.Select(p => Math.Round(p.Value!.Value, decimals, MidpointRounding.AwayFromZero)));
    }

    [Fact]
    public void E1_Loans_Retail_AllCountries_Quarterly_4()
    {
        var result = Run("Loans", "Retail", Selection.All, Frequency.Quarterly, 4);

        AssertSeries(result,
            ["Q3 24", "Q4 24", "Q1 25", "Q2 25"],
            [424_204m, 437_325m, 448_003m, 453_406m],
            decimals: 0);
    }

    [Fact]
    public void E2_Revenue_AllSegments_India_Monthly_6()
    {
        var result = Run("Revenue", Selection.All, "India", Frequency.Monthly, 6);

        AssertSeries(result,
            ["Jan 25", "Feb 25", "Mar 25", "Apr 25", "May 25", "Jun 25"],
            [10_420m, 9_833m, 9_338m, 10_295m, 10_549m, 10_265m],
            decimals: 0);
    }

    [Fact]
    public void E3_CostToIncome_Treasury_India_Quarterly_4()
    {
        var result = Run("CostToIncome", "Treasury", "India", Frequency.Quarterly, 4);

        AssertSeries(result,
            ["Q3 24", "Q4 24", "Q1 25", "Q2 25"],
            [39.0m, 31.2m, 36.5m, 37.4m],
            decimals: 1);
    }

    [Fact]
    public void E4_CostToIncome_AllSegments_India_Quarterly_4()
    {
        var result = Run("CostToIncome", Selection.All, "India", Frequency.Quarterly, 4);

        AssertSeries(result,
            ["Q3 24", "Q4 24", "Q1 25", "Q2 25"],
            [57.0m, 56.4m, 57.5m, 56.3m],
            decimals: 1);
    }

    [Fact]
    public void Sample_workbook_loads_as_expected()
    {
        var data = SampleWorkbook.Data;

        Assert.Equal(30, data.Periods.Count);
        Assert.Equal(new Period(2023, 1), data.Periods[0]);
        Assert.Equal(new Period(2025, 6), data.LatestPeriod);
        Assert.Equal(36, data.Records.Count);
        Assert.Equal(["Corporate", "Retail", "Treasury", "Wealth"], data.Segments);
        Assert.Equal(["India", "UK", "USA"], data.Countries);
        Assert.Equal(4, data.Metrics.Count);
    }
}
