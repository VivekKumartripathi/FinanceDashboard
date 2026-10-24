using FinanceDashboard.Core.Calculation;
using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Tests;

public class PeriodLabelTests
{
    [Theory]
    [InlineData("Jun 25", Frequency.Monthly, "Jun 25")]
    [InlineData("jun 2025", Frequency.Monthly, "Jun 25")]
    [InlineData("June 2025", Frequency.Monthly, "Jun 25")]
    [InlineData("2025M06", Frequency.Monthly, "Jun 25")]
    [InlineData("Q2 25", Frequency.Quarterly, "Q2 25")]
    [InlineData("q3 2024", Frequency.Quarterly, "Q3 24")]
    [InlineData("2024Q4", Frequency.Quarterly, "Q4 24")]
    [InlineData(" Q1 25 ", Frequency.Quarterly, "Q1 25")]
    public void Accepts_common_ways_of_writing_a_period(string text, Frequency frequency, string label)
    {
        Assert.True(PeriodLabel.TryParse(text, out var f, out var l));
        Assert.Equal(frequency, f);
        Assert.Equal(label, l);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("Q5 25")]
    [InlineData("Foo 25")]
    [InlineData("2025-06")]
    public void Rejects_anything_else(string? text)
    {
        Assert.False(PeriodLabel.TryParse(text, out _, out _));
    }

    [Fact]
    public void Metrics_can_be_found_by_display_name()
    {
        Assert.Equal("CostToIncome", SampleWorkbook.Data.GetMetric("Cost-to-income %").Code);
        Assert.Equal("Loans", SampleWorkbook.Data.GetMetric("loans outstanding").Code);
    }

    [Fact]
    public void A_formula_style_lookup_reproduces_E1()
    {
        // What =FIN.VALUE("Loans outstanding", "Retail", "All", "Q2 25") does.
        Assert.True(PeriodLabel.TryParse("Q2 25", out var frequency, out var label));
        var points = SeriesCalculator.Calculate(SampleWorkbook.Data,
            new Selection("Loans outstanding", "Retail", Selection.All, frequency, int.MaxValue)).Points;

        Assert.Equal(453_406m, points.Single(p => p.Label == label).Value);
    }
}
