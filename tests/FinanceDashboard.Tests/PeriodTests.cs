using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Tests;

public class PeriodTests
{
    [Theory]
    [InlineData("2023M01", 2023, 1)]
    [InlineData("2025M06", 2025, 6)]
    [InlineData(" 2024M12 ", 2024, 12)]
    public void TryParse_accepts_month_headers(string text, int year, int month)
    {
        Assert.True(Period.TryParse(text, out var period));
        Assert.Equal(new Period(year, month), period);
    }

    [Theory]
    [InlineData("Notes")]
    [InlineData("RowId")]
    [InlineData("2023M13")]
    [InlineData("2023-01")]
    [InlineData("")]
    [InlineData(null)]
    public void TryParse_rejects_other_headers(string? text)
    {
        Assert.False(Period.TryParse(text, out _));
    }

    [Theory]
    [InlineData(1, 1, false)]
    [InlineData(3, 1, true)]
    [InlineData(4, 2, false)]
    [InlineData(9, 3, true)]
    [InlineData(12, 4, true)]
    public void Quarter_and_quarter_end(int month, int quarter, bool isQuarterEnd)
    {
        var period = new Period(2024, month);
        Assert.Equal(quarter, period.Quarter);
        Assert.Equal(isQuarterEnd, period.IsQuarterEnd);
    }

    [Fact]
    public void Labels_match_the_brief()
    {
        Assert.Equal("Jan 25", new Period(2025, 1).MonthLabel);
        Assert.Equal("Q3 24", new Period(2024, 9).QuarterLabel);
    }

    [Fact]
    public void AddMonths_crosses_year_boundaries()
    {
        Assert.Equal(new Period(2025, 1), new Period(2024, 12).AddMonths(1));
        Assert.Equal(new Period(2024, 10), new Period(2025, 6).AddMonths(-8));
    }

    [Fact]
    public void Periods_sort_chronologically()
    {
        Assert.True(new Period(2024, 12).CompareTo(new Period(2025, 1)) < 0);
    }
}
