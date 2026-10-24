using FinanceDashboard.Core.Calculation;
using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Tests;

public class CommentaryWriterTests
{
    // Jan 24 to Mar 25 (15 months).
    private static readonly Period[] Months =
        Enumerable.Range(0, 15).Select(i => new Period(2024, 1).AddMonths(i)).ToArray();

    private static List<decimal?> Flat(decimal value) => Months.Select(_ => (decimal?)value).ToList();

    private static readonly FinancialData Data = new(
        Months,
        [
            // Retail UK revenue grows 100 a month from 1,000; Wealth UK is flat at 500, then drops 300 in Mar 25.
            new("R1", "Retail", "UK", "Revenue", Months.Select((_, i) => (decimal?)(1000 + 100 * i)).ToList(), null),
            new("R2", "Wealth", "UK", "Revenue", Months.Select((_, i) => (decimal?)(i == 14 ? 200 : 500)).ToList(), null),
            new("R3", "Retail", "UK", "Expenses", Flat(400), null),
            new("R4", "Wealth", "UK", "Expenses", Months.Select((_, i) => i == 14 ? null : (decimal?)100).ToList(), "Mar-25 not closed yet"),
        ],
        [
            new("Revenue", "Revenue", false, []),
            new("Expenses", "Operating expenses", false, []),
            new("CostToIncome", "Cost-to-income %", true, ["Expenses", "Revenue"]),
        ]);

    private static string Comment(string metric, string segment, string country, int periods = 6, Breakdown by = Breakdown.None)
    {
        var view = DashboardCalculator.Calculate(Data, new Selection(metric, segment, country, Frequency.Monthly, periods), by);
        return CommentaryWriter.Write(Data, view);
    }

    [Fact]
    public void Headline_gives_value_and_both_comparisons()
    {
        var text = Comment("Revenue", "Retail", "UK");

        // Mar 25 = 2,400; Feb 25 = 2,300 (+4.3%); Mar 24 = 1,200 (+100%).
        Assert.StartsWith("Revenue for Retail in UK was USD 2.4m in Mar 25, up 4.3% on Feb 25 and up 100.0% on Mar 24.", text);
    }

    [Fact]
    public void Names_the_segment_that_moved_most()
    {
        // Retail +100, Wealth -300: Wealth moved most.
        Assert.Contains("Wealth moved most (-USD 300k).", Comment("Revenue", Selection.All, "UK"));
    }

    [Fact]
    public void Uses_the_dashboard_comparison_when_there_is_one()
    {
        Assert.Contains("Wealth moved most", Comment("Revenue", Selection.All, "UK", by: Breakdown.Segment));
    }

    [Fact]
    public void No_mover_for_a_single_segment_and_country()
    {
        Assert.DoesNotContain("moved most", Comment("Revenue", "Retail", "UK"));
    }

    [Fact]
    public void Says_when_the_latest_period_is_the_highest_shown()
    {
        Assert.Contains("This is the highest of the 6 months shown.", Comment("Revenue", "Retail", "UK"));
    }

    [Fact]
    public void Says_when_the_latest_period_is_the_lowest_shown()
    {
        Assert.Contains("This is the lowest of the 6 months shown.", Comment("Revenue", "Wealth", "UK"));
    }

    [Fact]
    public void Percent_metrics_use_points_and_say_whether_it_is_better_or_worse()
    {
        var text = Comment("CostToIncome", "Retail", "UK");

        // Mar 25: 400 / 2,400 = 16.7%. Feb 25: 400 / 2,300 = 17.4%. Down 0.7 pp, which is good news.
        Assert.Contains("was 16.7% in Mar 25, down 0.7 pp (better) on Feb 25", text);
    }

    [Fact]
    public void Missing_latest_value_is_explained_instead_of_guessed()
    {
        var text = Comment("Expenses", "Wealth", "UK");

        Assert.StartsWith("There is no Operating expenses figure for Wealth in UK for Mar 25 yet.", text);
        Assert.Contains("Mar-25 not closed yet", text);
    }

    [Fact]
    public void Incomplete_latest_value_is_flagged()
    {
        Assert.Contains("Mar 25 is incomplete", Comment("Expenses", Selection.All, "UK"));
    }

    [Theory]
    [InlineData(10265, "USD 10.3m")]
    [InlineData(453406, "USD 453.4m")]
    [InlineData(850, "USD 850k")]
    public void Money_is_shown_in_thousands_or_millions(double thousands, string expected)
    {
        Assert.Equal(expected, CommentaryWriter.Money((decimal)thousands));
    }

    [Theory]
    [InlineData(2.74, false, true, "up 2.7%")]
    [InlineData(-2.74, false, true, "down 2.7%")]
    [InlineData(3, false, false, "up 3.0% (worse)")]
    [InlineData(-0.9, true, false, "down 0.9 pp (better)")]
    [InlineData(0.01, false, true, "unchanged")]
    public void Movement_wording(double change, bool isPercent, bool higherIsBetter, string expected)
    {
        Assert.Equal(expected, CommentaryWriter.Movement((decimal)change, isPercent, higherIsBetter));
    }

    [Fact]
    public void Real_workbook_E2_selection()
    {
        var view = DashboardCalculator.Calculate(SampleWorkbook.Data,
            new Selection("Revenue", Selection.All, "India", Frequency.Monthly, 6));

        var text = CommentaryWriter.Write(SampleWorkbook.Data, view);

        Assert.StartsWith("Revenue across all segments in India was USD 10.3m in Jun 25, down 2.7% on May 25", text);
    }
}
