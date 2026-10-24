using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Core.Calculation;

/// <summary>
/// The latest value compared with one other period.
/// </summary>
/// <param name="Change">Percent change for amounts; percentage-point difference for percent metrics.</param>
public sealed record Comparison(string Label, decimal? Value, decimal? Change);

/// <summary>
/// Headline numbers for the current selection: where the metric is now and how it moved.
/// </summary>
public sealed record KpiSummary(
    string LatestLabel,
    decimal? Latest,
    Comparison? PreviousPeriod,
    Comparison? SamePeriodLastYear,
    bool IsPercent,
    bool HigherIsBetter,
    bool HasMissingData);

public static class KpiCalculator
{
    /// <summary>For these metrics a fall is good news, so the change is coloured the other way round.</summary>
    private static readonly HashSet<string> LowerIsBetter =
        new(StringComparer.OrdinalIgnoreCase) { "Expenses", "CostToIncome" };

    public static KpiSummary Calculate(FinancialData data, Selection selection)
    {
        // Comparisons can reach further back than the periods on the chart, so use the full history.
        var points = SeriesCalculator.Calculate(data, selection with { PeriodCount = int.MaxValue }).Points;
        var metric = data.GetMetric(selection.Metric);
        var latest = points[points.Count - 1];
        var periodsPerYear = selection.Frequency == Frequency.Monthly ? 12 : 4;

        Comparison? Compare(int periodsBack, out bool missing)
        {
            missing = false;
            var index = points.Count - 1 - periodsBack;
            if (index < 0) return null;

            var other = points[index];
            missing = !other.IsComplete;
            return new Comparison(other.Label, other.Value, Change(latest.Value, other.Value, metric.IsPercent));
        }

        var previous = Compare(1, out var previousMissing);
        var lastYear = Compare(periodsPerYear, out var lastYearMissing);

        return new KpiSummary(
            latest.Label,
            latest.Value,
            previous,
            lastYear,
            metric.IsPercent,
            HigherIsBetter: !LowerIsBetter.Contains(metric.Code),
            HasMissingData: !latest.IsComplete || previousMissing || lastYearMissing);
    }

    /// <summary>
    /// Amounts: percent change. Percent metrics: difference in percentage points
    /// (37.4% after 36.5% is +0.9 pp, not +2.5%).
    /// </summary>
    internal static decimal? Change(decimal? current, decimal? earlier, bool isPercent)
    {
        if (current is not decimal now || earlier is not decimal then) return null;
        if (isPercent) return now - then;
        return then == 0m ? null : (now - then) / Math.Abs(then) * 100m;
    }
}
