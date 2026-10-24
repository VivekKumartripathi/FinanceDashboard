using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Core.Calculation;

/// <summary>Optionally split the chart into one series per segment or per country.</summary>
public enum Breakdown
{
    None,
    Segment,
    Country
}

/// <summary>
/// Everything the dashboard shows for one set of choices.
/// </summary>
/// <param name="Series">The total series; with a breakdown it covers every segment (or country) being split.</param>
/// <param name="Parts">One series per segment or country when broken down; empty otherwise.</param>
public sealed record DashboardView(
    SeriesResult Series,
    KpiSummary Kpis,
    Breakdown BreakdownBy,
    IReadOnlyList<SeriesResult> Parts);

public static class DashboardCalculator
{
    public static DashboardView Calculate(FinancialData data, Selection selection, Breakdown breakdownBy = Breakdown.None)
    {
        // Breaking down by segment means comparing every segment, so the total is "All" segments (likewise for country).
        var total = breakdownBy switch
        {
            Breakdown.Segment => selection with { Segment = Selection.All },
            Breakdown.Country => selection with { Country = Selection.All },
            _ => selection
        };

        var parts = breakdownBy switch
        {
            Breakdown.Segment => data.Segments.Select(s => total with { Segment = s }),
            Breakdown.Country => data.Countries.Select(c => total with { Country = c }),
            _ => []
        };

        return new DashboardView(
            SeriesCalculator.Calculate(data, total),
            KpiCalculator.Calculate(data, total),
            breakdownBy,
            parts.Select(p => TryCalculate(data, p)).OfType<SeriesResult>().ToList());
    }

    /// <summary>A part with no data (e.g. a country a segment does not operate in) is left out rather than failing.</summary>
    private static SeriesResult? TryCalculate(FinancialData data, Selection selection)
    {
        try
        {
            return SeriesCalculator.Calculate(data, selection);
        }
        catch (ArgumentException)
        {
            return null;
        }
    }
}
