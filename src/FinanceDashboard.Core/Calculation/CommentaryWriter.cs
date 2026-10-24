using System.Globalization;
using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Core.Calculation;

/// <summary>
/// Writes a short plain-English summary of the dashboard, the kind of sentence a finance team puts in a monthly pack:
/// "Revenue across all segments in India was USD 10.3m in Jun 25, down 2.7% on May 25 and up 4.1% on Jun 24.
/// Retail moved most (-USD 0.2m). This is the lowest of the 6 months shown."
/// </summary>
public static class CommentaryWriter
{
    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

    public static string Write(FinancialData data, DashboardView view)
    {
        var series = view.Series;
        var kpis = view.Kpis;
        var points = series.Points;
        var latest = points[points.Count - 1];
        var isPercent = series.Metric.IsPercent;

        if (latest.Value is null)
        {
            var reason = latest.Notes.Count > 0 ? " " + string.Join(" ", latest.Notes) : "";
            return $"There is no {series.Metric.DisplayName} figure {Scope(series.Selection)} for {latest.Label} yet.{reason}";
        }

        var sentences = new List<string> { Headline(series, kpis, isPercent) };

        if (!isPercent && BiggestMover(data, view) is string mover)
            sentences.Add(mover);

        if (Extreme(series) is string extreme)
            sentences.Add(extreme);

        if (!latest.IsComplete)
            sentences.Add($"{latest.Label} is incomplete: {string.Join(" ", latest.Notes)}");

        return string.Join(" ", sentences);
    }

    private static string Headline(SeriesResult series, KpiSummary kpis, bool isPercent)
    {
        var text = $"{series.Metric.DisplayName} {Scope(series.Selection)} was {Value(kpis.Latest!.Value, isPercent)} in {kpis.LatestLabel}";

        var changes = new[] { kpis.PreviousPeriod, kpis.SamePeriodLastYear }
            .Where(c => c?.Change is not null)
            .Select(c => $"{Movement(c!.Change!.Value, isPercent, kpis.HigherIsBetter)} on {c.Label}")
            .ToList();

        return changes.Count == 0 ? text + "." : $"{text}, {string.Join(" and ", changes)}.";
    }

    /// <summary>"down 2.7%", "up 0.9 pp (worse)", "unchanged".</summary>
    internal static string Movement(decimal change, bool isPercent, bool higherIsBetter)
    {
        if (Math.Round(change, 1) == 0m) return "unchanged";

        var direction = change > 0m ? "up" : "down";
        var size = Math.Abs(change).ToString("0.0", Invariant) + (isPercent ? " pp" : "%");

        // For metrics where a fall is good news, say so: "expenses up 3%" alone reads as neutral.
        var verdict = higherIsBetter ? "" : (change > 0m ? " (worse)" : " (better)");
        return $"{direction} {size}{verdict}";
    }

    /// <summary>
    /// Which segment (or country) moved most since the previous period. Uses the comparison on the dashboard if there
    /// is one; otherwise splits by whichever of segment or country is "All".
    /// </summary>
    private static string? BiggestMover(FinancialData data, DashboardView view)
    {
        var selection = view.Series.Selection;
        var by = view.BreakdownBy != Breakdown.None ? view.BreakdownBy
            : selection.Segment == Selection.All ? Breakdown.Segment
            : selection.Country == Selection.All ? Breakdown.Country
            : Breakdown.None;
        if (by == Breakdown.None) return null;

        var parts = view.BreakdownBy != Breakdown.None
            ? view.Parts
            : DashboardCalculator.Calculate(data, selection with { PeriodCount = 2 }, by).Parts;

        var moves = parts
            .Where(p => p.Points.Count >= 2 && p.Points[p.Points.Count - 1].Value is not null && p.Points[p.Points.Count - 2].Value is not null)
            .Select(p => (Part: p, Delta: p.Points[p.Points.Count - 1].Value!.Value - p.Points[p.Points.Count - 2].Value!.Value))
            .ToList();
        if (moves.Count < 2) return null;

        var top = moves.OrderByDescending(m => Math.Abs(m.Delta)).First();
        if (top.Delta == 0m) return null;

        var name = by == Breakdown.Segment ? top.Part.Selection.Segment : top.Part.Selection.Country;
        return $"{name} moved most ({(top.Delta > 0m ? "+" : "-")}{Money(Math.Abs(top.Delta))}).";
    }

    /// <summary>"This is the highest of the 12 months shown." Only for three or more periods.</summary>
    private static string? Extreme(SeriesResult series)
    {
        var values = series.Points.Where(p => p.Value is not null).Select(p => p.Value!.Value).ToList();
        if (values.Count < 3) return null;

        var latest = series.Points[series.Points.Count - 1].Value!.Value;
        var unit = series.Selection.Frequency == Frequency.Monthly ? "months" : "quarters";

        if (latest == values.Max()) return $"This is the highest of the {values.Count} {unit} shown.";
        if (latest == values.Min()) return $"This is the lowest of the {values.Count} {unit} shown.";
        return null;
    }

    /// <summary>"for Retail in India", "across all segments in India", "across all segments and countries".</summary>
    internal static string Scope(Selection s) => (s.Segment == Selection.All, s.Country == Selection.All) switch
    {
        (true, true) => "across all segments and countries",
        (true, false) => $"across all segments in {s.Country}",
        (false, true) => $"for {s.Segment} across all countries",
        _ => $"for {s.Segment} in {s.Country}"
    };

    private static string Value(decimal value, bool isPercent) =>
        isPercent ? value.ToString("0.0", Invariant) + "%" : Money(value);

    /// <summary>Amounts are in USD thousands: 10,265 -> "USD 10.3m", 850 -> "USD 850k".</summary>
    internal static string Money(decimal thousands) =>
        Math.Abs(thousands) >= 1000m
            ? $"USD {(thousands / 1000m).ToString("#,##0.0", Invariant)}m"
            : $"USD {thousands.ToString("#,##0", Invariant)}k";
}
