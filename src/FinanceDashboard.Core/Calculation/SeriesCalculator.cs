using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Core.Calculation;

/// <summary>
/// Turns a <see cref="Selection"/> into the series to chart.
/// </summary>
public static class SeriesCalculator
{
    public static SeriesResult Calculate(FinancialData data, Selection selection)
    {
        if (selection.PeriodCount < 1)
            throw new ArgumentOutOfRangeException(nameof(selection), "Choose at least one period.");

        var metric = data.GetMetric(selection.Metric);
        if (metric.IsCalculated)
            throw new NotSupportedException($"{metric.DisplayName} is not supported yet.");
        if (selection.Frequency != Frequency.Monthly)
            throw new NotSupportedException("Quarterly periods are not supported yet.");

        var totals = MonthlyTotals(Filter(data, metric.Code, selection.Segment, selection.Country), data.Periods.Count);

        // The chart always ends at the latest month; asking for more periods than exist shows them all.
        var count = Math.Min(selection.PeriodCount, data.Periods.Count);
        var first = data.Periods.Count - count;

        var points = Enumerable.Range(first, count)
            .Select(i => new SeriesPoint(data.Periods[i].MonthLabel, totals[i], []))
            .ToList();

        return new SeriesResult(selection, metric, points);
    }

    /// <summary>
    /// Rows for one metric, limited to the chosen segment and country. "All" means no limit.
    /// </summary>
    internal static IReadOnlyList<FinancialRecord> Filter(FinancialData data, string metric, string segment, string country)
    {
        var rows = data.Records
            .Where(r => Matches(r.Metric, metric) && Matches(r.Segment, segment) && Matches(r.Country, country))
            .ToList();

        if (rows.Count == 0)
            throw new ArgumentException($"There is no {metric} data for segment '{segment}' and country '{country}'.");

        return rows;
    }

    /// <summary>
    /// Adds the rows together month by month. A month where every row is blank stays null.
    /// </summary>
    internal static decimal?[] MonthlyTotals(IReadOnlyList<FinancialRecord> rows, int periodCount)
    {
        var totals = new decimal?[periodCount];
        foreach (var row in rows)
        {
            for (var i = 0; i < periodCount; i++)
            {
                if (row.Values[i] is decimal value)
                    totals[i] = (totals[i] ?? 0m) + value;
            }
        }
        return totals;
    }

    private static bool Matches(string value, string wanted) =>
        wanted == Selection.All || string.Equals(value, wanted, StringComparison.OrdinalIgnoreCase);
}
