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

        var totals = MonthlyTotals(Filter(data, metric.Code, selection.Segment, selection.Country), data.Periods.Count);
        var aggregation = AggregationRules.For(metric.Code);

        // The chart always ends at the latest month; asking for more periods than exist shows them all.
        var buckets = Buckets(data.Periods, selection.Frequency);
        var points = buckets
            .Skip(Math.Max(0, buckets.Count - selection.PeriodCount))
            .Select(b => new SeriesPoint(b.Label, Aggregate(totals, b, aggregation), b.Notes))
            .ToList();

        return new SeriesResult(selection, metric, points);
    }

    /// <summary>
    /// One chart period: its label and the month indexes (into <see cref="FinancialData.Periods"/>) it covers.
    /// </summary>
    internal sealed record Bucket(string Label, IReadOnlyList<int> MonthIndexes, IReadOnlyList<string> Notes);

    /// <summary>
    /// Groups the months into chart periods: one per month, or one per calendar quarter (Jan-Mar, Apr-Jun, ...).
    /// A quarter the data only partly covers is kept, so the chart still ends at the latest month, and is flagged.
    /// </summary>
    internal static IReadOnlyList<Bucket> Buckets(IReadOnlyList<Period> periods, Frequency frequency)
    {
        if (frequency == Frequency.Monthly)
            return periods.Select((p, i) => new Bucket(p.MonthLabel, [i], [])).ToList();

        return periods
            .Select((p, i) => (Period: p, Index: i))
            .GroupBy(x => (x.Period.Year, x.Period.Quarter))
            .Select(g =>
            {
                var indexes = g.Select(x => x.Index).ToList();
                var label = g.First().Period.QuarterLabel;
                IReadOnlyList<string> notes = indexes.Count < 3
                    ? [$"{label} covers only {indexes.Count} of 3 months."]
                    : [];
                return new Bucket(label, indexes, notes);
            })
            .ToList();
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

    /// <summary>
    /// Rolls monthly totals up into one bucket: the sum of its months, or the value at its last month.
    /// </summary>
    internal static decimal? Aggregate(decimal?[] monthlyTotals, Bucket bucket, Aggregation aggregation)
    {
        if (aggregation == Aggregation.PeriodEnd)
            return monthlyTotals[bucket.MonthIndexes[bucket.MonthIndexes.Count - 1]];

        decimal? sum = null;
        foreach (var i in bucket.MonthIndexes)
        {
            if (monthlyTotals[i] is decimal value)
                sum = (sum ?? 0m) + value;
        }
        return sum;
    }

    private static bool Matches(string value, string wanted) =>
        wanted == Selection.All || string.Equals(value, wanted, StringComparison.OrdinalIgnoreCase);
}
