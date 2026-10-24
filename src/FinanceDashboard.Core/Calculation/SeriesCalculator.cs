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

        // The chart always ends at the latest month; asking for more periods than exist shows them all.
        var buckets = Buckets(data.Periods, selection.Frequency);
        var shown = buckets.Skip(Math.Max(0, buckets.Count - selection.PeriodCount)).ToList();

        var points = metric.IsCalculated
            ? RatioPoints(data, selection, metric, shown)
            : AmountPoints(data, selection, metric, shown);

        return new SeriesResult(selection, metric, points);
    }

    private static List<SeriesPoint> AmountPoints(FinancialData data, Selection selection, MetricInfo metric, IReadOnlyList<Bucket> buckets)
    {
        var totals = MonthlyTotals(Filter(data, metric.Code, selection.Segment, selection.Country), data.Periods.Count);
        var aggregation = AggregationRules.For(metric.Code);

        return buckets
            .Select(b => new SeriesPoint(b.Label, Aggregate(totals, b, aggregation), b.Notes))
            .ToList();
    }

    /// <summary>
    /// A percent metric calculated from two others, e.g. Cost-to-income % = Expenses / Revenue.
    /// For one segment it is the ratio of that segment's totals over the period.
    /// For "All" segments it is the simple average of each segment's ratio: this is what worked example E4
    /// requires (the ratio of the combined totals gives 53.3% for Q3 24, not 57.0%).
    /// Countries within a segment are added together before the ratio is taken.
    /// </summary>
    private static List<SeriesPoint> RatioPoints(FinancialData data, Selection selection, MetricInfo metric, IReadOnlyList<Bucket> buckets)
    {
        if (metric.CalculatedFrom.Count != 2)
            throw new NotSupportedException($"{metric.DisplayName} must be calculated from exactly two metrics.");

        var numeratorCode = metric.CalculatedFrom[0];
        var denominatorCode = metric.CalculatedFrom[1];

        IReadOnlyList<string> segments = selection.Segment == Selection.All ? data.Segments : [selection.Segment];

        // Monthly numerator and denominator totals for each segment that has both in the chosen country.
        var perSegment = new List<(decimal?[] Numerator, decimal?[] Denominator)>();
        foreach (var segment in segments)
        {
            var numeratorRows = Rows(data, numeratorCode, segment, selection.Country);
            var denominatorRows = Rows(data, denominatorCode, segment, selection.Country);
            if (numeratorRows.Count == 0 || denominatorRows.Count == 0)
                continue;

            perSegment.Add((MonthlyTotals(numeratorRows, data.Periods.Count), MonthlyTotals(denominatorRows, data.Periods.Count)));
        }

        if (perSegment.Count == 0)
            throw new ArgumentException(
                $"There is no {numeratorCode} and {denominatorCode} data for segment '{selection.Segment}' and country '{selection.Country}'.");

        return buckets
            .Select(b =>
            {
                var ratios = perSegment
                    .Select(s => Ratio(Aggregate(s.Numerator, b, Aggregation.Sum), Aggregate(s.Denominator, b, Aggregation.Sum)))
                    .Where(r => r.HasValue)
                    .Select(r => r!.Value)
                    .ToList();

                decimal? value = ratios.Count > 0 ? ratios.Average() : null;
                return new SeriesPoint(b.Label, value, b.Notes);
            })
            .ToList();
    }

    /// <summary>Numerator / denominator as a percent; null when either is missing or the denominator is not positive.</summary>
    private static decimal? Ratio(decimal? numerator, decimal? denominator) =>
        numerator is decimal n && denominator is decimal d && d > 0 ? n / d * 100m : null;

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
        var rows = Rows(data, metric, segment, country);
        if (rows.Count == 0)
            throw new ArgumentException($"There is no {metric} data for segment '{segment}' and country '{country}'.");

        return rows;
    }

    private static List<FinancialRecord> Rows(FinancialData data, string metric, string segment, string country) =>
        data.Records
            .Where(r => Matches(r.Metric, metric) && Matches(r.Segment, segment) && Matches(r.Country, country))
            .ToList();

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
