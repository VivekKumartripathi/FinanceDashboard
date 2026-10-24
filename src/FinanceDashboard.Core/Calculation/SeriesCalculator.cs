using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Core.Calculation;

/// <summary>
/// Turns a <see cref="Selection"/> into the series to chart.
/// </summary>
/// <remarks>
/// Blank cells are never treated as zero. Amounts are calculated from the cells that do have values and the
/// point is flagged, quoting the sheet's note for that row. Ratios leave out any month that is blank on either
/// side, so a missing month of expenses cannot make cost-to-income look better than it is.
/// </remarks>
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
        var rows = Filter(data, metric.Code, selection.Segment, selection.Country);
        var totals = MonthlyTotals(rows, data.Periods.Count);
        var aggregation = AggregationRules.For(metric.Code);

        return buckets
            .Select(b =>
            {
                var notes = b.Notes.Concat(BlankNotes(data, rows, MonthsUsed(b, aggregation), "")).ToList();
                return new SeriesPoint(b.Label, Aggregate(totals, b, aggregation), notes);
            })
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

        // Numerator and denominator rows for each segment that has both in the chosen country.
        var perSegment = new List<(List<FinancialRecord> Numerator, List<FinancialRecord> Denominator)>();
        foreach (var segment in segments)
        {
            var numeratorRows = Rows(data, numeratorCode, segment, selection.Country);
            var denominatorRows = Rows(data, denominatorCode, segment, selection.Country);
            if (numeratorRows.Count > 0 && denominatorRows.Count > 0)
                perSegment.Add((numeratorRows, denominatorRows));
        }

        if (perSegment.Count == 0)
            throw new ArgumentException(
                $"There is no {numeratorCode} and {denominatorCode} data for segment '{selection.Segment}' and country '{selection.Country}'.");

        return buckets
            .Select(b =>
            {
                var notes = new List<string>(b.Notes);
                var ratios = new List<decimal>();

                foreach (var (numeratorRows, denominatorRows) in perSegment)
                {
                    var sourceRows = numeratorRows.Concat(denominatorRows).ToList();

                    // Only months where every source cell has a value, so both sides cover the same months.
                    var used = b.MonthIndexes.Where(i => sourceRows.All(r => r.Values[i].HasValue)).ToList();
                    var skipped = b.MonthIndexes.Except(used);
                    notes.AddRange(BlankNotes(data, sourceRows, skipped, "; that month is left out of the ratio"));

                    var ratio = Ratio(SumOver(numeratorRows, used), SumOver(denominatorRows, used));
                    if (ratio.HasValue)
                        ratios.Add(ratio.Value);
                }

                decimal? value = ratios.Count > 0 ? ratios.Average() : null;
                return new SeriesPoint(b.Label, value, notes);
            })
            .ToList();
    }

    /// <summary>Numerator / denominator as a percent; null when either is missing or the denominator is not positive.</summary>
    private static decimal? Ratio(decimal? numerator, decimal? denominator) =>
        numerator is decimal n && denominator is decimal d && d > 0 ? n / d * 100m : null;

    /// <summary>Sum of the rows over the given months; null when there are no months.</summary>
    private static decimal? SumOver(IReadOnlyList<FinancialRecord> rows, IReadOnlyList<int> months) =>
        months.Count == 0 ? null : rows.Sum(r => months.Sum(i => r.Values[i] ?? 0m));

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
        decimal? result = null;
        foreach (var i in MonthsUsed(bucket, aggregation))
        {
            if (monthlyTotals[i] is decimal value)
                result = (result ?? 0m) + value;
        }
        return result;
    }

    /// <summary>The months a bucket's value is built from: all of them for a sum, only the last for a balance.</summary>
    private static IReadOnlyList<int> MonthsUsed(Bucket bucket, Aggregation aggregation) =>
        aggregation == Aggregation.PeriodEnd
            ? [bucket.MonthIndexes[bucket.MonthIndexes.Count - 1]]
            : bucket.MonthIndexes;

    /// <summary>One note per blank cell among the rows and months, quoting the row's Notes text when there is one.</summary>
    private static IEnumerable<string> BlankNotes(FinancialData data, IEnumerable<FinancialRecord> rows, IEnumerable<int> months, string suffix) =>
        from i in months
        from r in rows
        where !r.Values[i].HasValue
        select $"{r.Segment} / {r.Country} / {r.Metric} is blank for {data.Periods[i].MonthLabel}"
             + (r.Note is null ? "" : $" ({r.Note})")
             + suffix + ".";

    private static bool Matches(string value, string wanted) =>
        wanted == Selection.All || string.Equals(value, wanted, StringComparison.OrdinalIgnoreCase);
}
