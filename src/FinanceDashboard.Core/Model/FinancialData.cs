namespace FinanceDashboard.Core.Model;

/// <summary>
/// Everything read from the workbook: the month columns, the data rows and the metric definitions.
/// </summary>
public sealed class FinancialData
{
    public FinancialData(
        IReadOnlyList<Period> periods,
        IReadOnlyList<FinancialRecord> records,
        IReadOnlyList<MetricInfo> metrics)
    {
        if (periods.Count == 0)
            throw new ArgumentException("The data has no month columns.", nameof(periods));

        Periods = periods;
        Records = records;
        Metrics = metrics;
    }

    /// <summary>Month columns in sheet order (oldest first).</summary>
    public IReadOnlyList<Period> Periods { get; }

    public IReadOnlyList<FinancialRecord> Records { get; }

    public IReadOnlyList<MetricInfo> Metrics { get; }

    /// <summary>The chart always ends here.</summary>
    public Period LatestPeriod => Periods[Periods.Count - 1];

    public IReadOnlyList<string> Segments =>
        Records.Select(r => r.Segment).Distinct().OrderBy(s => s).ToList();

    public IReadOnlyList<string> Countries =>
        Records.Select(r => r.Country).Distinct().OrderBy(c => c).ToList();

    public MetricInfo GetMetric(string code) =>
        Metrics.FirstOrDefault(m => string.Equals(m.Code, code, StringComparison.OrdinalIgnoreCase))
        ?? throw new ArgumentException($"Unknown metric '{code}'.", nameof(code));
}
