namespace FinanceDashboard.Core.Model;

/// <summary>
/// One point on the chart.
/// </summary>
/// <param name="Label">"Jan 25" or "Q3 24".</param>
/// <param name="Value">USD thousands, or a percent (39.0 means 39.0%); null when it cannot be calculated.</param>
/// <param name="Notes">Data-quality warnings for this point, e.g. a blank source cell.</param>
public sealed record SeriesPoint(
    string Label,
    decimal? Value,
    IReadOnlyList<string> Notes)
{
    public bool IsComplete => Notes.Count == 0;
}

/// <summary>
/// The calculated series for one <see cref="Selection"/>, ready to chart.
/// </summary>
public sealed record SeriesResult(
    Selection Selection,
    MetricInfo Metric,
    IReadOnlyList<SeriesPoint> Points);
