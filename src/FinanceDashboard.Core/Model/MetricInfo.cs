namespace FinanceDashboard.Core.Model;

/// <summary>
/// One row of the Metrics sheet.
/// </summary>
/// <param name="Code">Metric name used in the Financials sheet, e.g. "Loans".</param>
/// <param name="DisplayName">Name to show the user, e.g. "Loans outstanding".</param>
/// <param name="IsPercent">True when shown as a percent rather than USD thousands.</param>
/// <param name="CalculatedFrom">Source metric codes for derived metrics (e.g. Expenses, Revenue); empty otherwise.</param>
public sealed record MetricInfo(
    string Code,
    string DisplayName,
    bool IsPercent,
    IReadOnlyList<string> CalculatedFrom)
{
    /// <summary>True when the metric is not in the data and must be derived.</summary>
    public bool IsCalculated => CalculatedFrom.Count > 0;
}
