namespace FinanceDashboard.Core.Model;

/// <summary>
/// One row of the Financials sheet: a segment, country and metric with one value per month.
/// </summary>
/// <param name="Values">One entry per period in <see cref="FinancialData.Periods"/>; null means the cell was blank.</param>
/// <param name="Note">Free text from the Notes column, if any (e.g. "Jun-25 not closed yet").</param>
public sealed record FinancialRecord(
    string RowId,
    string Segment,
    string Country,
    string Metric,
    IReadOnlyList<decimal?> Values,
    string? Note);
