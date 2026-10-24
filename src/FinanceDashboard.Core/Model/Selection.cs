namespace FinanceDashboard.Core.Model;

public enum Frequency
{
    Monthly,
    Quarterly
}

/// <summary>
/// What the user picked in the task pane.
/// </summary>
/// <param name="Metric">Metric code, e.g. "Revenue" or "CostToIncome".</param>
/// <param name="Segment">A segment name, or <see cref="All"/>.</param>
/// <param name="Country">A country name, or <see cref="All"/>.</param>
/// <param name="PeriodCount">How many months or quarters to show, ending at the latest month.</param>
public sealed record Selection(
    string Metric,
    string Segment,
    string Country,
    Frequency Frequency,
    int PeriodCount)
{
    public const string All = "All";
}
