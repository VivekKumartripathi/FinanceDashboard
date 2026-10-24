namespace FinanceDashboard.Core.Calculation;

public enum Aggregation
{
    /// <summary>Flows (Revenue, Expenses): a quarter is the sum of its months.</summary>
    Sum,

    /// <summary>Balances (Loans outstanding): a quarter is the balance at its last month.</summary>
    PeriodEnd
}

/// <summary>
/// How each metric rolls up from months to quarters. The Metrics sheet does not say,
/// so this is our decision; E1 confirms Loans uses the quarter-end balance.
/// </summary>
public static class AggregationRules
{
    private static readonly HashSet<string> Balances = new(StringComparer.OrdinalIgnoreCase) { "Loans" };

    public static Aggregation For(string metricCode) =>
        Balances.Contains(metricCode) ? Aggregation.PeriodEnd : Aggregation.Sum;
}
