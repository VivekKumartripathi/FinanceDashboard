using ExcelDna.Integration;
using FinanceDashboard.Core.Calculation;
using FinanceDashboard.Core.Model;

namespace FinanceDashboard.AddIn;

/// <summary>
/// Worksheet functions that use the same calculation engine as the dashboard, so analysts can pull the numbers
/// into their own models:
///   =FIN.VALUE("Revenue", "All", "India", "Jun 25")
///   =FIN.SERIES("CostToIncome", "All", "India", "Q", 4)
///   =FIN.COMMENTARY("Revenue", "All", "India")
/// Missing data returns #N/A rather than zero, matching the dashboard. Bad arguments return #VALUE!.
/// </summary>
public static class FinanceFunctions
{
    private const string Category = "Finance Dashboard";

    [ExcelFunction(Name = "FIN.VALUE", Category = Category, IsMacroType = true,
        Description = "One value from the Financials data, calculated like the dashboard.")]
    public static object Value(
        [ExcelArgument(Name = "metric", Description = "Loans, Revenue, Expenses, CostToIncome, or a display name")] string metric,
        [ExcelArgument(Name = "segment", Description = "Corporate, Retail, Treasury, Wealth or All")] string segment,
        [ExcelArgument(Name = "country", Description = "India, UK, USA or All")] string country,
        [ExcelArgument(Name = "period", Description = "e.g. \"Jun 25\", \"2025M06\", \"Q2 25\" or \"2025Q2\"")] string period)
    {
        return Guard(() =>
        {
            if (!PeriodLabel.TryParse(period, out var frequency, out var label))
                return ExcelError.ExcelErrorValue;

            var points = Calculate(metric, segment, country, frequency, int.MaxValue).Points;
            var point = points.FirstOrDefault(p => p.Label == label);
            return point?.Value is decimal v ? (double)v : ExcelError.ExcelErrorNA;
        });
    }

    [ExcelFunction(Name = "FIN.SERIES", Category = Category, IsMacroType = true,
        Description = "The last N periods as a two-column array (period, value). Spills in Excel 365.")]
    public static object Series(
        [ExcelArgument(Name = "metric")] string metric,
        [ExcelArgument(Name = "segment")] string segment,
        [ExcelArgument(Name = "country")] string country,
        [ExcelArgument(Name = "[frequency]", Description = "\"M\" (monthly, default) or \"Q\" (quarterly)")] object frequency,
        [ExcelArgument(Name = "[periods]", Description = "How many periods, ending at the latest month (default 12)")] object periods)
    {
        return Guard(() =>
        {
            if (ParseFrequency(frequency) is not Frequency f || ParseCount(periods) is not int n)
                return ExcelError.ExcelErrorValue;

            var points = Calculate(metric, segment, country, f, n).Points;
            var result = new object[points.Count, 2];
            for (var i = 0; i < points.Count; i++)
            {
                result[i, 0] = points[i].Label;
                result[i, 1] = points[i].Value is decimal v ? (double)v : ExcelError.ExcelErrorNA;
            }
            return result;
        });
    }

    [ExcelFunction(Name = "FIN.COMMENTARY", Category = Category, IsMacroType = true,
        Description = "The dashboard's plain-English summary for a selection.")]
    public static object Commentary(
        [ExcelArgument(Name = "metric")] string metric,
        [ExcelArgument(Name = "segment")] string segment,
        [ExcelArgument(Name = "country")] string country,
        [ExcelArgument(Name = "[frequency]", Description = "\"M\" (monthly, default) or \"Q\" (quarterly)")] object frequency,
        [ExcelArgument(Name = "[periods]", Description = "How many periods (default 12)")] object periods)
    {
        return Guard(() =>
        {
            if (ParseFrequency(frequency) is not Frequency f || ParseCount(periods) is not int n)
                return ExcelError.ExcelErrorValue;

            var data = FormulaData.ForCaller();
            return DashboardCalculator.Calculate(data, new Selection(metric, segment, country, f, n)).Commentary;
        });
    }

    private static SeriesResult Calculate(string metric, string segment, string country, Frequency frequency, int periods) =>
        SeriesCalculator.Calculate(FormulaData.ForCaller(), new Selection(metric, segment, country, frequency, periods));

    /// <summary>Exceptions must never reach Excel from a worksheet function; they become #VALUE!.</summary>
    private static object Guard(Func<object> body)
    {
        try
        {
            return body();
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or FormatException)
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    private static Frequency? ParseFrequency(object value) => value switch
    {
        ExcelMissing or ExcelEmpty => Frequency.Monthly,
        string s when s.Trim().StartsWith("M", StringComparison.OrdinalIgnoreCase) => Frequency.Monthly,
        string s when s.Trim().StartsWith("Q", StringComparison.OrdinalIgnoreCase) => Frequency.Quarterly,
        _ => null
    };

    private static int? ParseCount(object value) => value switch
    {
        ExcelMissing or ExcelEmpty => 12,
        double d when d >= 1 && d == Math.Floor(d) => (int)d,
        _ => null
    };
}
