using System.Globalization;
using System.Text.RegularExpressions;

namespace FinanceDashboard.Core.Model;

/// <summary>
/// Reads a period typed by a user, e.g. in a worksheet formula, and turns it into the label the calculator uses.
/// Accepts "Jun 25", "Jun 2025", "2025M06", "Q2 25", "Q2 2025" and "2025Q2".
/// </summary>
public static class PeriodLabel
{
    private static readonly Regex MonthName = new(@"^([A-Za-z]{3})[a-z]*\s+(\d{2}|\d{4})$", RegexOptions.Compiled);
    private static readonly Regex Quarter = new(@"^Q([1-4])\s*(\d{2}|\d{4})$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex YearQuarter = new(@"^(\d{4})\s*Q([1-4])$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public static bool TryParse(string? text, out Frequency frequency, out string label)
    {
        frequency = Frequency.Monthly;
        label = "";
        if (string.IsNullOrWhiteSpace(text)) return false;
        text = text!.Trim();

        if (Period.TryParse(text, out var period))
        {
            label = period.MonthLabel;
            return true;
        }

        var m = MonthName.Match(text);
        if (m.Success && DateTime.TryParseExact(m.Groups[1].Value, "MMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out var month))
        {
            label = new Period(Year(m.Groups[2].Value), month.Month).MonthLabel;
            return true;
        }

        var q = Quarter.Match(text);
        var yq = YearQuarter.Match(text);
        if (q.Success || yq.Success)
        {
            var quarter = int.Parse(q.Success ? q.Groups[1].Value : yq.Groups[2].Value, CultureInfo.InvariantCulture);
            var year = Year(q.Success ? q.Groups[2].Value : yq.Groups[1].Value);
            frequency = Frequency.Quarterly;
            label = new Period(year, quarter * 3).QuarterLabel;
            return true;
        }

        return false;
    }

    private static int Year(string digits)
    {
        var year = int.Parse(digits, CultureInfo.InvariantCulture);
        return digits.Length == 2 ? 2000 + year : year;
    }
}
