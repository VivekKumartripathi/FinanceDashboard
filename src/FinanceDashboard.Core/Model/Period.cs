using System.Globalization;
using System.Text.RegularExpressions;

namespace FinanceDashboard.Core.Model;

/// <summary>
/// A calendar month, as used in the Financials sheet headers ("2023M01").
/// </summary>
public readonly record struct Period(int Year, int Month) : IComparable<Period>
{
    private static readonly Regex HeaderPattern = new(@"^(\d{4})M(\d{2})$", RegexOptions.Compiled);

    /// <summary>Quarter number 1-4 (Jan-Mar = 1).</summary>
    public int Quarter => (Month - 1) / 3 + 1;

    /// <summary>True for Mar, Jun, Sep and Dec.</summary>
    public bool IsQuarterEnd => Month % 3 == 0;

    /// <summary>"Jan 25"</summary>
    public string MonthLabel =>
        new DateTime(Year, Month, 1).ToString("MMM yy", CultureInfo.InvariantCulture);

    /// <summary>"Q3 24"</summary>
    public string QuarterLabel => $"Q{Quarter} {Year % 100:00}";

    public Period AddMonths(int months)
    {
        var index = Year * 12 + (Month - 1) + months;
        return new Period(index / 12, index % 12 + 1);
    }

    /// <summary>Parses a header such as "2023M01". Returns false for anything else (e.g. "Notes").</summary>
    public static bool TryParse(string? text, out Period period)
    {
        period = default;
        if (text is null) return false;

        var match = HeaderPattern.Match(text.Trim());
        if (!match.Success) return false;

        var year = int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
        var month = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
        if (month is < 1 or > 12) return false;

        period = new Period(year, month);
        return true;
    }

    public int CompareTo(Period other) =>
        Year != other.Year ? Year.CompareTo(other.Year) : Month.CompareTo(other.Month);

    public override string ToString() => $"{Year}M{Month:00}";
}
