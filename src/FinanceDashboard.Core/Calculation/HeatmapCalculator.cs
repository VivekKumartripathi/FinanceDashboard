using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Core.Calculation;

/// <summary>One cell of the heatmap; <see cref="Selection.All"/> in either name marks a total.</summary>
public sealed record HeatmapCell(string Segment, string Country, decimal? Value, bool IsComplete);

/// <summary>
/// Every segment against every country for the latest period, with totals: shows where a metric is high or low.
/// </summary>
public sealed record Heatmap(
    string PeriodLabel,
    IReadOnlyList<string> Segments,
    IReadOnlyList<string> Countries,
    IReadOnlyList<HeatmapCell> Cells)
{
    public HeatmapCell Cell(string segment, string country) =>
        Cells.First(c => c.Segment == segment && c.Country == country);
}

public static class HeatmapCalculator
{
    /// <summary>
    /// Each cell, including the totals, is calculated by <see cref="SeriesCalculator"/> exactly as the dashboard
    /// would for that segment and country, so cost-to-income totals follow the same rules as E3 and E4.
    /// </summary>
    public static Heatmap Calculate(FinancialData data, string metric, Frequency frequency)
    {
        var segments = data.Segments;
        var countries = data.Countries;
        var cells = new List<HeatmapCell>();
        var label = "";

        foreach (var segment in segments.Append(Selection.All))
        {
            foreach (var country in countries.Append(Selection.All))
            {
                try
                {
                    var point = SeriesCalculator.Calculate(data, new Selection(metric, segment, country, frequency, 1)).Points[0];
                    label = point.Label;
                    cells.Add(new HeatmapCell(segment, country, point.Value, point.IsComplete));
                }
                catch (ArgumentException)
                {
                    // That segment does not operate in that country.
                    cells.Add(new HeatmapCell(segment, country, null, true));
                }
            }
        }

        return new Heatmap(label, segments, countries, cells);
    }
}
