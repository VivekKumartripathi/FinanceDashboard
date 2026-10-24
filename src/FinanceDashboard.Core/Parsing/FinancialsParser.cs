using FinanceDashboard.Core.Model;

namespace FinanceDashboard.Core.Parsing;

/// <summary>
/// Turns the raw cells of the Financials and Metrics sheets into <see cref="FinancialData"/>.
/// Works on plain object arrays so it can be tested without Excel.
/// </summary>
public static class FinancialsParser
{
    public static FinancialData Parse(object?[,] financialsCells, object?[,] metricsCells)
    {
        var (periods, records) = ParseFinancials(financialsCells);
        var metrics = ParseMetrics(metricsCells);
        return new FinancialData(periods, records, metrics);
    }

    public static (IReadOnlyList<Period> Periods, IReadOnlyList<FinancialRecord> Records) ParseFinancials(object?[,] cells)
    {
        const string sheet = "Financials";
        var grid = new SheetGrid(cells);
        var headers = grid.HeaderIndex();

        var segmentCol = SheetGrid.RequireColumn(headers, "Segment", sheet);
        var countryCol = SheetGrid.RequireColumn(headers, "Country", sheet);
        var metricCol = SheetGrid.RequireColumn(headers, "Metric", sheet);
        int? rowIdCol = headers.TryGetValue("RowId", out var r) ? r : null;
        int? notesCol = headers.TryGetValue("Notes", out var n) ? n : null;

        // Month columns are found by their header ("2023M01"), so extra columns such as Notes are ignored.
        var monthColumns = new List<(int Col, Period Period)>();
        for (var col = 0; col < grid.ColumnCount; col++)
        {
            if (Period.TryParse(grid.Text(0, col), out var period))
                monthColumns.Add((col, period));
        }
        if (monthColumns.Count == 0)
            throw new FormatException("The Financials sheet has no month columns (expected headers like 2023M01).");

        monthColumns.Sort((a, b) => a.Period.CompareTo(b.Period));
        var periods = monthColumns.Select(m => m.Period).ToList();

        var records = new List<FinancialRecord>();
        var seen = new HashSet<(string, string, string)>();
        for (var row = 1; row < grid.RowCount; row++)
        {
            var segment = grid.Text(row, segmentCol);
            var country = grid.Text(row, countryCol);
            var metric = grid.Text(row, metricCol);
            if (segment is null && country is null && metric is null)
                continue; // trailing empty row

            if (segment is null || country is null || metric is null)
                throw new FormatException($"Financials row {row + 1} is missing its segment, country or metric.");

            if (!seen.Add((segment, country, metric)))
                throw new FormatException($"Financials has more than one row for {segment} / {country} / {metric}.");

            var values = monthColumns.Select(m => grid.Number(row, m.Col)).ToList();
            var rowId = rowIdCol is int ri ? grid.Text(row, ri) ?? $"row {row + 1}" : $"row {row + 1}";
            var note = notesCol is int nc ? grid.Text(row, nc) : null;

            records.Add(new FinancialRecord(rowId, segment, country, metric, values, note));
        }

        return (periods, records);
    }

    public static IReadOnlyList<MetricInfo> ParseMetrics(object?[,] cells)
    {
        const string sheet = "Metrics";
        var grid = new SheetGrid(cells);
        var headers = grid.HeaderIndex();

        var codeCol = SheetGrid.RequireColumn(headers, "Metric", sheet);
        var nameCol = SheetGrid.RequireColumn(headers, "Display Name", sheet);
        var shownAsCol = SheetGrid.RequireColumn(headers, "Shown as", sheet);
        int? fromCol = headers.TryGetValue("Calculated from", out var f) ? f : null;

        var metrics = new List<MetricInfo>();
        for (var row = 1; row < grid.RowCount; row++)
        {
            var code = grid.Text(row, codeCol);
            if (code is null) continue;

            var displayName = grid.Text(row, nameCol) ?? code;
            var shownAs = grid.Text(row, shownAsCol) ?? "";
            var isPercent = shownAs.IndexOf("percent", StringComparison.OrdinalIgnoreCase) >= 0;
            var calculatedFrom = fromCol is int fc ? SplitMetricList(grid.Text(row, fc)) : [];

            metrics.Add(new MetricInfo(code, displayName, isPercent, calculatedFrom));
        }

        return metrics;
    }

    /// <summary>"Expenses and Revenue" -> ["Expenses", "Revenue"]. Order is kept.</summary>
    internal static IReadOnlyList<string> SplitMetricList(string? text)
    {
        if (text is null) return [];

        return text
            .Replace(" and ", ",").Replace("&", ",")
            .Split(',')
            .Select(part => part.Trim())
            .Where(part => part.Length > 0)
            .ToList();
    }
}
