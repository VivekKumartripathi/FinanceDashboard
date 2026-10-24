using System.Globalization;

namespace FinanceDashboard.Core.Parsing;

/// <summary>
/// Read-only view over a 2-D cell array as returned by Excel's Range.Value2.
/// Excel arrays are 1-based; this hides the bounds so callers use 0-based rows and columns.
/// </summary>
internal sealed class SheetGrid
{
    private readonly object?[,] _cells;
    private readonly int _rowOffset;
    private readonly int _colOffset;

    public SheetGrid(object?[,] cells)
    {
        _cells = cells ?? throw new ArgumentNullException(nameof(cells));
        _rowOffset = cells.GetLowerBound(0);
        _colOffset = cells.GetLowerBound(1);
    }

    public int RowCount => _cells.GetLength(0);

    public int ColumnCount => _cells.GetLength(1);

    public object? this[int row, int col] => _cells[row + _rowOffset, col + _colOffset];

    /// <summary>Cell as trimmed text; null when the cell is empty or whitespace.</summary>
    public string? Text(int row, int col)
    {
        var text = this[row, col] switch
        {
            null => null,
            string s => s,
            double d => d.ToString(CultureInfo.InvariantCulture),
            var other => Convert.ToString(other, CultureInfo.InvariantCulture)
        };
        return string.IsNullOrWhiteSpace(text) ? null : text!.Trim();
    }

    /// <summary>Cell as a number; null when empty or not numeric (e.g. a note typed into a month cell).</summary>
    public decimal? Number(int row, int col) => this[row, col] switch
    {
        double d => (decimal)d,
        int i => i,
        decimal m => m,
        string s when decimal.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out var parsed) => parsed,
        _ => null
    };

    /// <summary>Maps each non-empty header in row 0 to its column index (case-insensitive).</summary>
    public Dictionary<string, int> HeaderIndex()
    {
        var index = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var col = 0; col < ColumnCount; col++)
        {
            var header = Text(0, col);
            if (header != null && !index.ContainsKey(header))
                index[header] = col;
        }
        return index;
    }

    public static int RequireColumn(Dictionary<string, int> headers, string name, string sheet) =>
        headers.TryGetValue(name, out var col)
            ? col
            : throw new FormatException($"The {sheet} sheet has no '{name}' column.");
}
