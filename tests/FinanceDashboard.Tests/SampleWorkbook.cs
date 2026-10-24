using ClosedXML.Excel;
using FinanceDashboard.Core.Model;
using FinanceDashboard.Core.Parsing;

namespace FinanceDashboard.Tests;

/// <summary>
/// Loads the real Financials.xlsx once, the same way the add-in will:
/// each sheet's used range as a 1-based object array, then through <see cref="FinancialsParser"/>.
/// </summary>
public static class SampleWorkbook
{
    private static readonly Lazy<FinancialData> Lazy = new(Load);

    public static FinancialData Data => Lazy.Value;

    private static FinancialData Load()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Financials.xlsx");
        using var workbook = new XLWorkbook(path);
        return FinancialsParser.Parse(
            ReadUsedRange(workbook.Worksheet("Financials")),
            ReadUsedRange(workbook.Worksheet("Metrics")));
    }

    private static object?[,] ReadUsedRange(IXLWorksheet sheet)
    {
        var range = sheet.RangeUsed() ?? throw new InvalidOperationException($"{sheet.Name} is empty.");
        var rows = range.RowCount();
        var cols = range.ColumnCount();
        var cells = (object?[,])Array.CreateInstance(typeof(object), [rows, cols], [1, 1]);

        for (var r = 1; r <= rows; r++)
        {
            for (var c = 1; c <= cols; c++)
            {
                var value = range.Cell(r, c).Value;
                cells[r, c] = value.IsBlank ? null
                    : value.IsNumber ? value.GetNumber()
                    : value.ToString();
            }
        }
        return cells;
    }
}
