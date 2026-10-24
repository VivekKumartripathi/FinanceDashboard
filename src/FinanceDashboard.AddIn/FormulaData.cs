using ExcelDna.Integration;
using FinanceDashboard.Core.Model;
using FinanceDashboard.Core.Parsing;

namespace FinanceDashboard.AddIn;

/// <summary>
/// The data behind the worksheet functions, read from the calling workbook through the Excel C API
/// (worksheet functions should not call Excel's COM object model). Cached per workbook, and cleared when the
/// Financials or Metrics sheet changes.
/// </summary>
internal static class FormulaData
{
    private static readonly Dictionary<string, FinancialData> Cache = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>True once a worksheet function has been used, so data edits only trigger a recalculation when needed.</summary>
    public static bool InUse { get; private set; }

    public static void Clear() => Cache.Clear();

    /// <summary>The data for the workbook that holds the calling cell. Only valid inside a macro-type function.</summary>
    public static FinancialData ForCaller()
    {
        InUse = true;

        // "[Book1.xlsx]Sheet1" -> "[Book1.xlsx]"
        var caller = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
        var sheetName = (string)XlCall.Excel(XlCall.xlSheetNm, caller);
        var book = sheetName.Substring(0, sheetName.IndexOf(']') + 1);

        if (!Cache.TryGetValue(book, out var data))
        {
            data = FinancialsParser.Parse(ReadSheet(book, "Financials"), ReadSheet(book, "Metrics"));
            Cache[book] = data;
        }
        return data;
    }

    private static object?[,] ReadSheet(string book, string sheet)
    {
        var name = book + sheet;

        // GET.DOCUMENT(10) and (12): last used row and column of the sheet.
        if (XlCall.Excel(XlCall.xlfGetDocument, 10, name) is not double lastRow ||
            XlCall.Excel(XlCall.xlfGetDocument, 12, name) is not double lastColumn)
            throw new InvalidOperationException($"This workbook has no '{sheet}' sheet.");

        var sheetId = ((ExcelReference)XlCall.Excel(XlCall.xlSheetId, name)).SheetId;
        var range = new ExcelReference(0, (int)lastRow - 1, 0, (int)lastColumn - 1, sheetId);

        if (range.GetValue() is not object[,] values)
            throw new InvalidOperationException($"The {sheet} sheet is empty.");

        // The C API marks empty cells with ExcelEmpty and errors with ExcelError; the parser expects null for both.
        var cells = new object?[values.GetLength(0), values.GetLength(1)];
        for (var r = 0; r < values.GetLength(0); r++)
            for (var c = 0; c < values.GetLength(1); c++)
                cells[r, c] = values[r, c] is ExcelEmpty or ExcelError ? null : values[r, c];
        return cells;
    }
}
