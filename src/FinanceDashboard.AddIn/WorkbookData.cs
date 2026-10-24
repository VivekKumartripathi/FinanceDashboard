using System.Runtime.InteropServices;
using FinanceDashboard.Core.Model;
using FinanceDashboard.Core.Parsing;
using Microsoft.Office.Interop.Excel;

namespace FinanceDashboard.AddIn;

/// <summary>
/// Reads the Financials and Metrics sheets from a workbook. Read-only: the data sheets are never changed.
/// </summary>
internal static class WorkbookData
{
    public static FinancialData Load(Workbook? workbook)
    {
        if (workbook is null)
            throw new InvalidOperationException("Open the Financials workbook first.");

        return FinancialsParser.Parse(ReadSheet(workbook, "Financials"), ReadSheet(workbook, "Metrics"));
    }

    /// <summary>
    /// One Value2 call per sheet returns every cell at once; reading cell by cell over COM is very slow.
    /// </summary>
    private static object?[,] ReadSheet(Workbook workbook, string name)
    {
        Worksheet sheet;
        try
        {
            sheet = (Worksheet)workbook.Worksheets[name];
        }
        catch (COMException)
        {
            throw new InvalidOperationException($"This workbook has no '{name}' sheet.");
        }

        object? value = sheet.UsedRange.Value2;
        return value as object?[,] ?? throw new InvalidOperationException($"The {name} sheet is empty.");
    }
}
