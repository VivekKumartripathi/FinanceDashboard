using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace FinanceDashboard.AddIn;

/// <summary>
/// Excel-DNA entry point. Excel calls AutoOpen when the add-in loads
/// and AutoClose when it is removed. The ribbon and the FIN.* worksheet functions are discovered by Excel-DNA automatically.
/// </summary>
public class AddIn : IExcelAddIn
{
    // Held in a field so the COM event subscription stays alive.
    private Application? _excel;

    public void AutoOpen()
    {
        _excel = (Application)ExcelDnaUtil.Application;
        _excel.SheetChange += OnSheetChange;
    }

    public void AutoClose()
    {
        if (_excel != null)
            _excel.SheetChange -= OnSheetChange;
        TaskPaneManager.Close();
    }

    /// <summary>
    /// When the data changes, drop the cached copy the FIN.* functions use and recalculate,
    /// so formulas never show stale numbers.
    /// </summary>
    private void OnSheetChange(object sheet, Range target)
    {
        if (sheet is not Worksheet { Name: "Financials" or "Metrics" }) return;

        FormulaData.Clear();
        if (FormulaData.InUse)
            ExcelAsyncUtil.QueueAsMacro(() => _excel?.CalculateFull());
    }
}
