using ExcelDna.Integration;

namespace FinanceDashboard.AddIn;

/// <summary>
/// Excel-DNA entry point. Excel calls AutoOpen when the add-in loads
/// and AutoClose when it is removed. The ribbon is discovered by Excel-DNA automatically.
/// </summary>
public class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
    }

    public void AutoClose()
    {
        TaskPaneManager.Close();
    }
}
