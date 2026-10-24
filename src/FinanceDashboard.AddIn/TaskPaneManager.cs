using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;

namespace FinanceDashboard.AddIn;

/// <summary>
/// Creates the Finance Dashboard task pane once and shows it again on later clicks.
/// </summary>
internal static class TaskPaneManager
{
    private static CustomTaskPane? _pane;

    private static Application Excel => (Application)ExcelDnaUtil.Application;

    public static void Show()
    {
        if (_pane is null)
        {
            _pane = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(DashboardPane), "Finance Dashboard");
            _pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            _pane.Width = 380;

            var control = (DashboardPane)_pane.ContentControl;
            control.DataSource = () => WorkbookData.Load(Excel.ActiveWorkbook);
            control.Calculated += view =>
            {
                try
                {
                    DashboardRenderer.Render(Excel.ActiveWorkbook, view);
                }
                catch (COMException ex)
                {
                    control.ShowError($"Could not update the Dashboard sheet: {ex.Message}");
                }
            };
        }

        _pane.Visible = true;
        ((DashboardPane)_pane.ContentControl).Reload();
    }

    public static void Close()
    {
        _pane?.Delete();
        _pane = null;
    }
}
