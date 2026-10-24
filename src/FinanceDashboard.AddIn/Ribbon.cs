using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;

namespace FinanceDashboard.AddIn;

/// <summary>
/// Adds a "Finance" tab with one button that opens the dashboard pane.
/// </summary>
[ComVisible(true)]
public class Ribbon : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId) => """
        <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
          <ribbon>
            <tabs>
              <tab id="financeTab" label="Finance">
                <group id="dashboardGroup" label="Dashboard">
                  <button id="showDashboard" label="Finance Dashboard" size="large"
                          imageMso="ChartTypeLineInsertGallery" onAction="OnShowDashboard"
                          screentip="Open the Finance Dashboard pane" />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>
        """;

    public void OnShowDashboard(IRibbonControl control)
    {
        try
        {
            TaskPaneManager.Show();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Finance Dashboard", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
