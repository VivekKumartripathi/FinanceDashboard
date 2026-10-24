using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace FinanceDashboard.AddIn
{
    /// <summary>
    /// Excel-DNA entry point. Excel calls AutoOpen when the add-in loads
    /// and AutoClose when it is removed.
    /// </summary>
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            MessageBox.Show("Finance Dashboard add-in loaded.", "Finance Dashboard");
        }

        public void AutoClose()
        {
        }
    }
}
