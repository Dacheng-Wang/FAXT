using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using ExcelHelper.DropdownHelper;
using System.Threading;
using System.Windows.Threading;

namespace ExcelHelper.ExcelClient
{
    public partial class MainRibbon
    {
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }
        private void btnDropdownHelper_Click(object sender, RibbonControlEventArgs e)
        {
            var thread = new Thread(() =>
            {
                MainWindow mainWindow = new MainWindow();
                mainWindow.Show();
                mainWindow.Topmost = true;
                mainWindow.Closed += (sender2, e2) => mainWindow.Dispatcher.InvokeShutdown();

                Dispatcher.Run();
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }
    }
}
