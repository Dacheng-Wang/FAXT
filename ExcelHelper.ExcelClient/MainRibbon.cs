using System;
using Microsoft.Office.Tools.Ribbon;
using ExcelHelper.DropdownHelper;
using System.Threading;
using System.Windows.Threading;
using System.IO;
using Xl = Microsoft.Office.Interop.Excel;

namespace ExcelHelper.ExcelClient
{
    public partial class MainRibbon
    {
        private string _appSettingsPath = null;
        private MainWindow mainWindow;
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            _appSettingsPath = Directory.CreateDirectory(
                 Path.Combine(
                     Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                     "ProjectCVIA"
                     )
                 ).FullName;
        }
        private void btnDropdownHelper_Click(object sender, RibbonControlEventArgs e)
        {
            if (mainWindow == null)
            {
                var thread = new Thread(() =>
                {
                    mainWindow = new MainWindow(Globals.ThisAddIn.Application);
                    mainWindow.Show();
                    mainWindow.Topmost = true;
                    mainWindow.Closed += (sender2, e2) => mainWindow.Dispatcher.InvokeShutdown();

                    Dispatcher.Run();
                });

                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
            }
            else
            {
                mainWindow.Dispatcher.BeginInvoke((Action)mainWindow.Close);
                var thread = new Thread(() =>
                {
                    mainWindow = new MainWindow(Globals.ThisAddIn.Application);
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
}
