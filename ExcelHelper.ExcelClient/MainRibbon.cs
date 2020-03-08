using System;
using Microsoft.Office.Tools.Ribbon;
using ExcelHelper.DropdownHelper;
using System.Threading;
using System.Windows.Threading;
using System.IO;
using Xl = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

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
            mainWindow = new MainWindow(Globals.ThisAddIn.Application);

            Globals.ThisAddIn.Application.ActiveSheet.SelectionChange += new Xl.DocEvents_SelectionChangeEventHandler(mainWindow.SelectionChange);
            if (mainWindow == null)
            {
                mainWindow.StartNewWindow(Globals.ThisAddIn.Application);
            }
            else
            {
                mainWindow.Close();
                mainWindow.StartNewWindow(Globals.ThisAddIn.Application);
            }
        }
    }
}
