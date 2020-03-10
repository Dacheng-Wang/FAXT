using System;
using Microsoft.Office.Tools.Ribbon;
using Dh = ExcelHelper.DropdownHelper;
using System.Threading;
using System.Windows.Threading;
using System.IO;
using Xl = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Help = ExcelHelper.Help;
using Xml = ExcelHelper.XMLImporter;

namespace ExcelHelper.ExcelClient
{
    public partial class MainRibbon
    {
        private string _appSettingsPath = null;
        private Dh.MainWindow dhMainWindow;
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
            dhMainWindow = new Dh.MainWindow(Globals.ThisAddIn.Application);

            Globals.ThisAddIn.Application.ActiveSheet.SelectionChange += new Xl.DocEvents_SelectionChangeEventHandler(dhMainWindow.SelectionChange);
            if (dhMainWindow == null)
            {
                dhMainWindow.StartNewWindow(Globals.ThisAddIn.Application);
            }
            else
            {
                dhMainWindow.Close();
                dhMainWindow.StartNewWindow(Globals.ThisAddIn.Application);
            }
        }
        private void btnXMLImporter_Click(object sender, RibbonControlEventArgs e)
        {
            Xml.MainWindow xmlWindow = new Xml.MainWindow();
            xmlWindow.Show();
        }
        private void btnHelp_Click(object sender, RibbonControlEventArgs e)
        {
            Help.MainWindow helpWindow = new Help.MainWindow();
            helpWindow.Show();
        }
    }
}
