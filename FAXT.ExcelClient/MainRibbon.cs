using System;
using Microsoft.Office.Tools.Ribbon;
using Dh = FAXT.DropdownHelper;
using System.Threading;
using System.Windows.Threading;
using System.IO;
using Xl = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Help = FAXT.Help;
using Xml = FAXT.XMLImporter;
using System.Diagnostics;

namespace FAXT.ExcelClient
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
                     "FAXT"
                     )
                 ).FullName;
        }
        private void DropdownHelper_Click(object sender, RibbonControlEventArgs e)
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
        private void XMLImporter_Click(object sender, RibbonControlEventArgs e)
        {
            Xml.MainWindow xmlWindow = new Xml.MainWindow(Globals.ThisAddIn.Application, _appSettingsPath);
            xmlWindow.Show();
        }
        private void Help_Click(object sender, RibbonControlEventArgs e)
        {
            Help.MainWindow helpWindow = new Help.MainWindow();
            helpWindow.Show();
        }
        private void Tabula_Click(object sender, RibbonControlEventArgs e)
        {
            Directory.CreateDirectory(Path.Combine(_appSettingsPath, "Tabula"));
            try
            {
                if (File.Exists(Path.Combine(_appSettingsPath, "Tabula", "Tabula.exe"))) File.Delete(Path.Combine(_appSettingsPath, "Tabula", "Tabula.exe"));
                if (File.Exists(Path.Combine(_appSettingsPath, "Tabula", "Tabula.jar"))) File.Delete(Path.Combine(_appSettingsPath, "Tabula", "Tabula.jar"));
                string tempExeName = Path.Combine(_appSettingsPath, "Tabula", "Tabula.exe");
                using (FileStream fsDst = new FileStream(tempExeName, FileMode.CreateNew, FileAccess.Write))
                {
                    byte[] bytes = Resource.tabula_exe;

                    fsDst.Write(bytes, 0, bytes.Length);
                }
                string tempJarName = Path.Combine(_appSettingsPath, "Tabula", "Tabula.jar");
                using (FileStream fsDst = new FileStream(tempJarName, FileMode.CreateNew, FileAccess.Write))
                {
                    byte[] bytes = Resource.tabula_jar;

                    fsDst.Write(bytes, 0, bytes.Length);
                }
                var process = new Process();
                process.StartInfo.FileName = Path.Combine(_appSettingsPath, "Tabula", "Tabula.exe");
                process.StartInfo.WorkingDirectory = Path.Combine(_appSettingsPath, "Tabula");
                process.StartInfo.CreateNoWindow = true;
                process.Start();
                Thread.Sleep(10000);
                DialogResult dialogResult = MessageBox.Show("Please do not close command window while Tabula is running. Do you want to open Tabula?", "Starting Tabula", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (dialogResult == DialogResult.Yes)
                {
                    Process.Start("http://127.0.0.1:8080/");
                }
                else if (dialogResult == DialogResult.No)
                {
                    process.Close();
                }
            }
            catch
            {
                DialogResult dialogResult1 = MessageBox.Show("Tabula.exe is currently running. Click OK to start the web app", "Tabula", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (dialogResult1 == DialogResult.OK) Process.Start("http://127.0.0.1:8080/");
            }
        }
    }
}
