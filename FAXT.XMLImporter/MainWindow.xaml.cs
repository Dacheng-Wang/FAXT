using System.Windows;
using System.Xml.Linq;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Xml.Xsl;
using System;
using System.Windows.Input;
using Xl = Microsoft.Office.Interop.Excel;
using System.Xml;

namespace FAXT.XMLImporter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Xl.Application app = new Xl.Application();
        public MainWindow(Xl.Application xlApp, String _appPath)
        {
            app = xlApp;
            this.CommandBindings.Add(new CommandBinding(SystemCommands.CloseWindowCommand, this.OnCloseWindow));
            this.CommandBindings.Add(new CommandBinding(SystemCommands.MaximizeWindowCommand, this.OnMaximizeWindow, this.OnCanResizeWindow));
            this.CommandBindings.Add(new CommandBinding(SystemCommands.MinimizeWindowCommand, this.OnMinimizeWindow, this.OnCanMinimizeWindow));
            this.CommandBindings.Add(new CommandBinding(SystemCommands.RestoreWindowCommand, this.OnRestoreWindow, this.OnCanResizeWindow));
            InitializeComponent();
            using (OpenFileDialog fileDialog = new OpenFileDialog())
            {
                fileDialog.Multiselect = false;
                fileDialog.Filter = "XML Files (*.xml)|*.xml";
                fileDialog.CheckFileExists = true;
                fileDialog.Title = "Please Select Your XML File";
                if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var myXslTrans = new XslCompiledTransform();
                    var reader = new StringReader(Properties.Resources.Generic);
                    XmlReaderSettings settings = new XmlReaderSettings();
                    settings.DtdProcessing = DtdProcessing.Parse;
                    XmlReader xsltReader = XmlReader.Create(reader, settings);
                    myXslTrans.Load(xsltReader);
                    XmlReader xmlReader = new XmlTextReader(fileDialog.FileName);
                    StreamWriter inputXML = new StreamWriter(Path.Combine(_appPath, "result.html"));
                    myXslTrans.Transform(xmlReader, null, inputXML);
                    inputXML.Close();
                    inputXML.Dispose();
                    FileStream htmlFile = new FileStream(Path.Combine(_appPath, "result.html"), FileMode.Open);
                    xmlViewer.NavigateToStream(htmlFile);
                    xMLWindow.Topmost = true;
                }
            }
        }
        private void FillBlankInRange(Xl.Range range)
        {
            DataTable table = new DataTable();
            foreach (Xl.Range rng in range)
            {
                if (rng.Interior.Color == 16777215 && rng.Value2 == null)
                {
                    if (rng.Offset[-1].Value2 == null) rng.Value2 = rng.End[Xl.XlDirection.xlUp].Value2;
                    else rng.Value2 = rng.Offset[-1].Value2;
                }
            }
        }
        private void btnImportMerged(object sender, RoutedEventArgs e)
        {
            dynamic doc = xmlViewer.Document;
            doc.ExecCommand("SelectAll", true, null);
            doc.ExecCommand("Copy", false, null);
            app.Selection.PasteSpecial();
            doc.Selection.Empty();
        }
        private void btnImportUnmerged(object sender, RoutedEventArgs e)
        {
            btnImportMerged(sender, e);
            app.Selection.UnMerge();
            FillBlankInRange(app.Selection);
        }
        private void OnCanResizeWindow(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = this.ResizeMode == ResizeMode.CanResize || this.ResizeMode == ResizeMode.CanResizeWithGrip;
        }

        private void OnCanMinimizeWindow(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = this.ResizeMode != ResizeMode.NoResize;
        }

        private void OnCloseWindow(object target, ExecutedRoutedEventArgs e)
        {
            SystemCommands.CloseWindow(this);
        }

        private void OnMaximizeWindow(object target, ExecutedRoutedEventArgs e)
        {
            SystemCommands.MaximizeWindow(this);
        }

        private void OnMinimizeWindow(object target, ExecutedRoutedEventArgs e)
        {
            SystemCommands.MinimizeWindow(this);
        }

        private void OnRestoreWindow(object target, ExecutedRoutedEventArgs e)
        {
            SystemCommands.RestoreWindow(this);
        }

    }
}