using System.Linq;
using System.Windows;
using Crl = System.Windows.Controls;
using System.Xml;
using System.Xml.Linq;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Windows.Media;
using System.Collections.Generic;
using System.Xml.Xsl;
using System;
using System.Windows.Input;

namespace ExcelHelper.XMLImporter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string projectPath = Directory.GetParent(Environment.CurrentDirectory).Parent.FullName;
        public MainWindow()
        {
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
                    StreamReader file = File.OpenText(fileDialog.FileName);
                    XDocument xmlDoc = XDocument.Load(file);
                    var myXslTrans = new XslCompiledTransform();
                    myXslTrans.Load(Path.Combine(projectPath, "Generic.xslt"));
                    myXslTrans.Transform(fileDialog.FileName, Path.Combine(projectPath, "result.html"));
                    FileStream htmlFile = new FileStream(Path.Combine(projectPath, "result.html"), FileMode.Open);
                    xmlViewer.NavigateToStream(htmlFile);
                    xMLWindow.Topmost = true;
                }
            }
        }
        
        private DataTable ReturnDataTableFromNode(Crl.TreeViewItem treeViewItem)
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Child Node");
            foreach (Crl.TreeViewItem firstLayerItem in treeViewItem.Items)
            {
                DataRow dataRow = dataTable.NewRow();
                if (firstLayerItem.HasItems)
                {
                    foreach (Crl.TreeViewItem childItem in firstLayerItem.Items)
                    {
                        if (!dataTable.Columns.Contains(childItem.Header.ToString())) dataTable.Columns.Add(childItem.Header.ToString());
                        foreach (Crl.TreeViewItem grandchildItem in childItem.Items)
                        {
                            dataRow[childItem.Header.ToString()] = grandchildItem.Header;
                        }
                    }
                    dataRow["Child Node"] = firstLayerItem.Header;
                    dataTable.Rows.Add(dataRow);
                }
                else return null;
            }
            return dataTable;
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