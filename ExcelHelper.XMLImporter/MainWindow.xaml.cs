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
                    //string[,] data = LoadXMLToDataTable(xmlDoc);
                    var myXslTrans = new XslCompiledTransform();
                    myXslTrans.Load(Path.Combine(projectPath, "Generic.xslt"));
                    myXslTrans.Transform(fileDialog.FileName, Path.Combine(projectPath, "result.html"));
                    BuildTree(treeView, xmlDoc);
                    xMLWindow.Topmost = true;
                    dataGrid.FrozenColumnCount = 1;

                    //DataSet ds = new DataSet();
                    //ds.ReadXml(fileDialog.FileName);
                    //dataGrid.ItemsSource = ds.Tables["CATALOG"].DefaultView;
                }
            }
        }
        private void TreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            e.Handled = true;
            Crl.TreeViewItem selectedItem = (Crl.TreeViewItem)treeView.SelectedItem;
            //XmlNodeReader nodeReader = new XmlNodeReader();
            //Crl.ItemsControl parentItem = GetSelectedTreeViewItemParent(selectedItem);
            DataTable dataTable = ReturnDataTableFromNode(selectedItem);
            if (dataTable != null) dataGrid.ItemsSource = dataTable.DefaultView;
        }
        private void BuildTree(Crl.TreeView treeView, XDocument doc)
        {
            Crl.TreeViewItem treeNode = new Crl.TreeViewItem {Header = doc.Root.Name.LocalName};
            treeView.Items.Add(treeNode);
            BuildNodes(treeNode, doc.Root);
        }
        private void BuildNodes(Crl.TreeViewItem treeNode, XElement element)
        {
            foreach (XNode child in element.Nodes())
            {
                switch (child.NodeType)
                {
                    case XmlNodeType.Element:
                        XElement childElement = child as XElement;
                        Crl.TreeViewItem childTreeNode = new Crl.TreeViewItem {Header = childElement.Name.LocalName};
                        treeNode.Items.Add(childTreeNode);
                        BuildNodes(childTreeNode, childElement);
                        break;
                    case XmlNodeType.Text:
                        XText childText = child as XText;
                        treeNode.Items.Add(new Crl.TreeViewItem {Header = childText.Value});
                        break;
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
        private void Window_Resize(object sender, SizeChangedEventArgs e)
        {
            e.Handled = true;
            treeGrid.Width = xMLWindow.ActualWidth / 4;
            previewGrid.Width = xMLWindow.ActualWidth * 3 / 4;
        }
        private string[,] LoadXMLToDataTable(XDocument xDocument)
        {
            List<string> vs = new List<string>();
            XElement xElement = xDocument.Root;
            do
            {
                vs.Add(xElement.Name.ToString());
                xElement = xElement.Descendants().ElementAt(0);
            } while (xElement.HasElements);
            return null;
        }
    }
}
