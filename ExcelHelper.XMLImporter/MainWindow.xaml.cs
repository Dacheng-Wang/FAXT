using System.Linq;
using System.Windows;
using Crl = System.Windows.Controls;
using System.Xml;
using System.Xml.Linq;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Xml;
using System.Windows.Media;

namespace ExcelHelper.XMLImporter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
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
                    BuildTree(treeView, xmlDoc);

                    //MemoryStream xmlStream = new MemoryStream();
                    //xmlDoc.Save(xmlStream);
                    //xmlStream.Position = 0;
                    //DataTable newTable = new DataTable();
                    //newTable.ReadXml(xmlStream);

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
            dataGrid.ItemsSource = dataTable.DefaultView;
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
                }
                else return null;
                dataTable.Rows.Add(dataRow);
            }
            return dataTable;
        }

    }
}
