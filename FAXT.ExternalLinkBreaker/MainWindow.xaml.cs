using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Xl = Microsoft.Office.Interop.Excel;

namespace FAXT.ExternalLinkBreaker
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public bool isEmpty { get; private set; } = false;
        private Xl.Workbook openWb;
        public MainWindow(Xl.Application xlApp)
        {
            this.CommandBindings.Add(new CommandBinding(SystemCommands.CloseWindowCommand, this.OnCloseWindow));
            this.CommandBindings.Add(new CommandBinding(SystemCommands.MaximizeWindowCommand, this.OnMaximizeWindow, this.OnCanResizeWindow));
            this.CommandBindings.Add(new CommandBinding(SystemCommands.MinimizeWindowCommand, this.OnMinimizeWindow, this.OnCanMinimizeWindow));
            this.CommandBindings.Add(new CommandBinding(SystemCommands.RestoreWindowCommand, this.OnRestoreWindow, this.OnCanResizeWindow));
            InitializeComponent();
            openWb = xlApp.ActiveWorkbook;
            if ((Array)((object)openWb.LinkSources(Xl.XlLink.xlExcelLinks)) == null)
            {
                MessageBox.Show("Current workbook does not contain any external link.", "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
                isEmpty = true;
            }
            else
            {
                RefreshList();
            }
        }
        private void RefreshList()
        {
            ExternalLinkList.Items.Clear();
            if ((Array)((object)openWb.LinkSources(Xl.XlLink.xlExcelLinks)) != null)
            {
                foreach (object link in (Array)((object)openWb.LinkSources(Xl.XlLink.xlExcelLinks)))
                {
                    ExternalLinkList.Items.Add(link.ToString());
                }
            }
        }
        private void DeleteExternalLink(IEnumerable<object> links)
        {
            Xl.Names names = openWb.Names;
            openWb.Application.ScreenUpdating = false;
            openWb.Application.Calculation = Xl.XlCalculation.xlCalculationManual;
            foreach (object link in links)
            {
                //Delete all named range with the external link in ReferTo
                foreach (Xl.Name name in names)
                {
                    string formula = name.RefersTo;
                    if (formula.Replace(@"[", "").Replace(@"]", "").Contains(link.ToString()))
                    {
                        name.Delete();
                    }
                }
                //Delete all data validation with the external link in the formula
                foreach (Xl.Worksheet ws in openWb.Worksheets)
                {
                    if (SpecialCellsCatchError(ws.Cells, Xl.XlCellType.xlCellTypeAllValidation) != null)
                    {
                        foreach (Xl.Range cell in ws.Cells.SpecialCells(Xl.XlCellType.xlCellTypeAllValidation))
                        {
                            if (cell.Validation.Formula1.Replace(@"[", "").Replace(@"]", "").Contains(link.ToString()) || cell.Validation.Formula2.Replace(@"[", "").Replace(@"]", "").Contains(link.ToString()))
                            {
                                cell.Validation.Delete();
                            }
                        }
                    }
                    if (ws.Cells.FormatConditions.Count > 0)
                    {
                        for (int i = 1; i <= ws.Cells.FormatConditions.Count; i++)
                        {
                            Xl.FormatCondition formatCondition = ws.Cells.FormatConditions[i] as Xl.FormatCondition;
                            if (formatCondition != null)
                            {
                                if (formatCondition.Formula1.Replace(@"[", "").Replace(@"]", "").Contains(link.ToString()) || formatCondition.Formula2.Replace(@"[", "").Replace(@"]", "").Contains(link.ToString()))
                                {
                                    formatCondition.Delete();
                                }
                            }
                        }
                    }
                }
                openWb.BreakLink(link.ToString(), Xl.XlLinkType.xlLinkTypeExcelLinks);
            }
            openWb.Application.ScreenUpdating = false;
            openWb.Application.Calculation = Xl.XlCalculation.xlCalculationAutomatic;
        }
        private Xl.Range SpecialCellsCatchError(Xl.Range myRange, Xl.XlCellType cellType)
        {
            try
            {
                return myRange.SpecialCells(cellType);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                return null;
            }
        }
        private void btnDeleteAll(object sender, RoutedEventArgs e)
        {
            var links = ExternalLinkList.Items.Cast<object>();
            DeleteExternalLink(links);
            RefreshList();
        }
        private void btnDeleteSelected (object sender, RoutedEventArgs e)
        {
            var links = ExternalLinkList.SelectedItems.Cast<object>();
            DeleteExternalLink(links);
            RefreshList();
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
        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            ListGrid.Width = ExternalLinkBreaker.ActualWidth * 350 / 600;
            ButtonGrid.Width = ExternalLinkBreaker.ActualWidth * 230 / 600;
        }
    }
}