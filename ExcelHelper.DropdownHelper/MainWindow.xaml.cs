using System.Windows.Input;
using Xl = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Windows.Threading;
using System.Threading;
using System.Windows.Controls;

namespace ExcelHelper.DropdownHelper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Xl.Application activeApp;
        private Xl.Workbook activeWb;
        private Xl.Worksheet activeWs;
        private Xl.Range activeRange;
        private List<string> validationList;
        private MainWindow mainWindow;
        private bool isAuto = false;
        private bool previousValidation = false;
        private bool currentValidation = false;
        public MainWindow(Xl.Application excelApp)
        {
            InitializeComponent();
            activeApp = excelApp;
            RefreshActive();
            validationList = ReadDropDownValues(activeWb, activeRange);
            SearchBox.ItemsSource = validationList;
        }
        public void SelectionChange(Xl.Range Target)
        {
            RefreshActive();
            validationList = ReadDropDownValues(activeWb, activeRange);
            string formulaRange;
            
            try
            {
                previousValidation = currentValidation;
                formulaRange = activeApp.Selection.Validation.Formula1;
                currentValidation = true;
                currentValidation = true;
            }
            catch (COMException e)
            {
                currentValidation = false;
            }
            if (mainWindow != null)
            {
                AutoWindow(mainWindow);
            }
            else
            {
                AutoWindow(this);
            }
        }
        private void AutoWindow(MainWindow window)
        {
            window.Dispatcher.BeginInvoke(new Action(() =>
            {
                isAuto = (bool)window.AutoToggle.IsChecked;
                window.SearchBox.ItemsSource = validationList;
                if (isAuto)
                {
                    if (currentValidation && !previousValidation)
                    {
                        window.Close();
                    }
                    else if (!currentValidation)  window.Hide();
                }
            }), DispatcherPriority.Background);
            if (isAuto && currentValidation && !previousValidation) StartNewWindow(activeApp);
        }
        public void StartNewWindow(Xl.Application activeApp)
        {
            var thread = new Thread(() =>
            {
                mainWindow = new MainWindow(activeApp);
                mainWindow.Show();
                mainWindow.AutoToggle.IsChecked = isAuto;
                mainWindow.Topmost = true;
                mainWindow.Closed += (sender2, e2) => mainWindow.Dispatcher.InvokeShutdown();
                Dispatcher.Run();
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }
        void WindowActivated(object sender, EventArgs e)
        {
            RefreshActive();
            if (mainWindow == null)
            {
                this.SearchBox.ConvertNormalCBToAutoComplete();
                this.SearchBox.IsDropDownOpen = true;
            }
            else
            {
                mainWindow.SearchBox.ConvertNormalCBToAutoComplete();
                mainWindow.SearchBox.IsDropDownOpen = true;
            }
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Keyboard.Focus(this.SearchBox);
        }
        private void btnFill(object sender, RoutedEventArgs e)
        {
            RefreshActive();
            if (SearchBox.Items.Contains(SearchBox.Text))
            {
                activeRange.Value2 = SearchBox.Text;
                if (DirectionBox.SelectedItem != null)
                {
                    if (DirectionBox.Text == "Down") activeRange.Offset[1].Select();
                    if (DirectionBox.Text == "Up") activeRange.Offset[-1].Select();
                    if (DirectionBox.Text == "Right") activeRange.Offset[0, 1].Select();
                    if (DirectionBox.Text == "Left") activeRange.Offset[0, -1].Select();
                }
            }
            else MessageBox.Show("Invalid Input.", "Error");
            SearchBox.IsDropDownOpen = true;
            Keyboard.Focus(SearchBox);
        }
        private void RefreshActive()
        {
            activeWb = activeApp.ActiveWorkbook;
            activeWs = activeWb.ActiveSheet;
            activeRange = activeApp.Selection;
        }
        List<string> ReadDropDownValues(Xl.Workbook xlWorkBook, Xl.Range dropDownCell)
        {
            List<string> result = new List<string>();
            string formulaRange;
            //Test if cell has validation
            try
            {
                formulaRange = dropDownCell.Validation.Formula1;
                //Test if the validation is a list, a formula reference, or a named range reference
                Xl.Worksheet xlWorkSheet;
                string[] splitFormulaRange;
                Xl.Range valRange;
                if (formulaRange.Contains(","))
                {
                    result = formulaRange.Split(',').ToList();
                }
                else
                {
                    if (formulaRange.Contains(":"))
                    {
                        //test if there is external reference
                        if (formulaRange.Contains("!"))
                        {
                            string[] formulaRangeWorkSheetAndCells = formulaRange.Substring(1, formulaRange.Length - 1).Split('!');
                            splitFormulaRange = formulaRangeWorkSheetAndCells[1].Split(':');
                            xlWorkSheet = xlWorkBook.Worksheets.get_Item(formulaRangeWorkSheetAndCells[0]);
                        }
                        else
                        {

                            splitFormulaRange = formulaRange.Substring(1, formulaRange.Length - 1).Split(':');
                            xlWorkSheet = activeWs;
                        }
                        valRange = xlWorkSheet.get_Range(splitFormulaRange[0], splitFormulaRange[1]);
                    }
                    else
                    {
                        if (formulaRange.Contains("!"))
                        {
                            string[] formulaRangeWorkSheetAndCells = formulaRange.Substring(1, formulaRange.Length - 1).Split('!');
                            xlWorkSheet = xlWorkBook.Worksheets.get_Item(formulaRangeWorkSheetAndCells[0]);
                            valRange = xlWorkSheet.get_Range(formulaRangeWorkSheetAndCells[1]);
                        }
                        else
                        {
                            valRange = activeApp.get_Range(formulaRange.Substring(1, formulaRange.Length - 1));
                        }
                    }
                    for (int nRows = 1; nRows <= valRange.Rows.Count; nRows++)
                    {
                        for (int nCols = 1; nCols <= valRange.Columns.Count; nCols++)
                        {
                            Xl.Range aCell = (Xl.Range)valRange.Cells[nRows, nCols];
                            if (aCell.Value2 != null)
                            {
                                result.Add(aCell.Value2.ToString());
                            }
                        }
                    }
                }
            }
            catch (COMException e)
            {
            }
            return result;
        }
        private void btnClear(object sender, RoutedEventArgs e)
        {
            SearchBox.Text="";
        }
        private void btnSortAZ(object sender, RoutedEventArgs e)
        {
            RefreshActive();
            validationList = ReadDropDownValues(activeWb, activeRange);
            validationList.Sort();
            this.Dispatcher.BeginInvoke(new Action(() =>
            {
                SearchBox.ItemsSource = validationList;
            }), DispatcherPriority.Background);
        }
        private void btnSortZA(object sender, RoutedEventArgs e)
        {
            RefreshActive();
            validationList = ReadDropDownValues(activeWb, activeRange);
            validationList.Sort();
            validationList.Reverse();
            this.Dispatcher.BeginInvoke(new Action(() =>
            {
                SearchBox.ItemsSource = validationList;
            }), DispatcherPriority.Background);
        }
        private void SearchBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Equals(Key.Enter))
            {
                btnFill(sender, e);
            }
        }
    }
}
