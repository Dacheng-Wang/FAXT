using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Diagnostics;

namespace ExcelHelper.DropdownHelper
{
    public static class AutoComplete
    {
        public static void ConvertNormalCBToAutoComplete(this ComboBox targetComboBox)
        {
            targetComboBox = TransformBox(targetComboBox);
        }
        private static ComboBox TransformBox(ComboBox box)
        {
            var targetComboBox = box as ComboBox;
            var targetTextBox = box?.Template.FindName("PART_EditableTextBox", box) as TextBox;
            if (targetTextBox == null) return null;

            targetComboBox.Tag = "TextInput";
            targetComboBox.StaysOpenOnEdit = true;
            targetComboBox.IsEditable = true;
            targetComboBox.IsTextSearchEnabled = false;

            targetTextBox.TextChanged += (o, args) =>
            {
                var textBox = (TextBox)o;
                var searchText = textBox.Text;

                if (targetComboBox.Tag.ToString() == "Selection")
                {
                    targetComboBox.Tag = "TextInput";
                    targetComboBox.IsDropDownOpen = true;
                }
                else
                {
                    if (targetComboBox.SelectionBoxItem != null)
                    {
                        targetComboBox.SelectedItem = null;
                        targetTextBox.Text = searchText;
                        textBox.CaretIndex = int.MaxValue;
                    }

                    if (string.IsNullOrEmpty(searchText))
                    {
                        targetComboBox.Items.Filter = item => true;
                        targetComboBox.SelectedItem = default(object);
                    }
                    else
                        targetComboBox.Items.Filter = item =>
                                CultureInfo.InvariantCulture.CompareInfo.IndexOf(item.ToString(), searchText, CompareOptions.IgnoreCase) >= 0;
                                //Back up code here in case we want to add a case sensitivity toggle
                                //item.ToString().Contains(searchText);
                                //Back up code here if we want to build a StartWith/Contains toggle
                                //item.ToString().StartsWith(searchText, true, CultureInfo.InvariantCulture);

                    Keyboard.ClearFocus();
                    Keyboard.Focus(targetTextBox);
                    targetComboBox.IsDropDownOpen = true;
                    targetTextBox.SelectionStart = targetTextBox.Text.Length;
                }
            };


            targetComboBox.SelectionChanged += (o, args) =>
            {
                var comboBox = o as ComboBox;
                if (comboBox?.SelectedItem == null) return;
                comboBox.Tag = "Selection";
            };
            return targetComboBox;
        }
    }
}
