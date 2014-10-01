// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// 
// This file is part of HLUTool.
// 
// HLUTool is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// HLUTool is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with HLUTool.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using HLU.UI.UserControls;
using Microsoft.Windows.Controls;
using HLU.Properties;

namespace HLU
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class WindowMain : Window
    {
        private ComboBox[] _comboBoxes;
        private MenuItem[] _menuItems;
        public string lastStyle = null;

        public WindowMain()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            _comboBoxes = FindControls.FindLogicalChildren<ComboBox>(this.GridMain).ToArray();

            //---------------------------------------------------------------------
            // FIXED: KI15 (User Interface Style)
            // Create
            //
            // Create an array of all the menu items in the window.
            _menuItems = FindControls.FindLogicalChildren<MenuItem>(this.MenuBar).ToArray();

            // Set the last style to be the default style (already loaded).
            lastStyle = Settings.Default.InterfaceStyle;

            // Check the menu item for the default style.
            CheckMenuItem(lastStyle, true);
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            IInputElement focusedElement = Keyboard.FocusedElement;
            Control focusedControl = focusedElement as Control;
            if ((focusedControl != null) && (focusedControl.Parent is DataGridCell))
                return;

            foreach (ComboBox cbx in _comboBoxes)
                if (cbx.IsDropDownOpen) return;

            switch (e.Key)
            {
                case Key.Return:
                    if (this.TextBoxRecordNumber.IsFocused)
                    {
                        this.ComboBoxProcess.Focus();
                        this.TextBoxRecordNumber.Focus();
                    }
                    break;
                case Key.Home:
                    if ((focusedControl != null) && (focusedControl is TextBox)) return;
                    if (this.ButtonFirstRecord.Command.CanExecute(null))
                        this.ButtonFirstRecord.Command.Execute(null);
                    break;
                case Key.Prior:
                    if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
                        goto case Key.Prior;
                    else if (this.ButtonPreviousRecord.Command.CanExecute(null))
                        this.ButtonPreviousRecord.Command.Execute(null);
                    break;
                case Key.Next:
                    if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
                        goto case Key.End;
                    else if (this.ButtonNextRecord.Command.CanExecute(null))
                        this.ButtonNextRecord.Command.Execute(null);
                    break;
                case Key.End:
                    if ((focusedControl != null) && (focusedControl is TextBox)) return;
                    if (this.ButtonLastRecord.Command.CanExecute(null))
                        this.ButtonLastRecord.Command.Execute(null);
                    break;
            }
        }

        private void LabelStatusIncid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (this.MenuItemClearFilter.Command.CanExecute(null))
                this.MenuItemClearFilter.Command.Execute(null);
        }

        /// <summary>
        /// Since the IHS multiplex ComboBox controls are set to IsEditable="True" and IsReadOnly="True" 
        /// (i.e., restricted to list), their text does not get cleared when a repeated value has been 
        /// chosen and rejected by the view model.
        /// This event handler uses reflection to get the underlying view model property value and sets the
        /// ComboBox text to String.Empty if the view model property is null.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ComboBox_DropDownClosed(object sender, EventArgs e)
        {
            string s = ((ComboBox)sender).Text;
            ComboBox cmb = sender as ComboBox;
            if (cmb == null) return;

            object propValue = cmb.DataContext.GetType().GetProperty(cmb.GetBindingExpression(
                ComboBox.SelectedValueProperty).ParentBinding.Path.Path).GetValue(cmb.DataContext, null);

            string propStr = propValue != null ? propValue.ToString() : String.Empty;

            if (String.IsNullOrEmpty(propStr))
            {
                if (!String.IsNullOrEmpty(cmb.Text))
                    cmb.Text = String.Empty;
            }
            else
            {
                var q = cmb.Items.Cast<object>().Where(o => o.GetType().GetProperty(
                    cmb.SelectedValuePath).GetValue(o, null).Equals(cmb.SelectedValue));

                if (q.Count() == 1)
                {
                    PropertyInfo pi = q.ElementAt(0).GetType().GetProperty(cmb.DisplayMemberPath);
                    if (pi != null)
                    {
                        object displayValue = pi.GetValue(q.ElementAt(0), null);
                        string displayText = displayValue as string;
                        if (!String.IsNullOrEmpty(displayText) && (displayText != cmb.Text))
                            cmb.Text = displayText;
                    }
                }
            }
        }

        private void ButtonGetMapSelection2_MouseEnter(object sender, MouseEventArgs e)
        {
            if (this.ButtonGetMapSelection2.IsEnabled)
            {
                this.ButtonGetMapSelection2.Width *= 1.1;
                this.ButtonGetMapSelection2.Height *= 1.1;
            }
        }

        private void ButtonGetMapSelection2_MouseLeave(object sender, MouseEventArgs e)
        {
            if (this.ButtonGetMapSelection2.IsEnabled)
            {
                this.ButtonGetMapSelection2.Width = this.ButtonGetMapSelection2.Width / 11 * 10;
                this.ButtonGetMapSelection2.Height = this.ButtonGetMapSelection2.Width / 11 * 10;
            }
        }

        private void EditableComboBox_KeyUp(object sender, KeyEventArgs e)
        {
            ValidateComboboxText(sender as ComboBox);
        }

        private void ValidateComboboxText(ComboBox cb)
        {
            if ((cb == null) || (cb.Items.Count == 0)) return;

            PropertyInfo pi = cb.Items[0].GetType().GetProperty(cb.DisplayMemberPath);

            for (int i = 0; i < cb.Items.Count; i++)
            {
                if (pi.GetValue(cb.Items[i], null).ToString().Equals(cb.Text))
                    return;
            }

            if (cb.SelectedIndex != -1)
            {
                cb.Text = pi.GetValue(cb.SelectedItem, null).ToString();
            }
            else
            {
                TextBox tbx = (TextBox)cb.Template.FindName("PART_EditableTextBox", cb);
                int caretIx = tbx.CaretIndex;

                //---------------------------------------------------------------------
                // FIX: 031 Check combobox text is not null before finding list item
                if (cb.Text != null)
                //---------------------------------------------------------------------
                {
                    string validText = cb.Text.Substring(0, caretIx < 1 ? 0 : caretIx);
                    //string validText = cb.Text.Substring(0, caretIx < 1 ? 0 : caretIx - 1);
                    for (int i = 0; i < cb.Items.Count; i++)
                    {
                        if (pi.GetValue(cb.Items[i], null).ToString().StartsWith(validText))
                        {
                            cb.SelectedIndex = i;
                            tbx.CaretIndex = caretIx;
                            return;
                        }
                    }
                    cb.Text = null;
                }
            }
        }

        //---------------------------------------------------------------------
        // FIXED: KI15 (User Interface Style)
        // Switch the interface style to the style selected by the user
        // based on the menu item name.
        //
        /// <summary>
        /// Handles the Click event of the MenuItem_Style control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void MenuItem_Style_Click(object sender, RoutedEventArgs e)
        {
            // Set the style name from the menu item name.
            MenuItem mi = sender as MenuItem;
            string styleName = string.Format("{0}", mi.Name);

            // Switch the style to the selected menu item style.
            if (App.LoadStyleDictionaryFromFile(styleName))
            {
                // Clear the check against the last menu item style.
                CheckMenuItem(lastStyle, false);

                // Store the last style as the current menu item.
                lastStyle = mi.Name;
            }
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // FIXED: KI15 (User Interface Style)
        // Check or uncheck a named menu item.
        //
        /// <summary>
        /// Checks or unchecks a named menu item.
        /// </summary>
        /// <param name="mItemName">The name of the menu item.</param>
        /// <param name="check">If set to <c>true</c> [check] the menu item is checked.</param>
        private void CheckMenuItem(string mItemName, bool check)
        {
            foreach (MenuItem mItem in _menuItems)
            {
                if ((mItem is MenuItem) && (mItem.Name == mItemName))
                {
                    if (check)
                        mItem.IsChecked = true;
                    else
                        mItem.IsChecked = false;
                }
            }
        }
        //---------------------------------------------------------------------
    }
}
