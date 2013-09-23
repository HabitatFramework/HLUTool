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

namespace HLU
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class WindowMain : Window
    {
        private ComboBox[] _comboBoxes;

        public WindowMain()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            _comboBoxes = FindControls.FindLogicalChildren<ComboBox>(this.GridMain).ToArray(); 
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

                string validText = cb.Text.Substring(0, caretIx < 1 ? 0 : caretIx - 1);
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
}
