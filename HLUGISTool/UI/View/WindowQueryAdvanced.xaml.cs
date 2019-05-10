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

using System.Windows;
using System.Reflection;
using HLU.UI.UserControls;
using System.Windows.Controls;
using System.Windows.Input;
using System.Collections.Generic;

namespace HLU.UI.View
{
    /// <summary>
    /// Interaction logic for WindowQueryAdvanced.xaml
    /// </summary>
    public partial class WindowQueryAdvanced : Window
    {
        public WindowQueryAdvanced()
        {
            InitializeComponent();
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

    }

}
