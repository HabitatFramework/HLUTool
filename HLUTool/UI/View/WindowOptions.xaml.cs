// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// Copyright © 2019 London & South East Record Centres (LaSER)
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
using HLU.UI.UserControls;
using System.Windows.Controls;
using System.Collections.Generic;

namespace HLU.UI.View
{
    /// <summary>
    /// Interaction logic for WindowOptions.xaml
    /// </summary>
    public partial class WindowOptions : Window
    {
        public WindowOptions()
        {
            InitializeComponent();
        }

        // Initialise selected history column list items when History tab has focus.
        //
        /// <summary>
        /// Fakes a click on each history column list item so that their checkboxes are properly initialised.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TabHistory_GotFocus(object sender, RoutedEventArgs e)
        {
            List<DependencyObject> listBoxes = new List<DependencyObject>();
            FindControls.GetChildren(this, typeof(ListBox), ref listBoxes);

            foreach (ListBox l in listBoxes)
            {
                if ((l.Items.Count > 0) &&
                    (l.Items[0].GetType().ToString().StartsWith(this.GetType().Namespace.Replace("View", "UserControls") + ".SelectionItem")))
                    ForceIsSelectedUpdate(l);
            }
        }

        private void ForceIsSelectedUpdate(ListBox listBoxControl)
        {
            for (int i = 0; i < listBoxControl.Items.Count; i++)
            {
                SelectionItem<string> it = listBoxControl.Items[i] as SelectionItem<string>;
                if (it.IsSelected)
                {
                    it.IsSelected = !it.IsSelected;
                    it.IsSelected = !it.IsSelected;
                    listBoxControl.SelectedItems.Add(listBoxControl.Items[i]);
                }
            }
        }
    }
}
