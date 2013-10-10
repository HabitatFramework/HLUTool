// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2013 Andy Foy
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
using System.Windows.Input;

namespace HLU.UI.View
{
    /// <summary>
    /// Interaction logic for MergeFeaturesWindow.xaml
    /// </summary>
    public partial class WindowMergeFeatures : Window
    {
        public WindowMergeFeatures()
        {
            InitializeComponent();
        }

        private void DataGridQueryBuilder_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (this.ButtonFlashFeature.Command.CanExecute(null))
                this.ButtonFlashFeature.Command.Execute(null);
        }
    }
}
