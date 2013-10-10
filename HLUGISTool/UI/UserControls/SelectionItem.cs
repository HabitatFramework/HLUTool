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

using System;
using System.ComponentModel;

namespace HLU.UI.UserControls
{
    public class SelectionItem<T> : INotifyPropertyChanged
    {
        #region Fields

        private bool isSelected;

        private T item;

        #endregion

        #region Properties

        public bool IsSelected
        {
            get { return isSelected; }
            set
            {
                if (value == isSelected) return;
                isSelected = value;
                OnPropertyChanged("IsSelected");
                OnSelectionChanged();
            }
        }

        public T Item
        {
            get { return item; }
            set
            {
                if (value.Equals(item)) return;
                item = value;
                OnPropertyChanged("Item");
            }
        }

        #endregion

        #region Events

        public event PropertyChangedEventHandler PropertyChanged;

        public event EventHandler SelectionChanged;

        #endregion

        #region Constructor

        public SelectionItem(T item)
            : this(false, item)
        {
        }

        public SelectionItem(bool selected, T item)
        {
            this.isSelected = selected;
            this.item = item;
        }

        #endregion

        #region Event invokers

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler changed = PropertyChanged;
            if (changed != null) changed(this, new PropertyChangedEventArgs(propertyName));
        }

        private void OnSelectionChanged()
        {
            EventHandler changed = SelectionChanged;
            if (changed != null) changed(this, EventArgs.Empty);
        }

        #endregion
    }
}
