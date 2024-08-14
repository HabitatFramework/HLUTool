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
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace HLU.UI.UserControls
{
    public class SelectionList<T> : ObservableCollection<SelectionItem<T>> 
        where T : IComparable<T>
    {
        #region Properties

        /// <summary>
        /// Returns the selected items in the list
        /// </summary>
        public IEnumerable<T> SelectedItems
        {
            get { return this.Where(x => x.IsSelected).Select(x => x.Item); }
        }

        /// <summary>
        /// Returns all the items in the SelectionList
        /// </summary>
        public IEnumerable<T> AllItems
        {
            get { return this.Select(x => x.Item); }
        }

        #endregion

        #region Constructor

        public SelectionList(IEnumerable<T> col)
            : base(toSelectionItemEnumerable(col))
        {
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Adds the item to the list
        /// </summary>
        /// <param name="item"></param>
        public void Add(T item)
        {
            int i = 0;
            foreach (T existingItem in AllItems)
            {
                if (item.CompareTo(existingItem) < 0) break;
                i++;
            }
            Insert(i, new SelectionItem<T>(item));
        }

        /// <summary>
        /// Checks if the item exists in the list
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public bool Contains(T item)
        {
            return AllItems.Contains(item);
        }

        /// <summary>
        /// Selects all the items in the list
        /// </summary>
        public void SelectAll()
        {
            foreach (SelectionItem<T> selectionItem in this)
            {
                selectionItem.IsSelected = true;
            }
        }

        /// <summary>
        /// Unselects all the items in the list
        /// </summary>
        public void UnselectAll()
        {
            foreach (SelectionItem<T> selectionItem in this)
            {
                selectionItem.IsSelected = false;
            }
        }

        #endregion

        #region Helper methods

        /// <summary>
        /// Creates an SelectionList from any IEnumerable
        /// </summary>
        /// <param name="items"></param>
        /// <returns></returns>
        private static IEnumerable<SelectionItem<T>> toSelectionItemEnumerable(IEnumerable<T> items)
        {
            List<SelectionItem<T>> list = new List<SelectionItem<T>>();
            foreach (T item in items)
            {
                SelectionItem<T> selectionItem = new SelectionItem<T>(item);
                list.Add(selectionItem);
            }
            return list;
        }

        #endregion
    }
}
