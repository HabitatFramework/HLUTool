// HLUTool is used to view and maintain habitat and land use GIS data.
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
using System.Reflection;
using System.Windows.Controls;
//using System.Windows.Data;
//using System.Collections;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Text;
//using Microsoft.Windows.Controls;
using HLU.UI.ViewModel;

namespace HLU.UI.View
{
    //---------------------------------------------------------------------
    // CHANGED: CR49 Process proposed OSMM Updates
    // Functionality to process proposed OSMM Updates.
    //    
    /// <summary>
    /// Interaction logic for WindowQueryOSMM.xaml
    /// </summary>
    public partial class WindowQueryOSMM : Window
    {
        public WindowQueryOSMM()
        {
            InitializeComponent();

            //DataGridOSMMUpdatesSummary.Sorting += new DataGridSortingEventHandler(SortHandler);
        }

        void OSMMUpdates_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender != null)
            {
                ViewModelWindowQueryOSMM _viewModel = (ViewModelWindowQueryOSMM)this.DataContext;

                OSMMUpdates selectedRow = (OSMMUpdates)DataGridOSMMUpdatesSummary.SelectedItem;

                _viewModel.OSMMUpdatesSelectedRow(selectedRow);
            }
        }

        //void SortHandler(object sender, DataGridSortingEventArgs e)
        //{
        //    DataGridColumn column = e.Column;

        //    IComparer comparer = null;

        //    //i do some custom checking based on column to get the right comparer
        //    //i have different comparers for different columns. I also handle the sort direction
        //    //in my comparer

        //    // prevent the built-in sort from sorting
        //    e.Handled = true;

        //    ListSortDirection direction = (column.SortDirection != ListSortDirection.Ascending) ? ListSortDirection.Ascending : ListSortDirection.Descending;

        //    //set the sort order on the column
        //    column.SortDirection = direction;

        //    //use a ListCollectionView to do the sort.
        //    ListCollectionView lcv = (ListCollectionView)CollectionViewSource.GetDefaultView(this.DataGridOSMMUpdatesSummary);

        //    //this is my custom sorter it just derives from IComparer and has a few properties
        //    //you could just apply the comparer but i needed to do a few extra bits and pieces
        //    comparer = new ResultSort(direction);

        //    //apply the sort
        //    lcv.CustomSort = comparer;
        //}    
    
    }
    //---------------------------------------------------------------------

    //public class ResultSort : IComparer
    //{
    //    ListSortDirection _direction;

    //    public ResultSort(ListSortDirection direction)
    //    {
    //        _direction = direction;
    //    }
    //    public int Compare(object x, object y)
    //    {
    //        //string studentIdX = (x as StudentSearchResult).StudentID.Value;
    //        //string studentIdY = (y as StudentSearchResult).StudentID.Value;
    //        string idX = (x as string);
    //        string idY = (x as string);

    //        if (_direction == ListSortDirection.Ascending)
    //        {
    //            return idX.SortableStringValue().CompareTo(idY.SortableStringValue());
    //        }
    //        else
    //        {
    //            return idY.SortableStringValue().CompareTo(idX.SortableStringValue());
    //        }
    //    }

    //}
    
    //public static class Extensions
    //{
    //    public static string SortableStringValue(this string text)
    //    {
    //        StringBuilder textBuilder;
    //        StringBuilder numberBuilder;
    //        textBuilder = new StringBuilder();
    //        numberBuilder = new StringBuilder();
    //        //Look at each char in the string
    //        foreach (char value in text)
    //        {
    //            switch (value)
    //            {
    //                //If its a number add it to the number builder
    //                case '0':
    //                case '1':
    //                case '2':
    //                case '3':
    //                case '4':
    //                case '5':
    //                case '6':
    //                case '7':
    //                case '8':
    //                case '9':
    //                    numberBuilder.Append(value);
    //                    break;
    //                //Else add it to the text builder
    //                default:
    //                    //Before we add the text, format and add any number we may have
    //                    if (numberBuilder.Length > 0)
    //                    {
    //                        textBuilder.Append(numberBuilder.ToString().PadLeft(16, '0'));
    //                        numberBuilder.Length = 0;
    //                        numberBuilder.Capacity = 0;
    //                    }
    //                    textBuilder.Append(value);

    //                    break;
    //            }
    //        }

    //        //Check to see if we have any numbers left in the builder
    //        if (numberBuilder.Length > 0)
    //        {
    //            textBuilder.Append(numberBuilder.ToString().PadLeft(16, '0'));
    //        }
    //        //Return the string value (The replace will allow negative number to be sorted)
    //        return textBuilder.ToString().Replace("-0", "-00");
    //    }

    //}

}
