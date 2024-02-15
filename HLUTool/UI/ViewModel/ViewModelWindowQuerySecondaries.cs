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

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;
using System.Text.RegularExpressions;
using HLU.Data;
using HLU.Data.Model;

namespace HLU.UI.ViewModel
{
    class ViewModelWindowQuerySecondaries : ViewModelBase, IDataErrorInfo
    {
        #region Fields

        private ICommand _okCommand;
        private ICommand _cancelCommand;
        private string _displayName = "Add Secondary Habitats";
        private String _querySecondaries;

        #endregion

        #region ViewModelBase Members

        public override string DisplayName
        {
            get { return _displayName; }
            set { _displayName = value; }
        }

        public override string WindowTitle
        {
            get { return DisplayName; }
        }

        #endregion

        #region RequestClose

        // declare the delegate since using non-generic pattern
        public delegate void RequestCloseEventHandler(String querySecondaries);

        // declare the event
        public event RequestCloseEventHandler RequestClose;

        #endregion

        #region Ok Command

        /// <summary>
        /// Create Ok button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand OkCommand
        {
            get
            {
                if (_okCommand == null)
                {
                    Action<object> okAction = new Action<object>(this.OkCommandClick);
                    _okCommand = new RelayCommand(okAction, param => this.CanOk);
                }

                return _okCommand;
            }
        }

        /// <summary>
        /// Handles event when Ok button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void OkCommandClick(object param)
        {
            this.RequestClose(QuerySecondaries);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanOk
        { 
            get 
            {
                return (String.IsNullOrEmpty(Error) && (_querySecondaries != null));
            } 
        }

        #endregion

        #region Cancel Command

        /// <summary>
        /// Create Cancel button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand CancelCommand
        {
            get
            {
                if (_cancelCommand == null)
                {
                    Action<object> cancelAction = new Action<object>(this.CancelCommandClick);
                    _cancelCommand = new RelayCommand(cancelAction);
                }
                return _cancelCommand;
            }
        }

        /// <summary>
        /// Handles event when Cancel button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void CancelCommandClick(object param)
        {
            this.RequestClose(null);
        }

        #endregion

        #region Query Secondaries

        public string QuerySecondaries
        {
            get { return _querySecondaries; }
            set { _querySecondaries = value; }
        }

        #endregion

        #region IDataErrorInfo Members

        public string Error
        {
            get
            {
                // Validate the string entered
                string pattern = @"^[1-9][0-9]{0,3}(\s[1-9][0-9]{0,3})*$|^[1-9][0-9]{0,3}((\s)*,(\s)*[1-9][0-9]{0,3})*$|^[1-9][0-9]{0,3}((\s)*\.(\s)*[1-9][0-9]{0,3})*$";
                if ((!String.IsNullOrEmpty(QuerySecondaries)) && (!Regex.IsMatch(QuerySecondaries, pattern, RegexOptions.IgnoreCase)))
                    return "Please enter a valid list of secondaries code separated by either spaces, commas or points.";
                else return null;
            }
        }

        public string this[string columnName]
        {
            get
            {
                string error = null;

                // Validate the string entered
                string pattern = @"^[1-9][0-9]{0,3}(\s[1-9][0-9]{0,3})*$|^[1-9][0-9]{0,3}((\s)*,(\s)*[1-9][0-9]{0,3})*$|^[1-9][0-9]{0,3}((\s)*\.(\s)*[1-9][0-9]{0,3})*$";
                switch (columnName)
                {
                    case "QuerySecondaries":
                        if ((!String.IsNullOrEmpty(QuerySecondaries)) && (!Regex.IsMatch(QuerySecondaries, pattern, RegexOptions.IgnoreCase)))
                            error = "Error: You must enter a valid list of secondaries codes separated by either spaces, commas or points.";
                        break;
                }

                // dirty commands registered with CommandManager so they are queried to see if they can execute now
                CommandManager.InvalidateRequerySuggested();

                return error;
            }
        }

        #endregion
    }
    //---------------------------------------------------------------------
}
