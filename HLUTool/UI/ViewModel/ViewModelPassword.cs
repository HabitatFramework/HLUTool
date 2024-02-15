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
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;

namespace HLU.UI.ViewModel
{
    class ViewModelPassword : ViewModelBase, IDataErrorInfo
    {
        #region private Members

        private string _displayName;
        private RelayCommand _okCommand;
        private RelayCommand _cancelCommand;
        private string _userLabel;
        private string _userText;
        private string _password;

        #endregion

        #region Constructor

        public ViewModelPassword()
        {
        }

        #endregion

        #region Display Name

        public override string DisplayName
        {
            get { return _displayName; }
            set { _displayName = value; }
        }

        #endregion

        #region Window Title

        public override string WindowTitle { get { return DisplayName; } }

        #endregion

        #region RequestClose

        // declare the delegate since using non-generic pattern
        public delegate void RequestCloseEventHandler(string password, string errorMsg);

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
            try
            {
                this.RequestClose(_password, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Blank password may be ok, but this dialog shall only be raised when the password is known to be non-blank.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanOk { get { return !String.IsNullOrEmpty(_password); } }

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
            this.RequestClose(null, null);
        }

        #endregion

        #region User

        public string UserLabel
        {
            get { return _userLabel; }
            set { _userLabel = value; }
        }

        public string UserText
        {
            get { return _userText; }
            set { _userText = value; }
        }

        #endregion

        #region Password

        public string Password
        {
            get { return _password; }
            set { _password = value; }
        }

        #endregion

        #region IDataErrorInfo Members

        string IDataErrorInfo.Error
        {
            get
            {
                if (String.IsNullOrEmpty(_password))
                    return "Please provide a password";
                else
                    return null;
            }
        }

        string IDataErrorInfo.this[string columnName]
        {
            get
            {
                string error = null;

                switch (columnName)
                {
                    case "Password":
                        if (String.IsNullOrEmpty(_password))
                            error = "Error: You must provide a password";
                        break;
                }

                // dirty commands registered with CommandManager so they are queried to see if they can execute now
                CommandManager.InvalidateRequerySuggested();

                return error;
            }
        }

        #endregion
    }
}
