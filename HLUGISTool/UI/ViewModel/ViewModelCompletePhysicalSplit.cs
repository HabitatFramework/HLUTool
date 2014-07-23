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
using System.Text;
using System.Windows.Input;
using HLU.Data.Model;

namespace HLU.UI.ViewModel
{
    class ViewModelCompletePhysicalSplit : ViewModelBase, IDataErrorInfo
    {
        #region Fields

        string _displayName = "Physical Split: required values";
        private HluDataSet.lut_reasonRow[] _reasonCodes;
        private HluDataSet.lut_processRow[] _processCodes;
        private string _reason;
        private string _process;
        private ICommand _okCommand;
        private ICommand _cancelCommand;

        #endregion

        #region ctor

        public ViewModelCompletePhysicalSplit(string reason, string process, 
            HluDataSet.lut_reasonRow[] reasonCodes, HluDataSet.lut_processRow[] processCodes)
        {
            _reason = reason;
            _process = process;
            _reasonCodes = reasonCodes;
            _processCodes = processCodes;
        }

        #endregion

        #region ViewModelBase members

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

        public delegate void RequestCloseEventHandler(string reason, string process);

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
            if (this.RequestClose != null)
                this.RequestClose(_reason, _process);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanOk { get { return String.IsNullOrEmpty(Error); } }

        #endregion

        #region Properties

        public HluDataSet.lut_reasonRow[] ReasonCodes
        {
            get { return _reasonCodes; }
            set { }
        }

        public HluDataSet.lut_processRow[] ProcessCodes
        {
            get { return _processCodes; }
            set { }
        }

        public string Reason
        {
            get { return _reason; }
            set { _reason = value; }
        }

        public string Process
        {
            get { return _process; }
            set { _process = value; }
        }

        #endregion

        #region IDataErrorInfo Members

        public string Error
        {
            get
            {
                StringBuilder error = new StringBuilder();
                if (String.IsNullOrEmpty(_reason))
                    error.Append("Please select a value for Reason.");
                if (String.IsNullOrEmpty(_process))
                    error.Append("Please select a value for Process.");
                if (error.Length > 0)
                    return error.ToString();
                else
                    return null;
            }
        }

        public string this[string columnName]
        {
            get
            {
                string error = null;
                switch (columnName)
                {
                    case "Reason":
                        if (String.IsNullOrEmpty(_reason))
                            error = "Please select a value for Reason.";
                        break;
                    case "Process":
                        if (String.IsNullOrEmpty(_process))
                            error = "Please select a value for Process.";
                        break;
                }
                CommandManager.InvalidateRequerySuggested();
                return error;
            }
        }

        #endregion
    }
}
