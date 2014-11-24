// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// Copyright © 2014 Sussex Biodiversity Record Centre
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
using System.Windows.Input;
using HLU.Data.Model;
using HLU.GISApplication;

namespace HLU.UI.ViewModel
{
    class ViewModelExport : ViewModelBase, IDataErrorInfo
    {
        #region Fields
        
        private ICommand _okCommand;
        private ICommand _cancelCommand;
        private string _displayName = "Export";
        private string _layerName;
        private HluDataSet.exportsDataTable _exportFormats;
        private int _exportID = -1;
        private bool _selectedOnly;
        private int _selectedNumber;
        private GISApplications _gisApp;

        #endregion

        #region Constructor

        public ViewModelExport(int numberSelected, string layerName, 
            GISApplications gisApp, HluDataSet.exportsDataTable exportFormats)
        {
            _selectedNumber = numberSelected;
            _selectedOnly = _selectedNumber > 0;
            _layerName = layerName;
            _gisApp = gisApp;
            _exportFormats = exportFormats;
            if (_exportFormats.Count == 1)
                _exportID = _exportFormats[0].export_id;
        }

        #endregion

        #region ViewModelBase Members

        public override string DisplayName
        {
            get
            {
                return _displayName;
            }
            set
            {
                _displayName = value;
            }
        }

        public override string WindowTitle
        {
            get { return _displayName; }
        }

        #endregion

        #region RequestClose

        // declare the delegate since using non-generic pattern
        public delegate void RequestCloseEventHandler(int exportID, bool selectedOnly);

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
            //---------------------------------------------------------------------
            // CHANGED: CR14 (Exporting IHS codes or descriptions)
            // Enable users to specify if individual fields should be
            // exported with descriptions, rather than the whole export,
            // by moving this option to the exports_fields table.
            //
            this.RequestClose(_exportID, _selectedOnly);
            //this.RequestClose(_exportID, _exportDescriptions, _selectedOnly);
            //---------------------------------------------------------------------
        }

        /// <summary>
        /// 
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanOk { get { return _exportID != -1; } }

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
            this.RequestClose(-1, false);
        }

        #endregion

        public string LayerName
        {
            get { return _layerName; }
            set { _layerName = value; }
        }

        public HluDataSet.exportsDataTable ExportFormats
        {
            get { return _exportFormats; }
            set { }
        }

        public int ExportID
        {
            get { return _exportID; }
            set { _exportID = value; }
        }

        public bool HaveSelection { get { return _selectedNumber > 0; } }

        public bool SelectedOnly
        {
            get { return _selectedOnly; }
            set { _selectedOnly = value; }
        }

        public string SelectedNumber
        {
            get { return HaveSelection ? String.Format("({0} selected feature{1})", _selectedNumber, _selectedNumber > 1 ? "s" : String.Empty) : String.Empty; }
        }

        #region IDataErrorInfo Members

        public string Error
        {
            get
            {
                if (_exportID == -1) return "Please choosed an export format";
                else return null;
            }
        }

        public string this[string columnName]
        {
            get
            {
                string error = null;

                switch (columnName)
                {
                    case "ExportID":
                        if (_exportID == -1) 
                            error = "Please choosed an export format";
                        break;
                }

                return error;
            }
        }

        #endregion
    }
}
