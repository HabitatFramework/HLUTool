// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// Copyright © 2014, 2016 Thames Valley Environmental Records Centre
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
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;
using HLU.GISApplication;
using HLU.Data;

namespace HLU.UI.ViewModel
{
    //---------------------------------------------------------------------
    // CHANGED: CR31 (Switching between GIS layers)
    // Enable the user to switch between different HLU layers, where
    // there is more than one valid layer in the current document.
    //
    // This is the view model supporting the user interface that
    // shows the list of current valid HLU layers, indicates
    // which layer is currently selected and allows the user to
    // switch to another layer.
    //---------------------------------------------------------------------
    class ViewModelWindowSwitchGISLayer : ViewModelBase, IDataErrorInfo
    {
        #region Fields
        
        private string _displayName = "Select HLU Feature Layer";
        private List<GISLayer> _availableHLULayers;
        private GISLayer _selectedHLULayer;
        private int _mapWindowsCount;
        private ICommand _okCommand;
        private ICommand _cancelCommand;

        #endregion

        //---------------------------------------------------------------------
        // FIX: 059 Do not display map window number with layer name
        // if there is only one map window.
        // 
        public ViewModelWindowSwitchGISLayer(List<GISLayer> ValidHluLayers, GISLayer SelectedHLULayer, int MapWindowsCount)
        {
            if (ValidHluLayers != null)
            {
                _availableHLULayers = ValidHluLayers;
                _selectedHLULayer = SelectedHLULayer;
                _mapWindowsCount = MapWindowsCount;

                // Set the static variable for the total number of
                // map windows.
                GISLayer.MapWindowsCount = _mapWindowsCount;
            }
        }
        //---------------------------------------------------------------------

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

        public delegate void RequestCloseEventHandler(bool doSwitch, GISLayer selectedHLULayer);

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
            this.RequestClose(true, _selectedHLULayer);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanOk { get { return String.IsNullOrEmpty(Error); } }

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
            this.RequestClose(false, null);
        }

        #endregion

        #region Properties

        public List<GISLayer> AvailableHLULayers
        {
            get
            {
                return _availableHLULayers;
            }
            set { }
        }

        public GISLayer SelectedHLULayer
        {
            get
            {
                return _selectedHLULayer;
            }

            set { _selectedHLULayer = value; }
        }

        #endregion

        #region Implementation of IDataErrorInfo

        public string Error
        {
            get
            {
                if ((_selectedHLULayer == null) || !AvailableHLULayers.Contains(_selectedHLULayer))
                    return "This tool requires a valid HLU Layer.";
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
                    case "SelectedHLULayer":
                        if ((_selectedHLULayer == null) || !AvailableHLULayers.Contains(_selectedHLULayer))
                            error = "Error: This tool requires a valid HLU Layer.";
                        break;
                }
                CommandManager.InvalidateRequerySuggested();
                return error;
            }
        }
        
        #endregion
    }
}
