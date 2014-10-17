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
using System.Text;
using System.Globalization;
using System.Windows.Input;
using HLU.Properties;

namespace HLU.UI.ViewModel
{
    class ViewModelWindowWarnOnGISSelect : ViewModelBase
    {
        #region Fields

        private string _displayName = "Select on Map";
        private int _gisFeaturesNum;
        private int _gisIncidNum;
        private string _gisFeaturesType;
        private bool _selectByjoin;
        private ICommand _noCommand;
        private ICommand _yesCommand;

        #endregion

        #region ctor

        public ViewModelWindowWarnOnGISSelect(int numFeatures, int numIncids, ViewModelWindowMain.GeometryTypes typeFeatures, bool selectByjoin)
        {
            // Store the expected number of features to be selected in GIS.
            _gisFeaturesNum = numFeatures;
            //---------------------------------------------------------------------
            // CHANGED: CR12 (Select by attribute performance)
            // Store the expected number of incids to be selected in GIS.
            _gisIncidNum = numIncids;
            // If the type of feature is not known then just use 'feature'.
            if (typeFeatures != ViewModelWindowMain.GeometryTypes.Unknown)
                _gisFeaturesType = typeFeatures.ToString().ToLower();
            else
                _gisFeaturesType = "feature";
            // Store if a GIS table join will be used to perform the selection.
            _selectByjoin = selectByjoin;
            //---------------------------------------------------------------------
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

        public delegate void RequestCloseEventHandler(bool proceed);

        public event RequestCloseEventHandler RequestClose;

        #endregion

        #region Yes Command

        /// <summary>
        /// Set Yes button command.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand YesCommand
        {
            get
            {
                if (_yesCommand == null)
                {
                    Action<object> yesAction = new Action<object>(this.YesCommandClick);
                    _yesCommand = new RelayCommand(yesAction);
                }

                return _yesCommand;
            }
        }

        /// <summary>
        /// Handles events when the Yes button is clicked.
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void YesCommandClick(object param)
        {
            Settings.Default.Save();
            this.RequestClose(true);
        }

        #endregion

        #region No Command

        /// <summary>
        /// Set No button command.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand NoCommand
        {
            get
            {
                if (_noCommand == null)
                {
                    Action<object> noAction = new Action<object>(this.NoCommandClick);
                    _noCommand = new RelayCommand(noAction);
                }

                return _noCommand;
            }
        }

        /// <summary>
        /// Handles events when the Cancel button is clicked.
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void NoCommandClick(object param)
        {
            this.RequestClose(false);
        }

        #endregion

        public string GroupBoxWarnOnGISSelectHeader
        {
            get { return String.Format("Expected number of selected {0}s & incids", _gisFeaturesType); }
            set { }
        }

        public string LabelMessage
        {
            get
            {
                //---------------------------------------------------------------------
                // CHANGED: CR12 (Select by attribute performance)
                // Display the expected number of features and/or the expected
                // number of incids to be selected in GIS, depending upon which
                // values are valid.
                StringBuilder labelMsg = new StringBuilder();
                if (_gisFeaturesNum < 0 && _gisIncidNum < 0)
                {
                    labelMsg.Append(String.Format("Could not determine the number of {0}{1} or incids expected to be " +
                        "selected on the map by this operation.",
                        _gisFeaturesType,
                        _gisFeaturesNum > 1 ? "s" : String.Empty));
                }
                else if (_gisFeaturesNum < 0)
                {
                    labelMsg.Append(String.Format("This operation is expected to select {0} incid{1} on the map.\n" +
                        "The expected number of {2}s could not be determined.",
                        String.Format(CultureInfo.CurrentCulture, "{0:N0}", _gisIncidNum),
                        _gisIncidNum > 1 ? "s" : String.Empty,
                        _gisFeaturesType));
                }
                else if (_gisIncidNum < 0)
                {
                    labelMsg.Append(String.Format("This operation is expected to select {0} {1}{2} on the map.\n" +
                        "The expected number of incids could not be determined.",
                        String.Format(CultureInfo.CurrentCulture, "{0:N0}", _gisFeaturesNum),
                        _gisFeaturesType,
                        _gisFeaturesNum > 1 ? "s" : String.Empty));
                }
                else
                {
                    labelMsg.Append(String.Format("This operation is expected to select {0} {1}{2} from {3} incid{4} on the map.",
                        String.Format(CultureInfo.CurrentCulture, "{0:N0}", _gisFeaturesNum),
                        _gisFeaturesType,
                        _gisFeaturesNum > 1 ? "s" : String.Empty,
                        String.Format(CultureInfo.CurrentCulture, "{0:N0}", _gisIncidNum),
                        _gisIncidNum > 1 ? "s" : String.Empty));
                }

                // Advise the user that a GIS table join will be used to
                // perform the selection.
                if (_selectByjoin)
                    labelMsg.Append("\n\nThe operation will be performed using a table join in GIS which make take some time.");

                labelMsg.Append("\n\nWould you like to proceed?");

                return labelMsg.ToString();
                //---------------------------------------------------------------------
            }
            set { }
        }

        public bool DoNotAskAgain
        {
            get { return false; }
            set
            {
                //---------------------------------------------------------------------
                // CHANGED: CR12 (Select by attribute performance)
                // Only set the warning down a 'notch' based on the
                // current warning level.
                if (Settings.Default.WarnBeforeGISSelect == 0)
                    Settings.Default.WarnBeforeGISSelect = 1;
                else if ((_selectByjoin) & (Settings.Default.WarnBeforeGISSelect == 1))
                    Settings.Default.WarnBeforeGISSelect = 2;
                //---------------------------------------------------------------------
            }
        }
    }
}
