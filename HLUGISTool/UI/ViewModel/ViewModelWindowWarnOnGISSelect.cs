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
using System.Globalization;
using System.Windows.Input;
using HLU.Properties;

namespace HLU.UI.ViewModel
{
    class ViewModelWindowWarnOnGISSelect : ViewModelBase
    {
        #region Fields

        private string _displayName = "GIS Selection";
        private int _gisFeaturesNum;
        private ViewModelWindowMain.GeometryTypes _gisFeaturesType;
        private ICommand _noCommand;
        private ICommand _yesCommand;

        #endregion

        #region ctor

        public ViewModelWindowWarnOnGISSelect(int numFeatures, ViewModelWindowMain.GeometryTypes typeFeatures)
        {
            _gisFeaturesNum = numFeatures;
            _gisFeaturesType = typeFeatures;
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
        /// Create Yes button command
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
        /// Handles event when Yes button is clicked
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
        /// Create No button command
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
        /// Handles event when Cancel button is clicked
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
            get { return String.Format("Expected Number of Selected {0}s", _gisFeaturesType); }
            set { }
        }

        public string LabelMessage
        {
            get
            {
                if (_gisFeaturesNum < 0)
                {
                    return String.Format("Could not determine the number of {0}{1} expected to be " + 
                        "selected on the map by this operation.\nWould you like to proceed?",
                        _gisFeaturesType.ToString().ToLower(), _gisFeaturesNum > 1 ? "s" : String.Empty);
                }
                else if (_gisFeaturesType == ViewModelWindowMain.GeometryTypes.Unknown)
                {
                    return String.Format("The expected number of {0}{1} selected on the map could not be determined.\n" +
                        "This operation will select {0}{1} on the map related to {2} INCIDs.\nWould you like to proceed?",
                        _gisFeaturesType.ToString().ToLower(), _gisFeaturesNum > 1 ? "s" : String.Empty, 
                        _gisFeaturesNum.ToString(CultureInfo.CurrentCulture));
                }
                else
                {
                    return String.Format("This operation is expected to select {0} {1}{2} on the map.\nWould you like to proceed?",
                        _gisFeaturesNum.ToString(CultureInfo.CurrentCulture), _gisFeaturesType.ToString().ToLower(), 
                        _gisFeaturesNum > 1 ? "s" : String.Empty);
                }
            }
            set { }
        }

        public bool DoNotAskAgain
        {
            get { return !Settings.Default.WarnOnGISSelect; }
            set { Settings.Default.WarnOnGISSelect = !value; }
        }
    }
}
