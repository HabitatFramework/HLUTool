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
    class ViewModelWindowWarnOnSubsetUpdate : ViewModelBase
    {
        #region Fields

        private string _displayName = "Attribute Update";
        private int _numFrags;
        private int _numToids;
        private int _numTotalFrags;
        private int _numTotToids;
        private ViewModelWindowMain.GeometryTypes _gisFeaturesType;
        private ICommand _yesCommand;
        private ICommand _noCommand;
        private ICommand _cancelCommand;
        private bool _makeDefaultReponse;

        #endregion

        #region ctor

        public ViewModelWindowWarnOnSubsetUpdate(int numFrags, int numToids, int numTotalFrags, int numTotToids, ViewModelWindowMain.GeometryTypes typeFeatures)
        {
            _numFrags = numFrags;
            _numToids = numToids;
            _numTotalFrags = numTotalFrags;
            _numTotToids = numTotToids;
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

        public delegate void RequestCloseEventHandler(bool proceed, bool split);

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
            if (_makeDefaultReponse == true)
            {
                Settings.Default.WarnOnSubsetUpdate = "Selected";
                Settings.Default.Save();
            }
            this.RequestClose(true, true);
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
        /// Handles event when No button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void NoCommandClick(object param)
        {
            if (_makeDefaultReponse == true)
            {
                Settings.Default.WarnOnSubsetUpdate = "All";
                Settings.Default.Save();
            }
            this.RequestClose(true, false);
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
            this.RequestClose(false, false);
        }

        #endregion

        public string GroupBoxWarnOnSubsetUpdateHeader
        {
            get { return String.Format("Attempting to update subset of Incid"); }
            set { }
        }

        public string LabelMessage
        {
            get
            {
                return String.Format("Only {0} out of {2} {1}s have been selected for this Incid.\n" +
                    "Would you like to logically split the selected {1}{3} before applying the update?\n\n" +
                    "Clicking 'No' will apply the update to all of the {1}s for this Incid?",
                    _numFrags.ToString(CultureInfo.CurrentCulture),
                    _gisFeaturesType.ToString().ToLower(),
                    _numTotalFrags.ToString(CultureInfo.CurrentCulture),
                    _numFrags > 0 ? "s" : String.Empty);
            }
            set { }
        }

        public bool MakeDefaultReponse
        {
            get { return _makeDefaultReponse; }
            set { _makeDefaultReponse = value; }
        }
    }
}
