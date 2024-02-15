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
using System.Windows.Input;

namespace HLU.UI.ViewModel
{
    class ViewModelWindowAbout : ViewModelBase
    {
        #region Fields

        private RelayCommand _okCommand;
        private string _displayName = "About HLU Tool";
        private string _appVersion;
        private string _dbVersion;
        private string _connectionType;
        private string _connectionSettings;
        private string _userId;
        private string _userName;
        private string _copyright;
        private string _userGuideURL;
        private string _userGuideText;
        private string _technicalGuideURL;
        private string _technicalGuideText;

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
            get { return DisplayName; }
        }

        #endregion

        #region RequestClose

        // declare the delegate since using non-generic pattern
        public delegate void RequestCloseEventHandler();

        // declare the event
        public event RequestCloseEventHandler RequestClose;

        #endregion

        #region Ok Command

        /// <summary>
        /// Create Ok button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        public ICommand OkCommand
        {
            get
            {
                if (_okCommand == null)
                {
                    Action<object> okAction = new Action<object>(this.Ok);
                    _okCommand = new RelayCommand(okAction);
                }
                return _okCommand;
            }
        }

        /// <summary>
        /// Handles event when Ok button is clicked
        /// </summary>
        /// <param name="param"></param>
        private void Ok(object param)
        {
            this.RequestClose();
        }

        #endregion

        public string AppVersion
        {
            get { return _appVersion; }
            set { _appVersion = value; }
        }

        public string DbVersion
        {
            get { return _dbVersion; }
            set { _dbVersion = value; }
        }

        public string ConnectionType
        {
            get { return _connectionType; }
            set { _connectionType = value; }
        }

        public string ConnectionSettings
        {
            get { return _connectionSettings; }
            set { _connectionSettings = value; }
        }

        public string UserId
        {
            get { return _userId; }
            set { _userId = value; }
        }

        public string UserName
        {
            get { return _userName; }
            set { _userName = value; }
        }

        public string Copyright
        {
            get { return _copyright; }
            set { _copyright = value; }
        }

        public string UserGuideURL
        {
            get { return _userGuideURL; }
            set { _userGuideURL = value; }
        }

        public string UserGuideText
        {
            get { return _userGuideText; }
            set { _userGuideText = value; }
        }

        public string TechnicalGuideURL
        {
            get { return _technicalGuideURL; }
            set { _technicalGuideURL = value; }
        }

        public string TechnicalGuideText
        {
            get { return _technicalGuideText; }
            set { _technicalGuideText = value; }
        }
    }
}
