// HLUTool is used to view and maintain habitat and land use GIS data.
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
using System.Globalization;
using System.Windows.Input;
using HLU.Properties;

namespace HLU.UI.ViewModel
{
    class ViewModelWindowNotifyOnSplitMerge : ViewModelBase
    {
        #region Fields

        private string _displayName = "Split Merge";
        private string _msgText;
        private ICommand _okCommand;

        #endregion

        #region ctor

        public ViewModelWindowNotifyOnSplitMerge(string msgText)
        {
            _msgText = msgText;
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

        public delegate void RequestCloseEventHandler();

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
                    _okCommand = new RelayCommand(okAction);
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
            Settings.Default.Save();
            this.RequestClose();
        }

        #endregion

        public string GroupBoxNotifyOnSplitMergeHeader
        {
            get { return "HLU Tool"; }
            set { }
        }

        public string LabelMessage
        {
            get
            {
                return _msgText;
            }
            set { }
        }

        public bool DoNotTellAgain
        {
            get { return !Settings.Default.NotifyOnSplitMerge; }
            set { Settings.Default.NotifyOnSplitMerge = !value; }
        }
    }
}
