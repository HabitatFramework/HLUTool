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
using System.Windows.Input;

namespace HLU.UI.ViewModel
{
    class ViewModelMessageWindow : ViewModelBase
    {
        #region Private Members

        private RelayCommand _okCommand;
        private string _messageHeader;
        private string _messageText;

        #endregion

        #region Window Title

        public override string DisplayName
        {
            get { return _messageHeader; }
            set { _messageHeader = value; }
        }

        public override string WindowTitle
        {
            get { return DisplayName; }
        }

        #endregion

        #region RequestClose

        public EventHandler RequestClose;

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
            this.RequestClose(this, EventArgs.Empty);
        }

        #endregion

        #region Message

        public string MessageText
        {
            get { return _messageText; }
            set { _messageText = value; }
        }

        public string MessageHeader
        {
            get { return _messageHeader; }
            set { _messageHeader = value; }
        }

        #endregion
    }
}
