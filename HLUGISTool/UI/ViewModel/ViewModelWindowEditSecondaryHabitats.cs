// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2019 Greenspace Information for Greater London CIC
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
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using HLU.Data;
using HLU.Data.Model;
using HLU.UI.View;

namespace HLU.UI.ViewModel
{
    class ViewModelWindowEditSecondaryHabitats : ViewModelBase, IDataErrorInfo
    {
        public static HluDataSet HluDatasetStatic = null;

        #region Fields

        private ICommand _okCommand;
        private ICommand _cancelCommand;

        private string _displayName = "Secondary Habitats";
        private Cursor _cursorType = Cursors.Arrow;

        private ViewModelWindowMain _viewModelMain;

        private ObservableCollection<SecondaryHabitat> _incidSecondaryHabitats;

        #endregion

        #region Constructor

        public ViewModelWindowEditSecondaryHabitats(ViewModelWindowMain viewModelMain, ObservableCollection<SecondaryHabitat> incidSecondaryHabitats)
        {
            _viewModelMain = viewModelMain;

            IEnumerable<SecondaryHabitat> prevSecondaryHabitats = null;
            prevSecondaryHabitats = from p in incidSecondaryHabitats
                                    select new SecondaryHabitat(false, p.secondary_id, p.incid, p.secondary_habitat, p.secondary_group);

            _incidSecondaryHabitats = new ObservableCollection<SecondaryHabitat>(prevSecondaryHabitats);
            OnPropertyChanged("IncidSecondaryHabitats");
        }

        #endregion

        #region ViewModelBase Members

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

        // declare the delegate since using non-generic pattern
        public delegate void RequestCloseEventHandler(ObservableCollection<SecondaryHabitat> incidSecondaryHabitats);

        // declare the event
        public event RequestCloseEventHandler RequestClose;

        #endregion

        #region Ok Command

        /// <summary>
        /// Set the Ok button command.
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
        /// Handles events when the Ok button is clicked.
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void OkCommandClick(object param)
        {
            this.RequestClose(_incidSecondaryHabitats);
        }

        /// <summary>
        /// Determine if the Ok button can be clicked.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanOk
        {
            get
            {
                //if (_incidSecondaryHabitats != null && _incidSecondaryHabitats.Count > 0)
                if (_incidSecondaryHabitats != null)
                {
                    int countInvalid = _incidSecondaryHabitats.Count(sh => !sh.IsValid());
                    if (countInvalid > 0)
                        return false;
                    else
                        return true;
                }
                return false;
            }
        }

        #endregion

        #region Cancel Command

        /// <summary>
        /// Set the Cancel button command.
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
        /// Handles events when the Cancel button is clicked.
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void CancelCommandClick(object param)
        {
            this.RequestClose(null);
        }

        #endregion

        #region Secondary Habitat

        public HluDataSet.lut_secondary_groupRow[] SecondaryGroupCodes
        {
            get
            {
                return _viewModelMain.SecondaryGroupCodes;
            }
        }

        public HluDataSet.lut_secondaryRow[] SecondaryCodes
        {
            get
            {
                return _viewModelMain.SecondaryCodes;
            }
        }

        public bool SecondaryHabitatsEnabled
        {
            get
            {
                return true;
            }
        }

        public ObservableCollection<SecondaryHabitat> IncidSecondaryHabitats
        {
            get { return _incidSecondaryHabitats; }
            set
            {
                _incidSecondaryHabitats = value;
            }
        }

        #endregion

        #region IDataErrorInfo Members

        public string Error
        {
            get
            {
                StringBuilder error = new StringBuilder();

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

                CommandManager.InvalidateRequerySuggested();

                return error;
            }
        }

        #endregion

    }
}
