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
    class ViewModelWindowEditPriorityHabitats : ViewModelBase, IDataErrorInfo
    {
        public static HluDataSet HluDatasetStatic = null;

        #region Fields

        private ICommand _okCommand;
        private ICommand _cancelCommand;

        private string _displayName = "Priority Habitats";
        private Cursor _cursorType = Cursors.Arrow;

        private ViewModelWindowMain _viewModelMain;

        private ObservableCollection<BapEnvironment> _incidBapRowsAuto;

        #endregion

        #region Constructor

        public ViewModelWindowEditPriorityHabitats(ViewModelWindowMain viewModelMain, ObservableCollection<BapEnvironment> incidBapHabitatsAuto)
        {
            _viewModelMain = viewModelMain;

            IEnumerable<BapEnvironment> prevBapRowsAuto = null;
            prevBapRowsAuto = from p in incidBapHabitatsAuto
                         select new BapEnvironment(false, false, p.bap_id, p.incid, p.bap_habitat, p.quality_determination, p.quality_interpretation, p.interpretation_comments);

            _incidBapRowsAuto = new ObservableCollection<BapEnvironment>(prevBapRowsAuto);
            OnPropertyChanged("IncidBapHabitatsAuto");
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
        public delegate void RequestCloseEventHandler(ObservableCollection<BapEnvironment> incidBapRowsAuto);

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
            this.RequestClose(_incidBapRowsAuto);
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
                //if (_incidBapRowsAuto != null && _incidBapRowsAuto.Count > 0)
                if (_incidBapRowsAuto != null)
                {
                    int countInvalid = _incidBapRowsAuto.Count(be => !be.IsValid());
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

        #region BAP Habitat

        public HluDataSet.lut_habitat_typeRow[] BapHabitatCodes
        {
            get
            {
                return _viewModelMain.BapHabitatCodes;
            }
        }

        public HluDataSet.lut_quality_determinationRow[] BapDeterminationQualityCodesAuto
        {
            get
            {
                return _viewModelMain.BapDeterminationQualityCodesAuto;
            }
        }

        public HluDataSet.lut_quality_interpretationRow[] BapInterpretationQualityCodes
        {
            get
            {
                return _viewModelMain.InterpretationQualityCodes;
            }
        }

        public bool BapHabitatsAutoEnabled
        {
            get
            {
                return IncidBapHabitatsAuto != null && IncidBapHabitatsAuto.Count > 0;
            }
        }

        public ObservableCollection<BapEnvironment> IncidBapHabitatsAuto
        {
            get { return _incidBapRowsAuto; }
            set
            {
                _incidBapRowsAuto = value;
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
