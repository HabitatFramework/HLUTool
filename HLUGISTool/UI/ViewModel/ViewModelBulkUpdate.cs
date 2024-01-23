// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2019 London & South East Record Centres (LaSER)
// Copyright © 2020 Greenspace Information for Greater London CIC
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
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Text;
using HLU.Data.Model;
using HLU.GISApplication;
using HLU.Properties;

namespace HLU.UI.ViewModel
{
    class ViewModelBulkUpdate : ViewModelBase, IDataErrorInfo
    {
        #region Fields

        private ViewModelWindowMain _viewModelMain;

        private ICommand _okCommand;
        private ICommand _cancelCommand;
        private string _displayName = "Bulk Update";

        private bool _osmmBulkUpdateMode;

        private bool _deleteOrphanBapHabitats;
        private bool _deletePotentialBapHabitats;
        private int _sourceCount;
        private int _deleteSecondaryCodes;
        private bool _createHistory;
        private string _determinationQuality;
        private string _interpretationQuality;
        private bool _ihsHabitatChanged;

        #endregion

        #region Constructor

        public ViewModelBulkUpdate(ViewModelWindowMain viewModelMain,
            bool osmmBulkUpdateMode,
            bool deleteOrphanBapHabitats,
            bool deletePotentialBapHabitats,
            int sourceCount,
            int deleteMultiplexCodes,
            bool createHistory,
            string determinationQuality,
            string interpretationQuality,
            bool ihsHabitatChanged)
        {
            _viewModelMain = viewModelMain;
            _osmmBulkUpdateMode = osmmBulkUpdateMode;
            _deleteOrphanBapHabitats = deleteOrphanBapHabitats;
            _deletePotentialBapHabitats = deletePotentialBapHabitats;
            _sourceCount = sourceCount;
            _deleteSecondaryCodes = deleteMultiplexCodes;
            _createHistory = createHistory;
            _determinationQuality = determinationQuality;
            _interpretationQuality = interpretationQuality;
            _ihsHabitatChanged = ihsHabitatChanged;

            //OnPropertyChanged("ShowInOSMMBulkUpdateMode");
            //OnPropertyChanged("HideInOSMMBulkUpdateMode");
            //OnPropertyChanged("OSMMBulkUpdateCommandHeader");
            //OnPropertyChanged("BapDeterminationQualityCodesAuto");
            //OnPropertyChanged("DeterminationQuality");
            //OnPropertyChanged("BapInterpretationQualityCodes");
            //OnPropertyChanged("InterpretationQuality");
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
        public delegate void RequestCloseEventHandler(bool apply,
            bool bulkDeleteOrphanBapHabitats,
            bool bulkDeletePotentialBapHabitats,
            int bulkDeleteMultiplexCodes,
            bool bulkCreateHistory,
            string bulkDeterminationQuality,
            string bulkInterpretationQuality);

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
            this.RequestClose(true,
                _deleteOrphanBapHabitats,
                _deletePotentialBapHabitats,
                _deleteSecondaryCodes,
                _createHistory,
                _determinationQuality,
                _interpretationQuality);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanOk
        {
            get
            {
                return true;
            }
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
            this.RequestClose(false, false, false, 1, false, null, null);
        }

        #endregion

        #region Interface

        /// <summary>
        /// Gets or sets whether orphan bap habitats should be deleted.
        /// </summary>
        /// <value>
        ///   <c>true</c> if orphan bap habittas should be deleted; otherwise, <c>false</c>.
        /// </value>
        public bool DeleteOrphanBapHabitats
        {
            get { return _deleteOrphanBapHabitats; }
            set { _deleteOrphanBapHabitats = value; }
        }

        /// <summary>
        /// Gets or sets whether existing potential bap habitats should be deleted.
        /// </summary>
        /// <value>
        ///   <c>true</c> if existing potential bap habittas should be deleted; otherwise, <c>false</c>.
        /// </value>
        public bool DeletePotentialBapHabitats
        {
            get { return _deletePotentialBapHabitats; }
            set { _deletePotentialBapHabitats = value; }
        }

        /// <summary>
        /// Gets or sets whether existing sources should be deleted.
        /// </summary>
        /// <value>
        ///   <c>true</c> if existing sources should be deleted; otherwise, <c>false</c>.
        /// </value>
        public bool DeleteSources
        {
            get { return (_sourceCount != 0); }
        }

        /// <summary>
        /// Gets or sets whether existing multiplex codes should be deleted.
        /// </summary>
        /// <value>
        ///   <c>true</c> if existing multiplex codes should be deleted; otherwise, <c>false</c>.
        /// </value>
        public DeleteSecondaryCodesAction[] DeleteSecondaryCodesActions
        {
            get
            {
                return Enum.GetValues(typeof(DeleteSecondaryCodesAction)).Cast<DeleteSecondaryCodesAction>()
                    .ToArray();
            }
            set { }
        }

        /// <summary>
        /// Gets or sets whether existing multiplex codes should be deleted.
        /// </summary>
        /// <value>
        ///   <c>Always</c> if all existing multiplex codes should be deleted; otherwise,
        /// <c>Invalid</c> if only invalid mulitplex codes should be deleted.
        /// </value>
        public DeleteSecondaryCodesAction? DeleteSecondaryCodes
        {
            get { return (DeleteSecondaryCodesAction)_deleteSecondaryCodes; }
            set
            {
                _deleteSecondaryCodes = (int)value;
            }
        }

        /// <summary>
        /// Gets or sets whether history records should be created.
        /// </summary>
        /// <value>
        ///   <c>true</c> if history records should be created; otherwise, <c>false</c>.
        /// </value>
        public bool CreateHistory
        {
            get { return _createHistory; }
            set { _createHistory = value; }
        }

        /// <summary>
        /// Gets the determination quality codes that can be applied to primary
        /// bap habitats.
        /// </summary>
        /// <value>
        /// The primary bap determination quality codes.
        /// </value>
        public HluDataSet.lut_quality_determinationRow[] BapDeterminationQualityCodes
        {
            get
            {
                return _viewModelMain.BapDeterminationQualityCodesAuto;
            }
        }

        /// <summary>
        /// Gets the Determination Quality to apply for new
        /// Bap Habitats in OSMM Bulk Update mode.
        /// </summary>
        /// <value>
        /// The Determination Quality to apply for OSMM Bulk Updates.
        /// </value>
        public string DeterminationQuality
        {
            get
            {
                return _determinationQuality;
            }
            set { _determinationQuality = value; }
        }

        /// <summary>
        /// Gets the interpretation quality codes that can be applied to primary
        /// bap habitats.
        /// </summary>
        /// <value>
        /// The primary bap interpretation quality codes.
        /// </value>
        public HluDataSet.lut_quality_interpretationRow[] InterpretationQualityCodes
        {
            get
            {
                return _viewModelMain.InterpretationQualityCodes;
            }
        }

        /// <summary>
        /// Gets the Interpretation Quality to apply for new
        /// Bap Habitats in OSMM Bulk Update mode.
        /// </summary>
        /// <value>
        /// The Interpretation Quality to apply for OSMM Bulk Updates.
        /// </value>
        public string InterpretationQuality
        {
            get
            {
                return _interpretationQuality;
            }
            set { _interpretationQuality = value; }
        }

        /// <summary>
        /// Hide control when in OSMM Bulk Update mode.
        /// </summary>
        public Visibility HideInOSMMBulkUpdateMode
        {
            get
            {
                // Hide the control if in OSMM Bulk Update mode
                if (_osmmBulkUpdateMode == true)
                    return Visibility.Hidden;
                else
                    return Visibility.Visible;
            }
            set { }
        }

        /// <summary>
        /// Show control when in OSMM Bulk Update mode.
        /// </summary>
        public Visibility ShowInOSMMBulkUpdateMode
        {
            get
            {
                // Show the control if in OSMM Bulk Update mode
                if (_osmmBulkUpdateMode == true)
                {
                    return Visibility.Visible;
                }
                else
                {
                    return Visibility.Collapsed;
                }
            }
            set { }
        }

        /// <summary>
        /// Enable control when in OSMM Bulk Update mode.
        /// </summary>
        public bool EnableInOSMMBulkUpdateMode
        {
            get
            {
                // Enable the control if in OSMM Bulk Update mode
                if (_osmmBulkUpdateMode == true)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            set { }
        }

        /// <summary>
        /// Disable control when in OSMM Bulk Update mode.
        /// </summary>
        public bool DisableInOSMMBulkUpdateMode
        {
            get
            {
                // Disable the control if in OSMM Bulk Update mode
                if (_osmmBulkUpdateMode == false)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            set { }
        }

        /// <summary>
        /// Enable control when in OSMM Bulk Update mode and
        /// the IHS habitat has changed.
        /// </summary>
        public bool EnableDeleteOrphanBapHabitats
        {
            get
            {
                // Enable the control if the IHS habitat has changed
                if (_ihsHabitatChanged == true)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            set { }
        }

        /// <summary>
        /// Enable control when in OSMM Bulk Update mode and
        /// the IHS habitat has changed.
        /// </summary>
        public bool EnableDeletePotentialBapHabitats
        {
            get
            {
                // Enable the control if the IHS habitat has changed
                if (_ihsHabitatChanged == true)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            set { }
        }

        /// <summary>
        /// Enable control when in OSMM Bulk Update mode and
        /// the IHS habitat has changed.
        /// </summary>
        public bool EnableDeleteSecondaryCodes
        {
            get
            {
                // Enable the control if not in OSMM Bulk Update mode and
                // the IHS habitat has changed
                if ((_osmmBulkUpdateMode == false) && (_ihsHabitatChanged == true))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            set { }
        }

        #endregion

        #region IDataErrorInfo Members

        public string Error
        {
            get
            {
                StringBuilder error = new StringBuilder();

                if (String.IsNullOrEmpty(DeterminationQuality))
                    error.Append(Environment.NewLine).Append("You must choose a Determination Quality");

                if (String.IsNullOrEmpty(InterpretationQuality))
                    error.Append(Environment.NewLine).Append("You must choose an Interpretation Quality");

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

                switch (columnName)
                {
                    //TODO: Bulk Update - delete secondary codes and/or multiplex codes when relevant
                    case "DeleteSecondaryCodes":
                        if ((EnableDeleteSecondaryCodes == true) &&
                            (DeleteSecondaryCodes == (int)DeleteSecondaryCodesAction.All))
                            error="Warning: This option will delete ALL secondary codes from all affected incids";
                        break;
                    case "DeterminationQuality":
                        if (String.IsNullOrEmpty(DeterminationQuality))
                            error= "Error: You must choose a Determination Quality";
                        break;
                    case "InterpretationQuality":
                        if (String.IsNullOrEmpty(InterpretationQuality))
                            error = "Error: You must choose an Interpretation Quality";
                        break;
                }

                CommandManager.InvalidateRequerySuggested();

                return error;
            }
        }

        #endregion
    }
}
