using System;
using System.ComponentModel;
using System.Windows.Input;
using HLU.Data.Model;
using HLU.GISApplication;

namespace HLU.UI.ViewModel
{
    class ViewModelExport : ViewModelBase, IDataErrorInfo
    {
        #region Fields
        
        private ICommand _okCommand;
        private ICommand _cancelCommand;
        private string _displayName = "Export";
        private string _layerName;
        private HluDataSet.exportsDataTable _exportFormats;
        private int _exportID = -1;
        private bool _selectedOnly;
        private bool _exportDescriptions = false;
        private int _selectedNumber;
        private GISApplications _gisApp;

        #endregion

        #region Constructor

        public ViewModelExport(int numberSelected, string layerName, 
            GISApplications gisApp, HluDataSet.exportsDataTable exportFormats)
        {
            _selectedNumber = numberSelected;
            _selectedOnly = _selectedNumber > 0;
            _layerName = layerName;
            _gisApp = gisApp;
            _exportFormats = exportFormats;
            if (_exportFormats.Count == 1)
                _exportID = _exportFormats[0].export_id;
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
        public delegate void RequestCloseEventHandler(int exportID, bool exportDescriptions, bool selectedOnly);

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
            this.RequestClose(_exportID, _exportDescriptions, _selectedOnly);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanOk { get { return _exportID != -1; } }

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
            this.RequestClose(-1, false, false);
        }

        #endregion

        public string LayerName
        {
            get { return _layerName; }
            set { _layerName = value; }
        }

        public HluDataSet.exportsDataTable ExportFormats
        {
            get { return _exportFormats; }
            set { }
        }

        public int ExportID
        {
            get { return _exportID; }
            set { _exportID = value; }
        }

        public bool HaveSelection { get { return _selectedNumber > 0; } }

        public bool SelectedOnly
        {
            get { return _selectedOnly; }
            set { _selectedOnly = value; }
        }

        public string SelectedNumber
        {
            get { return HaveSelection ? String.Format("({0} selected)", _selectedNumber) : String.Empty; }
        }

        public bool ExportDescriptions
        {
            get { return CanExportDescriptions && _exportDescriptions; }
            set { if (CanExportDescriptions) _exportDescriptions = value; }
        }

        public bool CanExportDescriptions
        {
            get { return _gisApp != GISApplications.MapInfo; }
        }

        #region IDataErrorInfo Members

        public string Error
        {
            get
            {
                if (_exportID == -1) return "Please choosed an export format";
                else return null;
            }
        }

        public string this[string columnName]
        {
            get
            {
                string error = null;

                switch (columnName)
                {
                    case "ExportID":
                        if (_exportID == -1) 
                            error = "Please choosed an export format";
                        break;
                }

                return error;
            }
        }

        #endregion
    }
}
