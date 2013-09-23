using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;
using HLU.GISApplication;

namespace HLU.UI.ViewModel
{
    class ViewModelWindowSelectGIS : ViewModelBase, IDataErrorInfo
    {
        #region Fields
        
        private string _displayName = "Select GIS Application";
        private GISApplications[] _availableGISApps;
        private GISApplications _selectedGISApp;
        private ICommand _okCommand;
        private ICommand _cancelCommand;

        #endregion

        public ViewModelWindowSelectGIS()
        {
            _availableGISApps = Enum.GetValues(typeof(GISApplications)).Cast<GISApplications>().Where(g => g != GISApplications.None).ToArray();
            if (_availableGISApps.Length > 0)
                _selectedGISApp = _availableGISApps[0];
        }

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

        public delegate void RequestCloseEventHandler(bool cancelled, GISApplications selectedGISApp);

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
            this.RequestClose(false, _selectedGISApp);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanOk { get { return String.IsNullOrEmpty(Error); } }

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
            this.RequestClose(true, GISApplications.None);
        }

        #endregion

        #region Properties

        public GISApplications[] GISApps
        {
            get { return _availableGISApps; }
            set { }
        }

        public GISApplications SelectedGISApp
        {
            get { return _selectedGISApp; }
            set { _selectedGISApp = value; }
        }

        #endregion

        #region Implementation of IDataErrorInfo

        public string Error
        {
            get
            {
                if (_selectedGISApp == GISApplications.None)
                    return "This tool requires a GIS application.";
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
                    case "SelectedGISApp":
                        if (_selectedGISApp == GISApplications.None)
                            error = "This tool requires a GIS application.";
                        break;
                }
                CommandManager.InvalidateRequerySuggested();
                return error;
            }
        }
        
        #endregion
    }
}
