using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;
using HLU.Data;
using HLU.Data.Model;

namespace HLU.UI.ViewModel
{
    class ViewModelQueryBuilder : ViewModelBase, IDataErrorInfo
    {
        public static HluDataSet HluDatasetStatic = null;

        #region Fields

        private ICommand _okCommand;
        private ICommand _cancelCommand;
        private string _displayName = "Query Builder";
        private HluDataSet _hluDataset;
        private ObservableCollection<QueryItem> _queryItems;

        #endregion

        #region Constructor

        public ViewModelQueryBuilder(HluDataSet hluDataset)
        {
            HluDatasetStatic = hluDataset;
            _hluDataset = hluDataset;
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
        public delegate void RequestCloseEventHandler(ObservableCollection<QueryItem> queryItems);

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
            HluDatasetStatic = null;
            this.RequestClose(_queryItems);
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
                return _queryItems != null && _queryItems.Count > 0 &&
                    _queryItems.Count(q => q.IsComplete) == _queryItems.Count; 
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
            HluDatasetStatic = null;
            this.RequestClose(null);
        }

        #endregion

        #region Query Items

        public ObservableCollection<QueryItem> QueryItems
        {
            get
            {
                if (_queryItems == null)
                {
                    QueryItem qryItem = new QueryItem();
                    List<QueryItem> qryItems = new List<QueryItem>();
                    qryItems.Add(qryItem);
                    _queryItems = new ObservableCollection<QueryItem>(qryItems);
                }
                return _queryItems;
            }
            set { _queryItems = value; }
        }

        #endregion

        #region IDataErrorInfo Members

        public string Error
        {
            get
            {
                string error = null;
                return error;
            }
        }

        public string this[string columnName]
        {
            get 
            {
                string error = null;

                // dirty commands registered with CommandManager so they are queried to see if they can execute now
                CommandManager.InvalidateRequerySuggested();

                return error;
            }
        }

        #endregion
    }
}
