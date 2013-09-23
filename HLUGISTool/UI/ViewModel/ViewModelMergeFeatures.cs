using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Windows.Input;
using HLU.Data.Model;
using HLU.GISApplication;

namespace HLU.UI.ViewModel
{
    class ViewModelMergeFeatures<T, R> : ViewModelBase, IDataErrorInfo
        where T : DataTable
        where R : DataRow
    {
        #region Fields

        private ICommand _okCommand;
        private ICommand _cancelCommand;
        private ICommand _flashFeatureCommand;
        private string _displayName = "Merge Features";
        private T _selectedFeatures;
        private R _resultFeature;
        private HluDataSet.incid_mm_polygonsRow[] _childRows;
        private HluDataSet.incid_mm_polygonsRow[] _currChildRows;
        private int _incidOrdinal;
        private int _selectedIndex = -1;
        private int[] _keyOrdinals;
        private GISApp _gisApp;

        #endregion

        #region Constructor

        public ViewModelMergeFeatures(T selectedFeatures, int[] keyOrdinals, int incidOrdinal, 
            HluDataSet.incid_mm_polygonsRow[] childRows, GISApp gisApp)
        {
            _selectedFeatures = selectedFeatures;
            _keyOrdinals = keyOrdinals;
            _childRows = childRows;
            _incidOrdinal = incidOrdinal;
            _gisApp = gisApp;
        }

        #endregion

        #region ViewModelBase Members

        public override string DisplayName
        {
            get { return _displayName; }
            set { _displayName = value; }
        }

        public override string WindowTitle { get { return DisplayName; } }

        #endregion

        #region RequestClose

        public delegate void RequestCloseEventHandler(int resultFeatureIndex);

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
            this.RequestClose(_selectedIndex);
        }

        private bool CanOk { get { return String.IsNullOrEmpty(this.Error); } }

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
            this.RequestClose(-1);
        }

        #endregion

        #region Flash Feature Command

        public ICommand FlashFeatureCommand
        {
            get
            {
                if (_flashFeatureCommand == null)
                {
                    Action<object> flashFeatureAction = new Action<object>(this.FlashFeature);
                    _flashFeatureCommand = new RelayCommand(flashFeatureAction, param => this.CanFlashFeature);
                }

                return _flashFeatureCommand;
            }
        }

        private bool CanFlashFeature
        {
            get 
            { 
                return _resultFeature != null && (_resultFeature is HluDataSet.incid_mm_polygonsRow ||
                    ((_currChildRows != null) && (_currChildRows.Length > 0))); 
            }
        }

        private void FlashFeature(object param)
        {
            if (_resultFeature == null) return;

            if (_resultFeature is HluDataSet.incid_mm_polygonsRow)
            {
                List<List<SqlFilterCondition>> whereClause = 
                    ViewModelWindowMainHelpers.GisSelectionToWhereClause(new R[] { _resultFeature },
                    _keyOrdinals, 10, _selectedFeatures);
                if (whereClause.Count == 1)
                    _gisApp.FlashSelectedFeature(whereClause[0]);
            }
            else if ((_currChildRows != null) && (_currChildRows.Length > 0))
            {
                foreach (HluDataSet.incid_mm_polygonsRow r in _currChildRows)
                {
                    List<List<SqlFilterCondition>> whereClause =
                        ViewModelWindowMainHelpers.GisSelectionToWhereClause(new HluDataSet.incid_mm_polygonsRow[] { r },
                        _keyOrdinals, 100, _selectedFeatures);
                    if (whereClause.Count == 1)
                        _gisApp.FlashSelectedFeature(whereClause[0]);
                }
            }
        }

        #endregion

        #region Merge Features

        public T MergeFeatures
        {
            get { return _selectedFeatures; }
        }

        public int SelectedIndex
        {
            get { return _selectedIndex; }
            set 
            {
                _selectedIndex = value;
                _resultFeature = (R)_selectedFeatures.Rows[_selectedIndex];

                if ((_resultFeature is HluDataSet.incidRow) && (_childRows != null))
                    _currChildRows = _childRows.Where(r => r.incid == _resultFeature.Field<string>(_incidOrdinal)).ToArray();

                FlashFeature(null);
            }
        }

        #endregion

        #region IDataErrorInfo Members

        public string Error
        {
            get 
            {
                string error = String.Empty;

                if ((_resultFeature == null) || (_selectedIndex < 0) || (_selectedIndex >= _selectedFeatures.Rows.Count))
                    error = "Please select the feature whose attributes will be retained.";

                return error;
            }
        }

        public string this[string columnName]
        {
            get
            {
                string error = String.Empty;

                switch (columnName)
                {
                    case "SelectedIndex":
                        if ((_selectedIndex < 0) || (_selectedIndex >= _selectedFeatures.Rows.Count))
                            error = "Please select the feature whose attributes will be retained.";
                        break;
                    case "ResultFeature":
                        if (_resultFeature == null)
                            error = "Please select the feature whose attributes will be retained.";
                        break;
                }

                // dirty commands registered with CommandManager so they are queried to see if they can execute now
                CommandManager.InvalidateRequerySuggested();
                
                return error;
            }
        }

        #endregion
    }
}
