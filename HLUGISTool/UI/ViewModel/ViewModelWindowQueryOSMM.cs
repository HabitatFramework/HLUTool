// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// Copyright © 2013, 2016 Thames Valley Environmental Records Centre
// Copyright © 2014 Sussex Biodiversity Record Centre
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
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using HLU.Data.Model;
using HLU.Data.Connection;
using HLU.UI.ViewModel;
using HLU.GISApplication;
using HLU.Properties;
using HLU.UI.UserControls;

namespace HLU.UI.ViewModel
{
    //---------------------------------------------------------------------
    // CHANGED: CR49 Process proposed OSMM Updates
    // Functionality to process proposed OSMM Updates.
    //    

    public class OSMMUpdates
    {
        public string ProcessFlag { get; set; }
        public string SpatialFlag { get; set; }
        public string ChangeFlag { get; set; }
        public string Rejected { get; set; }
        public string Ignored { get; set; }
        public string Proposed { get; set; }
        public string Pending { get; set; }
        public string Applied { get; set; }
        public string Total { get; set; }
    }

    class ViewModelWindowQueryOSMM : ViewModelBase, INotifyPropertyChanged
    {
        //public delegate void SelectionChangedEventHandler(object sender, SelectionChangedEventArgs e);

        //public event SelectionChangedEventHandler SelectionChanged;

        ViewModelWindowMain _viewModelMain;

        public static HluDataSet HluDatasetStatic = null;

        #region Fields

        private HluDataSet _hluDataset;
        private DbBase _db;

        private Cursor _cursorType = Cursors.Arrow;

        private ICommand _applyOSMMFilterCommand;
        private ICommand _resetOSMMFilterCommand;
        private ICommand _cancelOSMMFilterCommand;

        private int _osmmUpdatesCountRejected = -1;
        private int _osmmUpdatesCountIgnored = -1;
        private int _osmmUpdatesCountPending = -1;
        private int _osmmUpdatesCountApplied = -1;
        private int _osmmUpdatesCountProposed = -1;

        private int _filterCount = -1;

        private string _osmmUpdatesStatus;

        //private DataRow _osmmUpdatesSelected;

        private HluDataSet.lut_osmm_updates_processRow[] _osmmProcessFlags;
        private HluDataSet.lut_osmm_updates_spatialRow[] _osmmSpatialFlags;
        private HluDataSet.lut_osmm_updates_changeRow[] _osmmChangeFlags;
        private string _osmmProcessFlag;
        private string _osmmSpatialFlag;
        private string _osmmChangeFlag;

        private string _codeAnyRow = Settings.Default.CodeAnyRow;

        private string _displayName = "OSMM Updates Filter";

        #endregion

        #region Constructor

        /// <summary>
        /// Get the default values from settings.
        /// </summary>
        /// <remarks></remarks>
        public ViewModelWindowQueryOSMM(HluDataSet hluDataset, DbBase hluDatabase, ViewModelWindowMain viewModelMain)
        {
            HluDatasetStatic = hluDataset;
            _hluDataset = hluDataset;
            _db = hluDatabase;
            _viewModelMain = viewModelMain;

            // Reset all the filter fields.
            _osmmProcessFlag = _codeAnyRow;
            _osmmSpatialFlag = _codeAnyRow;
            _osmmChangeFlag = _codeAnyRow;

            // Set the default status value
            if (_viewModelMain.OSMMBulkUpdateMode == true)
                _osmmUpdatesStatus = "Pending";
            else
                _osmmUpdatesStatus = "Proposed";

            // Count the incid_osmm_update rows for the initial values.
            CountOSMMUpdates();
            
        }

        public void OSMMUpdatesSelectedRow(OSMMUpdates selectedRow)
        {
            if (selectedRow != null && selectedRow.ChangeFlag != "Total")
            {
                _osmmProcessFlag = selectedRow.ProcessFlag;
                _osmmSpatialFlag = selectedRow.SpatialFlag;
                _osmmChangeFlag = selectedRow.ChangeFlag;

                // Count the incid_osmm_update rows for the initial values.
                CountOSMMUpdates();

                OnPropertyChanged("IncidOSMMUpdatesProcessFlag");
                OnPropertyChanged("IncidOSMMUpdatesSpatialFlag");
                OnPropertyChanged("IncidOSMMUpdatesChangeFlag");
            }
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
            get { return _displayName; }
        }

        #endregion

        #region RequestClose

        // declare the delegate since using non-generic pattern
        public delegate void RequestCloseEventHandler(string processFlag, string spatialFlag, string changeFlag, string status, bool apply);

        // declare the event
        public event RequestCloseEventHandler RequestClose;

        #endregion

        #region Apply Command

        /// <summary>
        /// OSMM Updates Apply Filter command.
        /// </summary>
        public ICommand ApplyOSMMFilterCommand
        {
            get
            {
                if (_applyOSMMFilterCommand == null)
                {
                    Action<object> applyOSMMFilterAction = new Action<object>(this.ApplyOSMMFilterClicked);
                    _applyOSMMFilterCommand = new RelayCommand(applyOSMMFilterAction, param => this.CanApplyOSMMFilter);
                }
                return _applyOSMMFilterCommand;
            }
        }

        /// <summary>
        /// Handles event when Apply button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void ApplyOSMMFilterClicked(object param)
        {
            HluDatasetStatic = null;
            this.RequestClose(IncidOSMMUpdatesProcessFlag, IncidOSMMUpdatesSpatialFlag, IncidOSMMUpdatesChangeFlag, IncidOSMMUpdatesStatus, true);
        }

        public bool CanApplyOSMMFilter
        {
            get
            {
                return (_filterCount > 0);
            }
        }

        #endregion

        #region Reset Command

        /// <summary>
        /// OSMM Updates Reset Filter command.
        /// </summary>
        public ICommand ResetOSMMFilterCommand
        {
            get
            {
                if (_resetOSMMFilterCommand == null)
                {
                    Action<object> resetOSMMFilterAction = new Action<object>(this.ResetOSMMFilterClicked);
                    _resetOSMMFilterCommand = new RelayCommand(resetOSMMFilterAction, param => this.CanResetOSMMFilter);
                }
                return _resetOSMMFilterCommand;
            }
        }

        /// <summary>
        /// Handles event when Reset button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void ResetOSMMFilterClicked(object param)
        {
            // Reset all the filter fields.
            _osmmProcessFlag = _codeAnyRow;
            _osmmSpatialFlag = _codeAnyRow;
            _osmmChangeFlag = _codeAnyRow;

            // Set the default status value
            if (_viewModelMain.OSMMBulkUpdateMode == true)
                _osmmUpdatesStatus = "Pending";
            else
                _osmmUpdatesStatus = "Proposed";

            // Count the incid_osmm_update rows for the initial values.
            CountOSMMUpdates();

            OnPropertyChanged("IncidOSMMUpdatesProcessFlag");
            OnPropertyChanged("IncidOSMMUpdatesSpatialFlag");
            OnPropertyChanged("IncidOSMMUpdatesChangeFlag");
            OnPropertyChanged("IncidOSMMUpdatesStatus");
        }

        public bool CanResetOSMMFilter
        {
            get
            {
                return true;
            }
        }

        #endregion

        #region Cancel Command

        /// <summary>
        /// OSMM Updates Cancel Filter command.
        /// </summary>
        public ICommand CancelOSMMFilterCommand
        {
            get
            {
                if (_cancelOSMMFilterCommand == null)
                {
                    Action<object> cancelOSMMFilterAction = new Action<object>(this.CancelCommandClicked);
                    _cancelOSMMFilterCommand = new RelayCommand(cancelOSMMFilterAction);
                }

                return _cancelOSMMFilterCommand;
            }
        }

        /// <summary>
        /// Handles event when Cancel button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void CancelCommandClicked(object param)
        {
            HluDatasetStatic = null;
            this.RequestClose(null, null, null, null, false);
        }

        #endregion

        public HluDataSet.lut_osmm_updates_processRow[] IncidOSMMUpdatesProcessFlags
        {
            get
            {
                if ((_osmmProcessFlags == null) || (_osmmProcessFlags.Length == 0))
                {
                    _osmmProcessFlags = (from m in _hluDataset.lut_osmm_updates_process
                                         select m).OrderBy(m => m.sort_order).ThenBy(m => m.description).ToArray();
                }

                HluDataSet.lut_osmm_updates_processRow[] osmmProcessFlags;
                osmmProcessFlags = AnyRowOSMMUpdatesProcess(-3).Concat(_osmmProcessFlags).OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();

                return osmmProcessFlags;
            }
            set { }
        }

        public string IncidOSMMUpdatesProcessFlag
        {
            get { return _osmmProcessFlag; }
            set
            {
                _osmmProcessFlag = value;

                // Count the incid_osmm_update rows for the selected flag.
                CountOSMMUpdates();
            }
        }

        public HluDataSet.lut_osmm_updates_spatialRow[] IncidOSMMUpdatesSpatialFlags
        {
            get
            {
                if ((_osmmSpatialFlags == null) || (_osmmSpatialFlags.Length == 0))
                {
                    _osmmSpatialFlags = (from m in _hluDataset.lut_osmm_updates_spatial
                                         select m).OrderBy(m => m.sort_order).ThenBy(m => m.description).ToArray();
                }

                HluDataSet.lut_osmm_updates_spatialRow[] osmmSpatialFlags;
                osmmSpatialFlags = AnyRowOSMMUpdatesSpatial(-3).Concat(_osmmSpatialFlags).OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();

                return osmmSpatialFlags;
            }
            set { }
        }

        public string IncidOSMMUpdatesSpatialFlag
        {
            get { return _osmmSpatialFlag; }
            set
            {
                _osmmSpatialFlag = value;

                // Count the incid_osmm_update rows for the selected flag.
                CountOSMMUpdates();
            }
        }

        public HluDataSet.lut_osmm_updates_changeRow[] IncidOSMMUpdatesChangeFlags
        {
            get
            {
                if ((_osmmChangeFlags == null) || (_osmmChangeFlags.Length == 0))
                {
                    _osmmChangeFlags = (from m in _hluDataset.lut_osmm_updates_change
                                        select m).OrderBy(m => m.sort_order).ThenBy(m => m.description).ToArray();
                }

                HluDataSet.lut_osmm_updates_changeRow[] osmmChangeFlags;
                osmmChangeFlags = AnyRowOSMMUpdatesChange(-3).Concat(_osmmChangeFlags).OrderBy(r => r.sort_order).ThenBy(r => r.description).ToArray();

                return osmmChangeFlags;
            }
            set { }
        }

        public string IncidOSMMUpdatesChangeFlag
        {
            get { return _osmmChangeFlag; }
            set
            {
                _osmmChangeFlag = value;

                // Count the incid_osmm_update rows for the selected flag.
                CountOSMMUpdates();
            }
        }

        /// <summary>
        /// Gets or sets the list of available show OSMM Update options from
        /// the class.
        /// </summary>
        /// <value>
        /// The list of subset update actions.
        /// </value>
        public string[] IncidOSMMUpdatesStatuses
        {
            get
            {
                string[] osmmUpdateStatuses;
                if (_viewModelMain.OSMMBulkUpdateMode == true)
                    osmmUpdateStatuses = new[] { "Pending" };
                else
                    osmmUpdateStatuses = Settings.Default.OSMMUpdatesStatuses.Cast<string>().ToArray();

                return osmmUpdateStatuses;
            }
            set { }
        }

        /// <summary>
        /// Gets or sets the preferred show OSMM Update option.
        /// </summary>
        /// <value>
        /// The preferred show OSMM Update option.
        /// </value>
        public string IncidOSMMUpdatesStatus
        {
            get { return _osmmUpdatesStatus; }
            set
            {
                _osmmUpdatesStatus = value;

                // Count the incid_osmm_update rows for the selected status.
                CountOSMMUpdates();
            }
        }

        private HluDataSet.lut_osmm_updates_processRow[] AnyRowOSMMUpdatesProcess(int sortOrder)
        {
            HluDataSet.lut_osmm_updates_processRow anyRow = _hluDataset.lut_osmm_updates_process.Newlut_osmm_updates_processRow();
            anyRow.code = _codeAnyRow;
            anyRow.sort_order = sortOrder;
            anyRow.description = String.Empty;
            return new HluDataSet.lut_osmm_updates_processRow[] { anyRow };
        }

        private HluDataSet.lut_osmm_updates_spatialRow[] AnyRowOSMMUpdatesSpatial(int sortOrder)
        {
            HluDataSet.lut_osmm_updates_spatialRow anyRow = _hluDataset.lut_osmm_updates_spatial.Newlut_osmm_updates_spatialRow();
            anyRow.code = _codeAnyRow;
            anyRow.sort_order = sortOrder;
            anyRow.description = String.Empty;
            return new HluDataSet.lut_osmm_updates_spatialRow[] { anyRow };
        }

        private HluDataSet.lut_osmm_updates_changeRow[] AnyRowOSMMUpdatesChange(int sortOrder)
        {
            HluDataSet.lut_osmm_updates_changeRow anyRow = _hluDataset.lut_osmm_updates_change.Newlut_osmm_updates_changeRow();
            anyRow.code = _codeAnyRow;
            anyRow.sort_order = sortOrder;
            anyRow.description = String.Empty;
            return new HluDataSet.lut_osmm_updates_changeRow[] { anyRow };
        }

        public string IncidOSMMUpdatesRejectedCount
        {
            get { return String.Format("{0:n0}", _osmmUpdatesCountRejected); }
        }

        public string IncidOSMMUpdatesIgnoredCount
        {
            get { return String.Format("{0:n0}", _osmmUpdatesCountIgnored); }
        }

        public string IncidOSMMUpdatesPendingCount
        {
            get { return String.Format("{0:n0}", _osmmUpdatesCountPending); }
        }

        public string IncidOSMMUpdatesProposedCount
        {
            get { return String.Format("{0:n0}", _osmmUpdatesCountProposed); }
        }

        public string IncidOSMMUpdatesAppliedCount
        {
            get { return String.Format("{0:n0}", _osmmUpdatesCountApplied); }
        }

        /// <summary>
        /// Count the number of OSMM Updates in the database for the selected flags.
        /// </summary>
        public void CountOSMMUpdates()
        {
            // Show the wait cursor whilst loading the values.
            ChangeCursor(Cursors.Wait);

            _osmmUpdatesCountRejected = -1;
            _osmmUpdatesCountIgnored = -1;
            _osmmUpdatesCountApplied = -1;
            _osmmUpdatesCountPending = -1;
            _osmmUpdatesCountProposed = -1;

            StringBuilder whereClause = new StringBuilder();

            if (!String.IsNullOrEmpty(_osmmProcessFlag) && _osmmProcessFlag != _codeAnyRow)
            {
                whereClause.Append(String.Format(" AND {0} = {1}",
                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.process_flagColumn.ColumnName), _osmmProcessFlag));
            }

            if (!String.IsNullOrEmpty(_osmmSpatialFlag) && _osmmSpatialFlag != _codeAnyRow)
            {
                whereClause.Append(String.Format(" AND {0} = {1}",
                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.spatial_flagColumn.ColumnName), _db.QuoteValue(_osmmSpatialFlag)));
            }

            if (!String.IsNullOrEmpty(_osmmChangeFlag) && _osmmChangeFlag != _codeAnyRow)
            {
                whereClause.Append(String.Format(" AND {0} = {1}",
                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.change_flagColumn.ColumnName), _db.QuoteValue(_osmmChangeFlag)));
            }

            // Count the total number of rejected OSMM updates in the database for
            // the selected flags.
            _osmmUpdatesCountRejected = (int)_db.ExecuteScalar(String.Format(
                "SELECT COUNT(*) FROM {0} WHERE {1} = {2}{3}",
                _db.QualifyTableName(_hluDataset.incid_osmm_updates.TableName),
                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.statusColumn.ColumnName),
                -99,
                whereClause),
                _db.Connection.ConnectionTimeout, CommandType.Text);

            // Count the total number of ignored OSMM updates in the database for
            // the selected flags.
            _osmmUpdatesCountIgnored = (int)_db.ExecuteScalar(String.Format(
                "SELECT COUNT(*) FROM {0} WHERE {1} = {2}{3}",
                _db.QualifyTableName(_hluDataset.incid_osmm_updates.TableName),
                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.statusColumn.ColumnName),
                -2,
                whereClause),
                _db.Connection.ConnectionTimeout, CommandType.Text);

            // Count the total number of applied OSMM updates in the database for
            // the selected flags.
            _osmmUpdatesCountApplied = (int)_db.ExecuteScalar(String.Format(
                "SELECT COUNT(*) FROM {0} WHERE {1} = {2}{3}",
                _db.QualifyTableName(_hluDataset.incid_osmm_updates.TableName),
                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.statusColumn.ColumnName),
                -1,
                whereClause),
                _db.Connection.ConnectionTimeout, CommandType.Text);

            // Count the total number of pending OSMM updates in the database for
            // the selected flags.
            _osmmUpdatesCountPending = (int)_db.ExecuteScalar(String.Format(
                "SELECT COUNT(*) FROM {0} WHERE {1} = {2}{3}",
                _db.QualifyTableName(_hluDataset.incid_osmm_updates.TableName),
                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.statusColumn.ColumnName),
                0,
                whereClause),
                _db.Connection.ConnectionTimeout, CommandType.Text);

            // Count the total number of proposed OSMM updates in the database for
            // the selected flags.
            _osmmUpdatesCountProposed = (int)_db.ExecuteScalar(String.Format(
                "SELECT COUNT(*) FROM {0} WHERE {1} > {2}{3}",
                _db.QualifyTableName(_hluDataset.incid_osmm_updates.TableName),
                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.statusColumn.ColumnName),
                0,
                whereClause),
                _db.Connection.ConnectionTimeout, CommandType.Text);

            if (!String.IsNullOrEmpty(_osmmUpdatesStatus) && _osmmUpdatesStatus != _codeAnyRow)
            {
                switch (_osmmUpdatesStatus)
                {
                    case "Rejected":
                        _filterCount = _osmmUpdatesCountRejected;
                        break;
                    case "Ignored":
                        _filterCount = _osmmUpdatesCountIgnored;
                        break;
                    case "Proposed":
                        _filterCount = _osmmUpdatesCountProposed;
                        break;
                    case "Pending":
                        _filterCount = _osmmUpdatesCountPending;
                        break;
                }
            }

            OnPropertyChanged("IncidOSMMUpdatesRejectedCount");
            OnPropertyChanged("IncidOSMMUpdatesIgnoredCount");
            OnPropertyChanged("IncidOSMMUpdatesAppliedCount");
            OnPropertyChanged("IncidOSMMUpdatesPendingCount");
            OnPropertyChanged("IncidOSMMUpdatesProposedCount");

            OnPropertyChanged("CanApplyOSMMFilter");

            // Reset the cursor back to normal.
            ChangeCursor(Cursors.Arrow);

        }

        /// <summary>
        /// Count the total number of OSMM Updates in the database.
        /// </summary>
        //public DataTable OSMMUpdatesSummary
        public List<OSMMUpdates> OSMMUpdatesSummary
        {
            get
            {
                int rejectedCount = 0;
                int ignoredCount = 0;
                int proposedCount = 0;
                int pendingCount = 0;
                int appliedCount = 0;
                int allCount = 0;
                int rejectedTotal = 0;
                int ignoredTotal = 0;
                int proposedTotal = 0;
                int pendingTotal = 0;
                int appliedTotal = 0;
                int allTotal = 0;

                // Show the wait cursor whilst loading the values.
                ChangeCursor(Cursors.Wait);

                // Define a new data table to hold the results.
                List<OSMMUpdates> dataTable = new List<OSMMUpdates>();

                // Create a data reader to retrieve the rows for
                // the required column.
                IDataReader dataReader = null;

                try
                {
                    // Load the data reader to retrieve the rows for
                    // the required column.
                    string sqlColumns = String.Format("{0}, {1}, {2}, {3}, COUNT(*) As RecCount",
                        _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.process_flagColumn.ColumnName),
                        _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.spatial_flagColumn.ColumnName),
                        _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.change_flagColumn.ColumnName),
                        _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.statusColumn.ColumnName));
                    string sqlGroupBy = String.Format("{0}, {1}, {2}, {3}",
                        _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.process_flagColumn.ColumnName),
                        _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.spatial_flagColumn.ColumnName),
                        _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.change_flagColumn.ColumnName),
                        _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.statusColumn.ColumnName));
                    dataReader = _db.ExecuteReader(String.Format(
                        "SELECT DISTINCT {0} FROM {1} GROUP BY {3} ORDER BY {3}",
                        sqlColumns,
                        _db.QualifyTableName(_hluDataset.incid_osmm_updates.TableName),
                        _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.statusColumn.ColumnName),
                        sqlGroupBy),
                        _db.Connection.ConnectionTimeout, CommandType.Text);

                    if (dataReader == null) throw new Exception(String.Format("Error reading values from {0}", _hluDataset.incid_osmm_updates.TableName));

                    string processFlag, lastProcessFlag = null;
                    string spatialFlag, lastSpatialFlag = null;
                    string changeFlag, lastChangeFlag = null;
                    int status;
                    int recs;
                    OSMMUpdates dataRow;

                    // Load the list with the results.
                    while (dataReader.Read())
                    {
                        processFlag = dataReader.GetValue(0).ToString();
                        spatialFlag = dataReader.GetValue(1).ToString();
                        changeFlag = dataReader.GetValue(2).ToString();
                        status = (int)dataReader.GetValue(3);
                        recs = (int)dataReader.GetValue(4);

                        if (lastProcessFlag == null)
                            lastProcessFlag = processFlag;
                        if (lastSpatialFlag == null)
                            lastSpatialFlag = spatialFlag;
                        if (lastChangeFlag == null)
                            lastChangeFlag = changeFlag;

                        // If this is a different group.
                        if (processFlag != lastProcessFlag || spatialFlag != lastSpatialFlag || changeFlag != lastChangeFlag)
                        {
                            // Add the results as a new row.
                            dataRow = new OSMMUpdates();
                            dataRow.ProcessFlag = lastProcessFlag;
                            dataRow.SpatialFlag = lastSpatialFlag;
                            dataRow.ChangeFlag = lastChangeFlag;
                            dataRow.Rejected = String.Format("{0:n0}", rejectedCount);
                            dataRow.Ignored = String.Format("{0:n0}", ignoredCount);
                            dataRow.Proposed = String.Format("{0:n0}", proposedCount);
                            dataRow.Pending = String.Format("{0:n0}", pendingCount);
                            dataRow.Applied = String.Format("{0:n0}", appliedCount);
                            dataRow.Total = String.Format("{0:n0}", allCount);
                            dataTable.Add(dataRow);

                            // Update the totals.
                            rejectedTotal = rejectedTotal + rejectedCount;
                            ignoredTotal = ignoredTotal + ignoredCount;
                            proposedTotal = proposedTotal + proposedCount;
                            pendingTotal = pendingTotal + pendingCount;
                            appliedTotal = appliedTotal + appliedCount;
                            allTotal = allTotal + allCount;

                            // Reset the counts.
                            rejectedCount = 0;
                            ignoredCount = 0;
                            proposedCount = 0;
                            pendingCount = 0;
                            appliedCount = 0;
                            allCount = 0;

                            // Save the last group values;
                            lastProcessFlag = processFlag;
                            lastSpatialFlag = spatialFlag;
                            lastChangeFlag = changeFlag;
                        }

                        // Update the counts.
                        if (status == -99)
                            rejectedCount = recs;
                        if (status == -2)
                            ignoredCount = recs;
                        if (status == 0)
                            pendingCount = recs;
                        if (status == -1)
                            appliedCount = recs;
                        if (status > 0)
                            proposedCount = proposedCount + recs;
                        allCount = allCount + recs;

                    }

                    // Add the last results as a new row.
                    dataRow = new OSMMUpdates();
                    dataRow.ProcessFlag = lastProcessFlag;
                    dataRow.SpatialFlag = lastSpatialFlag;
                    dataRow.ChangeFlag = lastChangeFlag;
                    dataRow.Rejected = String.Format("{0:n0}", rejectedCount);
                    dataRow.Ignored = String.Format("{0:n0}", ignoredCount);
                    dataRow.Proposed = String.Format("{0:n0}", proposedCount);
                    dataRow.Pending = String.Format("{0:n0}", pendingCount);
                    dataRow.Applied = String.Format("{0:n0}", appliedCount);
                    dataRow.Total = String.Format("{0:n0}", allCount);
                    dataTable.Add(dataRow);

                    // Update the totals.
                    rejectedTotal = rejectedTotal + rejectedCount;
                    ignoredTotal = ignoredTotal + ignoredCount;
                    proposedTotal = proposedTotal + proposedCount;
                    pendingTotal = pendingTotal + pendingCount;
                    appliedTotal = appliedTotal + appliedCount;
                    allTotal = allTotal + allCount;

                    // Add the totals as a new row.
                    dataRow = new OSMMUpdates();
                    dataRow.ProcessFlag = "";
                    dataRow.SpatialFlag = "";
                    dataRow.ChangeFlag = "Total";
                    dataRow.Rejected = String.Format("{0:n0}", rejectedTotal);
                    dataRow.Ignored = String.Format("{0:n0}", ignoredTotal);
                    dataRow.Proposed = String.Format("{0:n0}", proposedTotal);
                    dataRow.Pending = String.Format("{0:n0}", pendingTotal);
                    dataRow.Applied = String.Format("{0:n0}", appliedTotal);
                    dataRow.Total = String.Format("{0:n0}", allTotal);
                    dataTable.Add(dataRow);

                    // Reset the cursor back to normal.
                    ChangeCursor(Cursors.Arrow);

                    return dataTable;
                }
                catch (Exception ex)
                {
                    // Reset the cursor back to normal.
                    ChangeCursor(Cursors.Arrow);

                    MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);

                    return null;
                }
                finally
                {
                    // Close the data reader.
                    if (!dataReader.IsClosed)
                        dataReader.Close();
                }
            }

        }

        ///// <summary>
        ///// Count the total number of OSMM Updates in the database.
        ///// </summary>
        ////public DataTable OSMMUpdatesSummary
        //public List<test> OSMMUpdatesSummary
        //{
        //    get
        //    {
        //        int rejectedCount = 0;
        //        int ignoredCount = 0;
        //        int proposedCount = 0;
        //        int pendingCount = 0;
        //        int appliedCount = 0;
        //        int allCount = 0;
        //        int rejectedTotal = 0;
        //        int ignoredTotal = 0;
        //        int proposedTotal = 0;
        //        int pendingTotal = 0;
        //        int appliedTotal = 0;
        //        int allTotal = 0;

        //        // Show the wait cursor whilst loading the values.
        //        ChangeCursor(Cursors.Wait);

        //        // Define a new data table to hold the results.
        //        DataTable dataTable = new DataTable();
        //        DataColumn dataColumn = new DataColumn("ProcessFlag");
        //        dataTable.Columns.Add(dataColumn);
        //        dataColumn = new DataColumn("SpatialFlag");
        //        dataTable.Columns.Add(dataColumn);
        //        dataColumn = new DataColumn("ChangeFlag");
        //        dataTable.Columns.Add(dataColumn);
        //        dataColumn = new DataColumn("Rejected");
        //        dataTable.Columns.Add(dataColumn);
        //        dataColumn = new DataColumn("Ignored");
        //        dataTable.Columns.Add(dataColumn);
        //        dataColumn = new DataColumn("Proposed");
        //        dataTable.Columns.Add(dataColumn);
        //        dataColumn = new DataColumn("Pending");
        //        dataTable.Columns.Add(dataColumn);
        //        dataColumn = new DataColumn("Applied");
        //        dataTable.Columns.Add(dataColumn);
        //        dataColumn = new DataColumn("Total");
        //        dataTable.Columns.Add(dataColumn);
        //        DataRow dataRow = null;

        //        // Create a data reader to retrieve the rows for
        //        // the required column.
        //        IDataReader dataReader = null;

        //        try
        //        {
        //            // Load the data reader to retrieve the rows for
        //            // the required column.
        //            string sqlColumns = String.Format("{0}, {1}, {2}, {3}, COUNT(*) As RecCount",
        //                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.process_flagColumn.ColumnName),
        //                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.spatial_flagColumn.ColumnName),
        //                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.change_flagColumn.ColumnName),
        //                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.statusColumn.ColumnName));
        //            string sqlGroupBy = String.Format("{0}, {1}, {2}, {3}",
        //                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.process_flagColumn.ColumnName),
        //                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.spatial_flagColumn.ColumnName),
        //                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.change_flagColumn.ColumnName),
        //                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.statusColumn.ColumnName));
        //            dataReader = _db.ExecuteReader(String.Format(
        //                "SELECT DISTINCT {0} FROM {1} GROUP BY {3} ORDER BY {3}",
        //                sqlColumns,
        //                _db.QualifyTableName(_hluDataset.incid_osmm_updates.TableName),
        //                _db.QuoteIdentifier(_hluDataset.incid_osmm_updates.statusColumn.ColumnName),
        //                sqlGroupBy),
        //                _db.Connection.ConnectionTimeout, CommandType.Text);

        //            if (dataReader == null) throw new Exception(String.Format("Error reading values from {0}", _hluDataset.incid_osmm_updates.TableName));

        //            string processFlag, lastProcessFlag = null;
        //            string spatialFlag, lastSpatialFlag = null;
        //            string changeFlag, lastChangeFlag = null;
        //            int status;
        //            int recs;

        //            // Load the list with the results.
        //            while (dataReader.Read())
        //            {
        //                processFlag = dataReader.GetValue(0).ToString();
        //                spatialFlag = dataReader.GetValue(1).ToString();
        //                changeFlag = dataReader.GetValue(2).ToString();
        //                status = (int)dataReader.GetValue(3);
        //                recs = (int)dataReader.GetValue(4);

        //                if (lastProcessFlag == null)
        //                    lastProcessFlag = processFlag;
        //                if (lastSpatialFlag == null)
        //                    lastSpatialFlag = spatialFlag;
        //                if (lastChangeFlag == null)
        //                    lastChangeFlag = changeFlag;

        //                // If this is a different group.
        //                if (processFlag != lastProcessFlag || spatialFlag != lastSpatialFlag || changeFlag != lastChangeFlag)
        //                {
        //                    // Add the results as a new row.
        //                    dataRow = dataTable.NewRow();
        //                    dataRow["ProcessFlag"] = lastProcessFlag;
        //                    dataRow["SpatialFlag"] = lastSpatialFlag;
        //                    dataRow["ChangeFlag"] = lastChangeFlag;
        //                    dataRow["Rejected"] = String.Format("{0:n0}", rejectedCount);
        //                    dataRow["Ignored"] = String.Format("{0:n0}", ignoredCount);
        //                    dataRow["Proposed"] = String.Format("{0:n0}", proposedCount);
        //                    dataRow["Pending"] = String.Format("{0:n0}", pendingCount);
        //                    dataRow["Applied"] = String.Format("{0:n0}", appliedCount);
        //                    dataRow["Total"] = String.Format("{0:n0}", allCount);
        //                    dataTable.Rows.Add(dataRow);

        //                    // Update the totals.
        //                    rejectedTotal = rejectedTotal + rejectedCount;
        //                    ignoredTotal = ignoredTotal + ignoredCount;
        //                    proposedTotal = proposedTotal + proposedCount;
        //                    pendingTotal = pendingTotal + pendingCount;
        //                    appliedTotal = appliedTotal + appliedCount;
        //                    allTotal = allTotal + allCount;

        //                    // Reset the counts.
        //                    rejectedCount = 0;
        //                    ignoredCount = 0;
        //                    proposedCount = 0;
        //                    pendingCount = 0;
        //                    appliedCount = 0;
        //                    allCount = 0;

        //                    // Save the last group values;
        //                    lastProcessFlag = processFlag;
        //                    lastSpatialFlag = spatialFlag;
        //                    lastChangeFlag = changeFlag;
        //                }

        //                // Update the counts.
        //                if (status == -99)
        //                    rejectedCount = recs;
        //                if (status == -2)
        //                    ignoredCount = recs;
        //                if (status == 0)
        //                    pendingCount = recs;
        //                if (status == -1)
        //                    appliedCount = recs;
        //                if (status > 0)
        //                    proposedCount = proposedCount + recs;
        //                allCount = allCount + recs;

        //            }

        //            // Add the last results as a new row.
        //            dataRow = dataTable.NewRow();
        //            dataRow["ProcessFlag"] = lastProcessFlag;
        //            dataRow["SpatialFlag"] = lastSpatialFlag;
        //            dataRow["ChangeFlag"] = lastChangeFlag;
        //            dataRow["Rejected"] = String.Format("{0:n0}", rejectedCount);
        //            dataRow["Ignored"] = String.Format("{0:n0}", ignoredCount);
        //            dataRow["Proposed"] = String.Format("{0:n0}", proposedCount);
        //            dataRow["Pending"] = String.Format("{0:n0}", pendingCount);
        //            dataRow["Applied"] = String.Format("{0:n0}", appliedCount);
        //            dataRow["Total"] = String.Format("{0:n0}", allCount);
        //            dataTable.Rows.Add(dataRow);

        //            // Update the totals.
        //            rejectedTotal = rejectedTotal + rejectedCount;
        //            ignoredTotal = ignoredTotal + ignoredCount;
        //            proposedTotal = proposedTotal + proposedCount;
        //            pendingTotal = pendingTotal + pendingCount;
        //            appliedTotal = appliedTotal + appliedCount;
        //            allTotal = allTotal + allCount;

        //            // Add the totals as a new row.
        //            dataRow = dataTable.NewRow();
        //            dataRow["ProcessFlag"] = "";
        //            dataRow["SpatialFlag"] = "";
        //            dataRow["ChangeFlag"] = "Total";
        //            dataRow["Rejected"] = String.Format("{0:n0}", rejectedTotal);
        //            dataRow["Ignored"] = String.Format("{0:n0}", ignoredTotal);
        //            dataRow["Proposed"] = String.Format("{0:n0}", proposedTotal);
        //            dataRow["Pending"] = String.Format("{0:n0}", pendingTotal);
        //            dataRow["Applied"] = String.Format("{0:n0}", appliedTotal);
        //            dataRow["Total"] = String.Format("{0:n0}", allTotal);
        //            dataTable.Rows.Add(dataRow);

        //            // Reset the cursor back to normal.
        //            ChangeCursor(Cursors.Arrow);

        //            //return dataTable;

        //            test test1 = new test();
        //            test1.processflag = "A";
        //            test1.spatialflag = "1";

        //            List<test> test2 = new List<test> { test1 };

        //            return test2;
        //        }
        //        catch (Exception ex)
        //        {
        //            // Reset the cursor back to normal.
        //            ChangeCursor(Cursors.Arrow);

        //            MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);

        //            return null;
        //        }
        //        finally
        //        {
        //            // Close the data reader.
        //            if (!dataReader.IsClosed)
        //                dataReader.Close();
        //        }
        //    }

        //}

        #region Cursor

        /// <summary>
        /// Gets the cursor type to use when the cursor is over the window.
        /// </summary>
        /// <value>
        /// The window cursor type.
        /// </value>
        public Cursor WindowCursor { get { return _cursorType; } }

        public void ChangeCursor(Cursor cursorType)
        {
            _cursorType = cursorType;
            OnPropertyChanged("WindowCursor");
            if (cursorType == Cursors.Wait)
                DispatcherHelper.DoEvents();
        }

        #endregion

    }
    //---------------------------------------------------------------------
}
