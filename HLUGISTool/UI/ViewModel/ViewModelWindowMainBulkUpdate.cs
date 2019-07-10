// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// Copyright © 2013 Thames Valley Environmental Records Centre
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
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Input;
using HLU.Data;
using HLU.Data.Model;
using HLU.Data.Model.HluDataSetTableAdapters;
using HLU.GISApplication;
using HLU.Properties;
using HLU.UI.View;
using HLU.Date;

namespace HLU.UI.ViewModel
{
    //---------------------------------------------------------------------
    // CHANGED: CR49 Process proposed OSMM Updates
    // Functionality for applying pending OSMM Bulk Updates.
    //    
    /// <summary>
    /// An enumeration of the different options for when
    /// to delete multiplex codes.
    /// </summary>
    public enum DeleteMultiplexCodesAction
    {
        All,
        Invalid,
        None
    };
    //---------------------------------------------------------------------

    class ViewModelWindowMainBulkUpdate
    {
        private ViewModelWindowMain _viewModelMain;
        private WindowBulkUpdate _windowBulkUpdate;
        private ViewModelBulkUpdate _viewModelBulkUpdate;

        private HluDataSet.incid_ihs_matrixDataTable _ihsMatrixTable = new HluDataSet.incid_ihs_matrixDataTable();
        private HluDataSet.incid_ihs_formationDataTable _ihsFormationTable = new HluDataSet.incid_ihs_formationDataTable();
        private HluDataSet.incid_ihs_managementDataTable _ihsManagementTable = new HluDataSet.incid_ihs_managementDataTable();
        private HluDataSet.incid_ihs_complexDataTable _ihsComplexTable = new HluDataSet.incid_ihs_complexDataTable();

        private bool _bulkDeleteOrphanBapHabitats;
        private bool _bulkDeletePotentialBapHabitats;
        private int _bulkDeleteMultiplexCodes;
        private bool _bulkCreateHistory;
        private string _bulkDeterminationQuality;
        private string _bulkInterpretationQuality;

        private bool _osmmBulkUpdateMode;

        //private DataTable _incidSelectionBackup;

        #region #ctor

        public ViewModelWindowMainBulkUpdate(ViewModelWindowMain viewModelMain)
        {
            _viewModelMain = viewModelMain;
        }

        #endregion

        #region Bulk Update

        /// <summary>
        /// Starts the bulk update mode.
        /// </summary>
        public void StartBulkUpdate(bool osmmBulkUpdateMode)
        {
            //---------------------------------------------------------------------
            // CHANGED: CR49 Process proposed OSMM Updates
            // Functionality for applying pending OSMM Bulk Updates.
            //    
            // Store whether we are starting in OSMM Bulk update
            // mode or just standard Bulk Update mode
            _osmmBulkUpdateMode = osmmBulkUpdateMode;

            // Update the modes in the main view model
            _viewModelMain.BulkUpdateMode = true;
            if (_osmmBulkUpdateMode == true)
                _viewModelMain.OSMMBulkUpdateMode = true;

            // Clear all the form fields (except the habitat class
            // and habitat type).
            _viewModelMain.ClearForm();

            // Select another tab if the currently selected tab
            // will be disabled.
            if (_osmmBulkUpdateMode == true)
            {
                // Disable the IHS tab
                _viewModelMain.TabItemIhsEnabled = false;

                // Clear the selection (filter).
                _viewModelMain.IncidSelection = null;

                // If the ihs or history tab is currently selected then
                // select the details tab
                if (_viewModelMain.TabItemSelected == 0 || _viewModelMain.TabItemSelected == 3)
                    _viewModelMain.TabItemSelected = 1;
            }
            else
            {
                // If the history tab is currently selected then
                // select the ihs tab
                if (_viewModelMain.TabItemSelected == 3)
                    _viewModelMain.TabItemSelected = 0;
            }
            //---------------------------------------------------------------------

            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            // Clear any interface warning and error messages
            _viewModelMain.ResetWarningsErrors();
            _viewModelMain.RefreshAll();

            //DispatcherHelper.DoEvents();
            //---------------------------------------------------------------------

            if (_osmmBulkUpdateMode == true)
            {
                // Reset the database counts
                _viewModelMain.NumIncidSelectedDB = 0;
                _viewModelMain.NumToidSelectedDB = 0;
                _viewModelMain.NumFragmentsSelectedDB = 0;

                // Open the OSMM Update filter
                _viewModelMain.OpenWindowQueryOSMM(true);
            }

        }

        //---------------------------------------------------------------------
        // FIX: 078 Bulk update overhaul/improvements.
        // 
        /// <summary>
        /// Displays the bulk update window.
        /// </summary>
        public void BulkUpdate()
        {
            // Backup the current selection (filter).
            //_incidSelectionBackup = _viewModelMain.IncidSelection;

            // Get the default settings
            bool deleteOrphanBapHabitats = Settings.Default.BulkUpdateDeleteOrphanBapHabitats;
            bool deletePotentialBapHabitats = Settings.Default.BulkUpdateDeletePotentialBapHabitats;
            int deleteMultiplexCodes = Settings.Default.BulkUpdateDeleteMultiplexCodes;
            bool createHistory = Settings.Default.BulkUpdateCreateHistoryRecords;
            string determinationQuality = Settings.Default.BulkUpdateDeterminationQuality;
            string interpretationQuality = Settings.Default.BulkUpdateInterpretationQuality;

            // Set any mandatory options
            if (_osmmBulkUpdateMode == true)
            {
                deleteMultiplexCodes = 0;
                createHistory = true;
            }

            _windowBulkUpdate = new WindowBulkUpdate();
            _windowBulkUpdate.Owner = App.Current.MainWindow;
            _windowBulkUpdate.WindowStartupLocation = WindowStartupLocation.CenterScreen;

            // Count the non-blank rows from the user interface
            int sourceCount = (from nr in _viewModelMain.IncidSourcesRows
                           where nr.source_id != Int32.MinValue
                           select nr).Count();

            // Display the bulk update interface to prompt the user
            // to select the options they want to use.
            _viewModelBulkUpdate = new ViewModelBulkUpdate(_viewModelMain, _osmmBulkUpdateMode,
                deleteOrphanBapHabitats, deletePotentialBapHabitats, sourceCount, deleteMultiplexCodes,
                createHistory, determinationQuality, interpretationQuality);
            if (_osmmBulkUpdateMode == true)
                _viewModelBulkUpdate.DisplayName = "OSMM Bulk Update";
            else
                _viewModelBulkUpdate.DisplayName = "Bulk Update";
            _viewModelBulkUpdate.RequestClose +=
                new ViewModelBulkUpdate.RequestCloseEventHandler(_viewModelBulkUpdate_RequestClose);

            _windowBulkUpdate.DataContext = _viewModelBulkUpdate;

            _windowBulkUpdate.ShowDialog();
        }

        private void _viewModelBulkUpdate_RequestClose(bool apply,
            bool bulkDeleteOrphanBapHabitats,
            bool bulkDeletePotentialBapHabitats,
            int bulkDeleteMultiplexCodes,
            bool bulkCreateHistory,
            string bulkDeterminationQuality,
            string bulkInterpretationQuality)
        {
            _viewModelBulkUpdate.RequestClose -= _viewModelBulkUpdate_RequestClose;
            _windowBulkUpdate.Close();

            // If the user selected to apply the update then
            // perform the update using the options set.
            if (apply == true)
            {
                //DispatcherHelper.DoEvents();

                // Set the options for processing the bulk update
                _bulkDeleteOrphanBapHabitats = bulkDeleteOrphanBapHabitats;
                _bulkDeletePotentialBapHabitats = bulkDeletePotentialBapHabitats;
                _bulkDeleteMultiplexCodes = bulkDeleteMultiplexCodes;
                _bulkCreateHistory = bulkCreateHistory;
                _bulkDeterminationQuality = bulkDeterminationQuality;
                _bulkInterpretationQuality = bulkInterpretationQuality;

                // Apply the bulk update
                if (_osmmBulkUpdateMode == true)
                    ApplyOSMMBulkUpdate();
                else
                    ApplyBulkUpdate();
            }
        }
        //---------------------------------------------------------------------

        /// <summary>
        /// Applies the bulk update.
        /// </summary>
        public void ApplyBulkUpdate()
        {
            _viewModelMain.ChangeCursor(Cursors.Wait, "Bulk updating ...");

            _viewModelMain.DataBase.BeginTransaction(true, IsolationLevel.ReadCommitted);

            try
            {
                // Build a blank WHERE clause based on incid column
                string incidWhereClause = String.Format(" WHERE {0} = {1}",
                    _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid.incidColumn.ColumnName), 
                    _viewModelMain.DataBase.QuoteValue("{0}"));

                // Build the SELECT statement based on the incid where clause
                string selectCommandIncid = String.Format("SELECT {0} FROM {1}{2}",
                    _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid.ihs_habitatColumn.ColumnName),
                    _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid.TableName), incidWhereClause);

                // Update the last modified date & user fields and the ihs version on
                // the incid table
                //ViewModelWindowMainUpdate.IncidCurrentRowDerivedValuesUpdate(_viewModelMain);
                DateTime currDtTm = DateTime.Now;
                DateTime nowDtTm = new DateTime(currDtTm.Year, currDtTm.Month, currDtTm.Day, currDtTm.Hour, currDtTm.Minute, currDtTm.Second, DateTimeKind.Local);

                _viewModelMain.IncidCurrentRow.last_modified_date = nowDtTm;
                _viewModelMain.IncidCurrentRow.last_modified_user_id = _viewModelMain.UserID;
                _viewModelMain.IncidCurrentRow.ihs_version = _viewModelMain.IhsVersion;
                _viewModelMain.IncidCurrentRow.ihs_habitat = _viewModelMain.IncidIhsHabitat;

                // Build a collection of the updated columns in the incid table
                var incidUpdateVals = _viewModelMain.HluDataset.incid.Columns.Cast<DataColumn>()
                    .Where(c => c.Ordinal != _viewModelMain.HluDataset.incid.incidColumn.Ordinal &&
                        !_viewModelMain.IncidCurrentRow.IsNull(c.Ordinal));

                // Build an UPDATE statement for the incid table
                string updateCommandIncid = incidUpdateVals.Count() == 0 ? String.Empty :
                    new StringBuilder(String.Format("UPDATE {0} SET ",
                    _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid.TableName)))
                    .Append(_viewModelMain.HluDataset.incid.Columns.Cast<DataColumn>()
                    .Where(c => c.Ordinal != _viewModelMain.HluDataset.incid.incidColumn.Ordinal && 
                        !_viewModelMain.IncidCurrentRow.IsNull(c.Ordinal))
                    .Aggregate(new StringBuilder(), (sb, c) => sb.Append(String.Format(", {0} = {1}",
                        _viewModelMain.DataBase.QuoteIdentifier(c.ColumnName), 
                        _viewModelMain.DataBase.QuoteValue(_viewModelMain.IncidCurrentRow[c.Ordinal]))))
                        .Remove(0, 2)).Append(incidWhereClause).ToString();

                // Get the column ordinal for the incid column on the incid table
                int incidOrdinal = 
                    _viewModelMain.IncidSelection.Columns[_viewModelMain.HluDataset.incid.incidColumn.ColumnName].Ordinal;

                List<string> ihsMultiplexDelStatements = new List<string>();
                // If all multiplex codes are to be deleted
                if (_bulkDeleteMultiplexCodes == (int)DeleteMultiplexCodesAction.All)
                {
                    // Build DELETE statements for all IHS multiplex rows
                    ihsMultiplexDelStatements.Add(
                        String.Format("DELETE FROM {0} WHERE {1} = {2}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_ihs_matrix.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_matrix.incidColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue("{0}")));

                    ihsMultiplexDelStatements.Add(
                        String.Format("DELETE FROM {0} WHERE {1} = {2}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_ihs_formation.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_formation.incidColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue("{0}")));

                    ihsMultiplexDelStatements.Add(
                        String.Format("DELETE FROM {0} WHERE {1} = {2}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_ihs_management.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_management.incidColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue("{0}")));

                    ihsMultiplexDelStatements.Add(
                        String.Format("DELETE FROM {0} WHERE {1} = {2}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_ihs_complex.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_complex.incidColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue("{0}")));
                }
                // If the IHS habitat code has changed and invalid multiplex codes
                // are to be deleted
                else if ((incidUpdateVals.Contains(_viewModelMain.HluDataset.incid.ihs_habitatColumn)) &&
                    (_bulkDeleteMultiplexCodes == (int)DeleteMultiplexCodesAction.Invalid))
                {
                    // Get the new IHS habitat code
                    string newIncidIhsHabitatCode = _viewModelMain.IncidCurrentRow[_viewModelMain.HluDataset.incid.ihs_habitatColumn.Ordinal].ToString();

                    // Build DELETE statements for IHS multiplex rows rendered obsolete by new IHS habitat
                    ihsMultiplexDelStatements.Add(
                        String.Format("DELETE FROM {0} WHERE {1} = {2} AND NOT EXISTS (SELECT {3} FROM {4} WHERE {3} = {0}.{5} AND {6} = {7})",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_ihs_matrix.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_matrix.incidColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue("{0}"),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.lut_ihs_habitat_ihs_matrix.code_matrixColumn.ColumnName),
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.lut_ihs_habitat_ihs_matrix.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_matrix.matrixColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.lut_ihs_habitat_ihs_matrix.code_habitatColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue(newIncidIhsHabitatCode)));
                    
                    ihsMultiplexDelStatements.Add(
                        String.Format("DELETE FROM {0} WHERE {1} = {2} AND NOT EXISTS (SELECT {3} FROM {4} WHERE {3} = {0}.{5} AND {6} = {7})",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_ihs_formation.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_formation.incidColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue("{0}"),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.lut_ihs_habitat_ihs_formation.code_formationColumn.ColumnName),
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.lut_ihs_habitat_ihs_formation.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_formation.formationColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.lut_ihs_habitat_ihs_formation.code_habitatColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue(newIncidIhsHabitatCode)));

                    ihsMultiplexDelStatements.Add(
                        String.Format("DELETE FROM {0} WHERE {1} = {2} AND NOT EXISTS (SELECT {3} FROM {4} WHERE {3} = {0}.{5} AND {6} = {7})",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_ihs_management.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_management.incidColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue("{0}"),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.lut_ihs_habitat_ihs_management.code_managementColumn.ColumnName),
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.lut_ihs_habitat_ihs_management.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_management.managementColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.lut_ihs_habitat_ihs_management.code_habitatColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue(newIncidIhsHabitatCode)));

                    ihsMultiplexDelStatements.Add(
                        String.Format("DELETE FROM {0} WHERE {1} = {2} AND NOT EXISTS (SELECT {3} FROM {4} WHERE {3} = {0}.{5} AND {6} = {7})",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_ihs_complex.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_complex.incidColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue("{0}"),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.lut_ihs_habitat_ihs_complex.code_complexColumn.ColumnName),
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.lut_ihs_habitat_ihs_complex.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_complex.complexColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.lut_ihs_habitat_ihs_complex.code_habitatColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue(newIncidIhsHabitatCode)));
                }

                // Filter out any rows not set (because the maximum number of blank rows are
                // created at the start of the bulk process so any not used need to be removed)
                _viewModelMain.IncidIhsMatrixRows = FilterUpdateRows<HluDataSet.incid_ihs_matrixDataTable,
                    HluDataSet.incid_ihs_matrixRow>(_viewModelMain.IncidIhsMatrixRows);

                _viewModelMain.IncidIhsFormationRows = FilterUpdateRows<HluDataSet.incid_ihs_formationDataTable,
                    HluDataSet.incid_ihs_formationRow>(_viewModelMain.IncidIhsFormationRows);

                _viewModelMain.IncidIhsManagementRows = FilterUpdateRows<HluDataSet.incid_ihs_managementDataTable,
                    HluDataSet.incid_ihs_managementRow>(_viewModelMain.IncidIhsManagementRows);

                _viewModelMain.IncidIhsComplexRows = FilterUpdateRows<HluDataSet.incid_ihs_complexDataTable,
                    HluDataSet.incid_ihs_complexRow>(_viewModelMain.IncidIhsComplexRows);

                _viewModelMain.IncidSourcesRows = FilterUpdateRows<HluDataSet.incid_sourcesDataTable,
                    HluDataSet.incid_sourcesRow>(_viewModelMain.IncidSourcesRows);

                //---------------------------------------------------------------------
                // FIX: 078 Bulk update overhaul/improvements.
                // 
                // Loop through each row in the incid selection
                foreach (DataRow r in _viewModelMain.IncidSelection.Rows)
                {
                    // Set the incid for the current row
                    string currIncid = r[incidOrdinal].ToString();

                    // Perform the bulk updates on the data tables
                    if (Settings.Default.BulkUpdateUsesAdo)
                        BulkUpdateAdo(currIncid, selectCommandIncid, updateCommandIncid, null,
                            ihsMultiplexDelStatements, _bulkDeleteMultiplexCodes, _bulkDeleteOrphanBapHabitats, _bulkDeletePotentialBapHabitats);
                    else
                        BulkUpdateDb(currIncid, selectCommandIncid, updateCommandIncid, null,
                            ihsMultiplexDelStatements, _bulkDeleteMultiplexCodes, _bulkDeleteOrphanBapHabitats, _bulkDeletePotentialBapHabitats);
                }
                //---------------------------------------------------------------------

                // Perform the bulk updates on the GIS data, shadow copy in DB and history
                BulkUpdateGis(incidOrdinal, _viewModelMain.IncidSelection, _bulkCreateHistory, ViewModelWindowMain.Operations.BulkUpdate, nowDtTm);

                // Commit the transaction and accept the changes ???
                _viewModelMain.DataBase.CommitTransaction();
                _viewModelMain.HluDataset.AcceptChanges();

                // force re-loading data from db
                _viewModelMain.IncidTable.Clear();
                
                MessageBox.Show("Bulk update succeeded.", "HLU: Bulk Update",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                _viewModelMain.DataBase.RollbackTransaction();
                MessageBox.Show(String.Format("Bulk update failed. The error message returned was:\n\n{0}",
                    ex.Message), "HLU: Bulk Update", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BulkUpdateResetControls();
            }
        }

        /// <summary>
        /// Cancels the bulk update mode.
        /// </summary>
        public void CancelBulkUpdate()
        {
            BulkUpdateResetControls();
        }

        /// <summary>
        /// Stops the bulk update mode and resets all
        /// the controls to normal.
        /// </summary>
        private void BulkUpdateResetControls()
        {
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            // Force the Incid table to be refilled because it has been
            // updated directly in the database rather than via the
            // local copy.
            _viewModelMain.RefillIncidTable = true;
            //---------------------------------------------------------------------

            // Reset the incid filter
            _viewModelMain.BulkUpdateMode = null;
            //_viewModelMain.IncidCurrentRowIndex = 1;

            // Clear the active filter.
            _viewModelMain.ClearFilter(true);

            // Select the first incid in the map.
            _viewModelMain.SelectOnMap(false);

            // Count the number of toids and fragments for the current incid
            // selected in the GIS and in the database.
            _viewModelMain.CountToidFrags();

            // Stop the bulk update mode
            _viewModelMain.BulkUpdateMode = false;

            // Enable the history tab
            _viewModelMain.TabItemHistoryEnabled = true;

            // Restore the filter and re-select all the features
            //_viewModelMain.IncidSelection = _incidSelectionBackup;  // Skip this!!!???
            //_viewModelMain.SelectAllOnMap();                        // Skip this!!!???

            // Refresh all the controls
            _viewModelMain.RefreshAll();

            // Reset the cursor
            _viewModelMain.ChangeCursor(Cursors.Arrow, String.Empty);
        }

        #endregion

        #region OSMM Bulk Update

        //---------------------------------------------------------------------
        // CHANGED: CR49 Process proposed OSMM Updates
        // Functionality for applying pending OSMM Bulk Updates.
        //    
        /// <summary>
        /// Applies the osmm bulk update.
        /// </summary>
        public void ApplyOSMMBulkUpdate()
        {
            _viewModelMain.ChangeCursor(Cursors.Wait, "OSMM Bulk updating ...");

            _viewModelMain.DataBase.BeginTransaction(true, IsolationLevel.ReadCommitted);

            try
            {
                // Build a blank WHERE clause based on incid column
                string incidWhereClause = String.Format(" WHERE {0} = {1}",
                    _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid.incidColumn.ColumnName),
                    _viewModelMain.DataBase.QuoteValue("{0}"));

                // Build the SELECT statement based on the incid where clause
                string selectCommandIncid = String.Format("SELECT {0} FROM {1}{2}",
                    _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid.ihs_habitatColumn.ColumnName),
                    _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid.TableName), incidWhereClause);

                // Update the last modified date & user fields and the ihs version on
                // the incid table
                DateTime currDtTm = DateTime.Now;
                DateTime nowDtTm = new DateTime(currDtTm.Year, currDtTm.Month, currDtTm.Day, currDtTm.Hour, currDtTm.Minute, currDtTm.Second, DateTimeKind.Local);
                //---------------------------------------------------------------------
                _viewModelMain.IncidCurrentRow.last_modified_date = nowDtTm;
                _viewModelMain.IncidCurrentRow.last_modified_user_id = _viewModelMain.UserID;
                _viewModelMain.IncidCurrentRow.ihs_version = _viewModelMain.IhsVersion;

                // Set the ihs habitat on the incid table to a non-null value to
                // indicate that the column has/will change
                _viewModelMain.IncidCurrentRow.ihs_habitat = "";

                // Get the column ordinal for the incid column on the incid table
                int incidOrdinal =
                    _viewModelMain.IncidSelection.Columns[_viewModelMain.HluDataset.incid.incidColumn.ColumnName].Ordinal;

                // Build a collection of the updated columns in the incid table
                var incidUpdateVals = _viewModelMain.HluDataset.incid.Columns.Cast<DataColumn>()
                    .Where(c => c.Ordinal != _viewModelMain.HluDataset.incid.incidColumn.Ordinal &&
                        !_viewModelMain.IncidCurrentRow.IsNull(c.Ordinal));
                
                // Loop through each row in the incid selection
                foreach (DataRow r in _viewModelMain.IncidSelection.Rows)
                {
                    // Set the incid for the current row
                    string currIncid = r[incidOrdinal].ToString();
                    string[] relValues = new string[] { currIncid };

                    // Load the incid_osmm_updates table for the current incid
                    HluDataSet.incid_osmm_updatesDataTable incidOSMMUpdatesTable = (HluDataSet.incid_osmm_updatesDataTable)_viewModelMain.HluDataset.incid_osmm_updates.Copy();
                    HluDataSet.incid_osmm_updatesRow[] incidOSMMUpdatesRows = _viewModelMain.GetIncidChildRowsDb(relValues,
                        _viewModelMain.HluTableAdapterManager.incid_osmm_updatesTableAdapter, ref incidOSMMUpdatesTable);

                    // Get the osmm_xref_id
                    int incidOSMMXrefId = incidOSMMUpdatesRows[0].osmm_xref_id;

                    // Get the lut_osmm_ihs_xref row for the current osmm_xref_id
                    IEnumerable<HluDataSet.lut_osmm_ihs_xrefRow> ihsOSMMXref = from x in _viewModelMain.HluDataset.lut_osmm_ihs_xref
                                                                               where x.is_local && x.osmm_xref_id == incidOSMMXrefId
                                                                               select x;

                    ////////////////////////////////////////
                    // Exit if no row is found
                    if (ihsOSMMXref == null)
                    {
                    }
                    ////////////////////////////////////////

                    // Get the new IHS habitat code from the lut_osmm_ihs_xref table
                    string newIncidIhsHabitatCode = ihsOSMMXref.ElementAt(0).ihs_habitat;

                    // Set the ihs habitat on the incid table
                    _viewModelMain.IncidCurrentRow.ihs_habitat = newIncidIhsHabitatCode;

                    // Build an UPDATE statement for the incid table
                    string updateCommandIncid = incidUpdateVals.Count() == 0 ? String.Empty :
                        new StringBuilder(String.Format("UPDATE {0} SET ",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid.TableName)))
                        .Append(_viewModelMain.HluDataset.incid.Columns.Cast<DataColumn>()
                        .Where(c => c.Ordinal != _viewModelMain.HluDataset.incid.incidColumn.Ordinal &&
                            !_viewModelMain.IncidCurrentRow.IsNull(c.Ordinal))
                        .Aggregate(new StringBuilder(), (sb, c) => sb.Append(String.Format(", {0} = {1}",
                            _viewModelMain.DataBase.QuoteIdentifier(c.ColumnName),
                            _viewModelMain.DataBase.QuoteValue(_viewModelMain.IncidCurrentRow[c.Ordinal]))))
                            .Remove(0, 2)).Append(incidWhereClause).ToString();

                    // Build an UPDATE statement for the incid_osmm_updates table
                    string updateCommandIncidOSMMUpdates = new StringBuilder(String.Format("UPDATE {0} SET {1} = -1, {2} = {3}, {4} = {5}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_osmm_updates.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_osmm_updates.statusColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_osmm_updates.last_modified_dateColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue(nowDtTm),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_osmm_updates.last_modified_user_idColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue(_viewModelMain.UserID)))
                        .Append(incidWhereClause).ToString();

                    // Create new rows in the matrix table from the lut_osmm_ihs_xref table
                    string newMultiplexCode;
                    int multiplexRow = 0;
                    _viewModelMain.IncidIhsMatrixRows = new HluDataSet.incid_ihs_matrixRow[3]
                        .Select(mr => _viewModelMain.HluDataset.incid_ihs_matrix.Newincid_ihs_matrixRow()).ToArray();
                    if (!ihsOSMMXref.ElementAt(0).Isihs_matrix1Null())
                    {
                        newMultiplexCode = ihsOSMMXref.ElementAt(0).ihs_matrix1;
                        _viewModelMain.IncidIhsMatrixRows[multiplexRow].matrix_id = multiplexRow;
                        _viewModelMain.IncidIhsMatrixRows[multiplexRow].incid = currIncid;
                        _viewModelMain.IncidIhsMatrixRows[multiplexRow].matrix = (string)newMultiplexCode;
                        multiplexRow += 1;
                    }
                    if (!ihsOSMMXref.ElementAt(0).Isihs_matrix2Null())
                    {
                        newMultiplexCode = ihsOSMMXref.ElementAt(0).ihs_matrix2;
                        _viewModelMain.IncidIhsMatrixRows[multiplexRow].matrix_id = multiplexRow;
                        _viewModelMain.IncidIhsMatrixRows[multiplexRow].incid = currIncid;
                        _viewModelMain.IncidIhsMatrixRows[multiplexRow].matrix = (string)newMultiplexCode;
                        multiplexRow += 1;
                    }
                    if (!ihsOSMMXref.ElementAt(0).Isihs_matrix3Null())
                    {
                        newMultiplexCode = ihsOSMMXref.ElementAt(0).ihs_matrix3;
                        _viewModelMain.IncidIhsMatrixRows[multiplexRow].matrix_id = multiplexRow;
                        _viewModelMain.IncidIhsMatrixRows[multiplexRow].incid = currIncid;
                        _viewModelMain.IncidIhsMatrixRows[multiplexRow].matrix = (string)newMultiplexCode;
                    }

                    // Create new rows in the formation table from the lut_osmm_ihs_xref table
                    multiplexRow = 0;
                    _viewModelMain.IncidIhsFormationRows = new HluDataSet.incid_ihs_formationRow[2]
                        .Select(mr => _viewModelMain.HluDataset.incid_ihs_formation.Newincid_ihs_formationRow()).ToArray();
                    if (!ihsOSMMXref.ElementAt(0).Isihs_formation1Null())
                    {
                        newMultiplexCode = ihsOSMMXref.ElementAt(0).ihs_formation1;
                        _viewModelMain.IncidIhsFormationRows[multiplexRow].formation_id = multiplexRow;
                        _viewModelMain.IncidIhsFormationRows[multiplexRow].incid = currIncid;
                        _viewModelMain.IncidIhsFormationRows[multiplexRow].formation = (string)newMultiplexCode;
                        multiplexRow += 1;
                    }
                    if (!ihsOSMMXref.ElementAt(0).Isihs_formation2Null())
                    {
                        newMultiplexCode = ihsOSMMXref.ElementAt(0).ihs_formation2;
                        _viewModelMain.IncidIhsFormationRows[multiplexRow].formation_id = multiplexRow;
                        _viewModelMain.IncidIhsFormationRows[multiplexRow].incid = currIncid;
                        _viewModelMain.IncidIhsFormationRows[multiplexRow].formation = (string)newMultiplexCode;
                    }

                    // Create new rows in the management table from the lut_osmm_ihs_xref table
                    multiplexRow = 0;
                    _viewModelMain.IncidIhsManagementRows = new HluDataSet.incid_ihs_managementRow[2]
                        .Select(mr => _viewModelMain.HluDataset.incid_ihs_management.Newincid_ihs_managementRow()).ToArray();
                    if (!ihsOSMMXref.ElementAt(0).Isihs_management1Null())
                    {
                        newMultiplexCode = ihsOSMMXref.ElementAt(0).ihs_management1;
                        _viewModelMain.IncidIhsManagementRows[multiplexRow].management_id = multiplexRow;
                        _viewModelMain.IncidIhsManagementRows[multiplexRow].incid = currIncid;
                        _viewModelMain.IncidIhsManagementRows[multiplexRow].management = (string)newMultiplexCode;
                        multiplexRow += 1;
                    }
                    if (!ihsOSMMXref.ElementAt(0).Isihs_management2Null())
                    {
                        newMultiplexCode = ihsOSMMXref.ElementAt(0).ihs_management2;
                        _viewModelMain.IncidIhsManagementRows[multiplexRow].management_id = multiplexRow;
                        _viewModelMain.IncidIhsManagementRows[multiplexRow].incid = currIncid;
                        _viewModelMain.IncidIhsManagementRows[multiplexRow].management = (string)newMultiplexCode;
                    }

                    // Create new rows in the complex table from the lut_osmm_ihs_xref table
                    multiplexRow = 0;
                    _viewModelMain.IncidIhsComplexRows = new HluDataSet.incid_ihs_complexRow[2]
                        .Select(mr => _viewModelMain.HluDataset.incid_ihs_complex.Newincid_ihs_complexRow()).ToArray();
                    if (!ihsOSMMXref.ElementAt(0).Isihs_complex1Null())
                    {
                        newMultiplexCode = ihsOSMMXref.ElementAt(0).ihs_complex1;
                        _viewModelMain.IncidIhsComplexRows[multiplexRow].complex_id = multiplexRow;
                        _viewModelMain.IncidIhsComplexRows[multiplexRow].incid = currIncid;
                        _viewModelMain.IncidIhsComplexRows[multiplexRow].complex = (string)newMultiplexCode;
                        multiplexRow += 1;
                    }
                    if (!ihsOSMMXref.ElementAt(0).Isihs_complex2Null())
                    {
                        newMultiplexCode = ihsOSMMXref.ElementAt(0).ihs_complex2;
                        _viewModelMain.IncidIhsComplexRows[multiplexRow].complex_id = multiplexRow;
                        _viewModelMain.IncidIhsComplexRows[multiplexRow].incid = currIncid;
                        _viewModelMain.IncidIhsComplexRows[multiplexRow].complex = (string)newMultiplexCode;
                    }

                    // Filter out any rows not set (because the maximum number of blank rows are
                    // created above so any not used need to be removed)
                    _viewModelMain.IncidIhsMatrixRows = FilterUpdateRows<HluDataSet.incid_ihs_matrixDataTable,
                        HluDataSet.incid_ihs_matrixRow>(_viewModelMain.IncidIhsMatrixRows);

                    _viewModelMain.IncidIhsFormationRows = FilterUpdateRows<HluDataSet.incid_ihs_formationDataTable,
                        HluDataSet.incid_ihs_formationRow>(_viewModelMain.IncidIhsFormationRows);

                    _viewModelMain.IncidIhsManagementRows = FilterUpdateRows<HluDataSet.incid_ihs_managementDataTable,
                        HluDataSet.incid_ihs_managementRow>(_viewModelMain.IncidIhsManagementRows);

                    _viewModelMain.IncidIhsComplexRows = FilterUpdateRows<HluDataSet.incid_ihs_complexDataTable,
                        HluDataSet.incid_ihs_complexRow>(_viewModelMain.IncidIhsComplexRows);

                    _viewModelMain.IncidSourcesRows = FilterUpdateRows<HluDataSet.incid_sourcesDataTable,
                        HluDataSet.incid_sourcesRow>(_viewModelMain.IncidSourcesRows);


                    // Always assume all existing multiplex codes are to be deleted
                    List<string> ihsMultiplexDelStatements = new List<string>();
                    // Build DELETE statements for all IHS multiplex rows
                    ihsMultiplexDelStatements.Add(
                        String.Format("DELETE FROM {0} WHERE {1} = {2}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_ihs_matrix.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_matrix.incidColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue("{0}")));

                    ihsMultiplexDelStatements.Add(
                        String.Format("DELETE FROM {0} WHERE {1} = {2}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_ihs_formation.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_formation.incidColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue("{0}")));

                    ihsMultiplexDelStatements.Add(
                        String.Format("DELETE FROM {0} WHERE {1} = {2}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_ihs_management.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_management.incidColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue("{0}")));

                    ihsMultiplexDelStatements.Add(
                        String.Format("DELETE FROM {0} WHERE {1} = {2}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_ihs_complex.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_ihs_complex.incidColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue("{0}")));

                    // Perform the bulk updates on the data tables
                    if (Settings.Default.BulkUpdateUsesAdo)
                        BulkUpdateAdo(currIncid, selectCommandIncid, updateCommandIncid, updateCommandIncidOSMMUpdates,
                            ihsMultiplexDelStatements, _bulkDeleteMultiplexCodes, _bulkDeleteOrphanBapHabitats, _bulkDeletePotentialBapHabitats);
                    else
                        BulkUpdateDb(currIncid, selectCommandIncid, updateCommandIncid, updateCommandIncidOSMMUpdates,
                            ihsMultiplexDelStatements, _bulkDeleteMultiplexCodes, _bulkDeleteOrphanBapHabitats, _bulkDeletePotentialBapHabitats);
                }

                // Perform the bulk updates on the GIS data, shadow copy in DB and history
                BulkUpdateGis(incidOrdinal, _viewModelMain.IncidSelection, _bulkCreateHistory, ViewModelWindowMain.Operations.OSMMUpdate, nowDtTm);

                // Commit the transaction and accept the changes ???
                _viewModelMain.DataBase.CommitTransaction();
                _viewModelMain.HluDataset.AcceptChanges();

                // force re-loading data from db
                _viewModelMain.IncidTable.Clear();

                MessageBox.Show("OSMM Bulk update succeeded.", "HLU: OSMM Bulk Update",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                _viewModelMain.DataBase.RollbackTransaction();
                MessageBox.Show(String.Format("OSMM Bulk update failed. The error message returned was:\n\n{0}",
                    ex.Message), "HLU: OSMM Bulk Update", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                OSMMBulkUpdateResetControls();
            }
        }

        /// <summary>
        /// Cancels the osmm bulk update mode.
        /// </summary>
        public void CancelOSMMBulkUpdate()
        {
            OSMMBulkUpdateResetControls();
        }

        /// <summary>
        /// Stops the osmm bulk update mode and resets all
        /// the controls to normal.
        /// </summary>
        private void OSMMBulkUpdateResetControls()
        {
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            // Force the Incid table to be refilled because it has been
            // updated directly in the database rather than via the
            // local copy.
            _viewModelMain.RefillIncidTable = true;
            //---------------------------------------------------------------------

            // Reset the incid and map selections and move
            // to the first incid in the database.
            _viewModelMain.BulkUpdateMode = null;
            _viewModelMain.OSMMBulkUpdateMode = false;
            //_viewModelMain.IncidCurrentRowIndex = 1;      // Skip this!!!???
            
            // Clear the active filter.
            _viewModelMain.ClearFilter(true);

            // Select the first incid in the map.
            _viewModelMain.SelectOnMap(false);

            // Count the number of toids and fragments for the current incid
            // selected in the GIS and in the database.
            _viewModelMain.CountToidFrags();

            // Stop the bulk update mode
            _viewModelMain.BulkUpdateMode = false;

            // Enable the ihs and history tabs
            _viewModelMain.TabItemIhsEnabled = true;
            _viewModelMain.TabItemHistoryEnabled = true;

            // Select the ihs tab
            _viewModelMain.TabItemSelected = 0;

            // Refresh all the controls
            _viewModelMain.RefreshAll();

            // Reset the cursor
            _viewModelMain.ChangeCursor(Cursors.Arrow, String.Empty);
        }
        //---------------------------------------------------------------------

        #endregion

        #region Database & GIS Updates

        /// <summary>
        /// Perform the OSMM bulk update on the database tables using an ADO connection.
        /// </summary>
        /// <param name="currIncid">The incid to update.</param>
        /// <param name="selectCommandIncid">The SELECT command template.</param>
        /// <param name="updateCommandIncid">The UPDATE command template.</param>
        /// <param name="deleteExistingRows">if set to <c>1</c> [delete existing rows].</param>
        /// <param name="ihsMultiplexDelStatements">The ihs multiplex delete statements.</param>
        /// <exception cref="Exception">
        /// Failed to delete IHS multiplex rows.
        /// or
        /// Failed to update incid table.
        /// or
        /// No database row for incid
        /// </exception>
        private void BulkUpdateAdo(string currIncid,
            string selectCommandIncid,
            string updateCommandIncid,
            string updateCommandIncidOSMMUpdates,
            List<string> ihsMultiplexDelStatements,
            int bulkDeleteMultiplexCodes,
            bool bulkDeleteOrphanBapHabitats,
            bool bulkDeletePotentialBapHabitats)
        {
            // Delete any orphaned IHS multiplex rows
            if (ihsMultiplexDelStatements != null)
            {
                foreach (string s in ihsMultiplexDelStatements)
                    if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format(s, currIncid),
                        _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                        throw new Exception("Failed to delete IHS multiplex rows.");
            }

            // Execute the UPDATE incid statement for the current row
            if (!String.IsNullOrEmpty(updateCommandIncid))
            {
                if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format(updateCommandIncid, currIncid),
                    _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                    throw new Exception("Failed to update incid table.");
            }

            // Execute the UPDATE incid_osmm_updates statement for the current row
            if (!String.IsNullOrEmpty(updateCommandIncidOSMMUpdates))
            {
                if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format(updateCommandIncidOSMMUpdates, currIncid),
                    _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                    throw new Exception("Failed to update incid_osmm_updates table.");
            }

            // Execute the SELECT statement to get the IHS habitat code
            object retValue = _viewModelMain.DataBase.ExecuteScalar(String.Format(selectCommandIncid, currIncid),
                _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text);

            if (retValue == null) throw new Exception(
                String.Format("No database row for incid '{0}'", currIncid));

            // Set the IHS habitat code
            string ihsHabitat = retValue.ToString();

            // Create an array of the value of the current incid
            object[] relValues = new object[] { currIncid };

            // Store the rows from the user interface
            HluDataSet.incid_ihs_matrixRow[] incidIhsMatrixRows = _viewModelMain.IncidIhsMatrixRows;
            // Store a copy of the table for the current incid
            HluDataSet.incid_ihs_matrixDataTable ihsMatrixTable =
                (HluDataSet.incid_ihs_matrixDataTable)_viewModelMain.HluDataset.incid_ihs_matrix.Copy();
            // Update the rows in the database
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            BulkUpdateAdoMultiplexSourceTable(bulkDeleteMultiplexCodes, currIncid, relValues,
                _viewModelMain.HluTableAdapterManager.incid_ihs_matrixTableAdapter,
                ihsMatrixTable, ref incidIhsMatrixRows);
            //---------------------------------------------------------------------
            //_viewModelMain.IncidIhsMatrixRows = incidIhsMatrixRows;

            // Store the rows from the user interface
            HluDataSet.incid_ihs_formationRow[] incidIhsFormationRows = _viewModelMain.IncidIhsFormationRows;
            // Store a copy of the table for the current incid
            HluDataSet.incid_ihs_formationDataTable ihsFormationTable =
                (HluDataSet.incid_ihs_formationDataTable)_viewModelMain.HluDataset.incid_ihs_formation.Copy();
            // Update the rows in the database
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            BulkUpdateAdoMultiplexSourceTable(bulkDeleteMultiplexCodes, currIncid, relValues,
                _viewModelMain.HluTableAdapterManager.incid_ihs_formationTableAdapter,
                ihsFormationTable, ref incidIhsFormationRows);
            //---------------------------------------------------------------------
            //_viewModelMain.IncidIhsFormationRows = incidIhsFormationRows;

            // Store the rows from the user interface
            HluDataSet.incid_ihs_managementRow[] incidIhsManagementRows = _viewModelMain.IncidIhsManagementRows;
            // Store a copy of the table for the current incid
            HluDataSet.incid_ihs_managementDataTable ihsManagementTable =
                (HluDataSet.incid_ihs_managementDataTable)_viewModelMain.HluDataset.incid_ihs_management.Copy();
            // Update the rows in the database
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            BulkUpdateAdoMultiplexSourceTable(bulkDeleteMultiplexCodes, currIncid, relValues,
                _viewModelMain.HluTableAdapterManager.incid_ihs_managementTableAdapter,
                ihsManagementTable, ref incidIhsManagementRows);
            //---------------------------------------------------------------------
            //_viewModelMain.IncidIhsManagementRows = incidIhsManagementRows;

            // Store the rows from the user interface
            HluDataSet.incid_ihs_complexRow[] incidIhsComplexRows = _viewModelMain.IncidIhsComplexRows;
            // Store a copy of the table for the current incid
            HluDataSet.incid_ihs_complexDataTable ihsComplexTable =
                (HluDataSet.incid_ihs_complexDataTable)_viewModelMain.HluDataset.incid_ihs_complex.Copy();
            // Update the rows in the database
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            BulkUpdateAdoMultiplexSourceTable(bulkDeleteMultiplexCodes, currIncid, relValues,
                _viewModelMain.HluTableAdapterManager.incid_ihs_complexTableAdapter,
                ihsComplexTable, ref incidIhsComplexRows);
            //---------------------------------------------------------------------
            //_viewModelMain.IncidIhsComplexRows = incidIhsComplexRows;

            // Store a copy of the table for the current incid
            HluDataSet.incid_bapDataTable bapTable = (HluDataSet.incid_bapDataTable)_viewModelMain.HluDataset.incid_bap.Copy();
            // Load the child rows for the bap table for the supplied incid
            _viewModelMain.GetIncidChildRowsDb(relValues,
                _viewModelMain.HluTableAdapterManager.incid_bapTableAdapter, ref bapTable);
            // Update the rows in the database
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            BulkUpdateBap(currIncid, ihsHabitat, bapTable, incidIhsMatrixRows, incidIhsFormationRows,
                incidIhsManagementRows, incidIhsComplexRows, bulkDeleteOrphanBapHabitats, bulkDeletePotentialBapHabitats);
            //---------------------------------------------------------------------

            // Store the rows from the user interface
            HluDataSet.incid_sourcesRow[] incidSourcesRows = _viewModelMain.IncidSourcesRows;
            // Store a copy of the table for the current incid
            HluDataSet.incid_sourcesDataTable sourcesTable =
                (HluDataSet.incid_sourcesDataTable)_viewModelMain.HluDataset.incid_sources.Copy();
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            // Count the non-blank rows from the user interface
            int newRows = (from nr in incidSourcesRows
                           where nr.source_id != Int32.MinValue
                           select nr).Count();
            // If there are new source rows then delete the old sources
            int deleteSources = newRows > 0 ? 1 : 3;
            // Update the rows in the database
            BulkUpdateAdoMultiplexSourceTable(deleteSources, currIncid, relValues,
                _viewModelMain.HluTableAdapterManager.incid_sourcesTableAdapter,
                sourcesTable, ref incidSourcesRows);
            //_viewModelMain.IncidSourcesRows = incidSourcesRows;
            //---------------------------------------------------------------------

        }

        /// <summary>
        /// Perform the bulk update on the database tables using a DB connection.
        /// </summary>
        /// <param name="currIncid">The incid to update.</param>
        /// <param name="selectCommandIncid">The SELECT command template.</param>
        /// <param name="updateCommandIncid">The UPDATE command template.</param>
        /// <param name="deleteExistingRows">if set to <c>1</c> [delete existing rows].</param>
        /// <param name="ihsMultiplexDelStatements">The ihs multiplex delete statements.</param>
        /// <exception cref="Exception">
        /// Failed to delete IHS multiplex rows.
        /// or
        /// Failed to update incid table.
        /// or
        /// No database row for incid
        /// </exception>
        private void BulkUpdateDb(string currIncid,
            string selectCommandIncid,
            string updateCommandIncid,
            string updateCommandIncidOSMMUpdates,
            List<string> ihsMultiplexDelStatements,
            int bulkDeleteMultiplexCodes,
            bool bulkDeleteOrphanBapHabitats,
            bool bulkDeletePotentialBapHabitats)
        {
            // Delete any orphaned IHS multiplex rows
            if (ihsMultiplexDelStatements != null)
            {
                foreach (string s in ihsMultiplexDelStatements)
                    if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format(s, currIncid),
                        _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                        throw new Exception("Failed to delete IHS multiplex rows.");
            }

            // Execute the UPDATE incid statement for the current row
            if (!String.IsNullOrEmpty(updateCommandIncid))
            {
                if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format(updateCommandIncid, currIncid),
                    _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                    throw new Exception("Failed to update incid table.");
            }

            // Execute the UPDATE incid_osmm_updates statement for the current row
            if (!String.IsNullOrEmpty(updateCommandIncidOSMMUpdates))
            {
                if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format(updateCommandIncidOSMMUpdates, currIncid),
                    _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                    throw new Exception("Failed to update incid_osmm_updates table.");
            }

            // Execute the SELECT statement to get the IHS habitat code
            object retValue = _viewModelMain.DataBase.ExecuteScalar(String.Format(selectCommandIncid, currIncid),
                _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text);

            if (retValue == null) throw new Exception(
                String.Format("No database row for incid '{0}'", currIncid));

            // Set the IHS habitat code
            string ihsHabitat = retValue.ToString();

            // Create an array of the value of the current incid
            object[] relValues = new object[] { currIncid };

            // Store the rows from the user interface
            HluDataSet.incid_ihs_matrixRow[] incidIhsMatrixRows = _viewModelMain.IncidIhsMatrixRows;
            // Store a copy of the table for the current incid
            HluDataSet.incid_ihs_matrixDataTable ihsMatrixTable =
                (HluDataSet.incid_ihs_matrixDataTable)_viewModelMain.HluDataset.incid_ihs_matrix.Copy();
            // Update the rows in the database
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            BulkUpdateDbMultiplexSourceTable(bulkDeleteMultiplexCodes, currIncid, relValues,
                _viewModelMain.HluTableAdapterManager.incid_ihs_matrixTableAdapter,
                ihsMatrixTable, ref incidIhsMatrixRows);
            //---------------------------------------------------------------------

            // Store the rows from the user interface
            HluDataSet.incid_ihs_formationRow[] incidIhsFormationRows = _viewModelMain.IncidIhsFormationRows;
            // Store a copy of the table for the current incid
            HluDataSet.incid_ihs_formationDataTable ihsFormationTable =
                (HluDataSet.incid_ihs_formationDataTable)_viewModelMain.HluDataset.incid_ihs_formation.Copy();
            // Update the rows in the database
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            BulkUpdateDbMultiplexSourceTable(bulkDeleteMultiplexCodes, currIncid, relValues,
                _viewModelMain.HluTableAdapterManager.incid_ihs_formationTableAdapter,
                ihsFormationTable, ref incidIhsFormationRows);
            //---------------------------------------------------------------------

            // Store the rows from the user interface
            HluDataSet.incid_ihs_managementRow[] incidIhsManagementRows = _viewModelMain.IncidIhsManagementRows;
            // Store a copy of the table for the current incid
            HluDataSet.incid_ihs_managementDataTable ihsManagementTable =
                (HluDataSet.incid_ihs_managementDataTable)_viewModelMain.HluDataset.incid_ihs_management.Copy();
            // Update the rows in the database
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            BulkUpdateDbMultiplexSourceTable(bulkDeleteMultiplexCodes, currIncid, relValues,
                _viewModelMain.HluTableAdapterManager.incid_ihs_managementTableAdapter,
                ihsManagementTable, ref incidIhsManagementRows);
            //---------------------------------------------------------------------
            //_viewModelMain.IncidIhsManagementRows = incidIhsManagementRows;

            // Store the rows from the user interface
            HluDataSet.incid_ihs_complexRow[] incidIhsComplexRows = _viewModelMain.IncidIhsComplexRows;
            // Store a copy of the table for the current incid
            HluDataSet.incid_ihs_complexDataTable ihsComplexTable =
                (HluDataSet.incid_ihs_complexDataTable)_viewModelMain.HluDataset.incid_ihs_complex.Copy();
            // Update the rows in the database
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            BulkUpdateDbMultiplexSourceTable(bulkDeleteMultiplexCodes, currIncid, relValues,
                _viewModelMain.HluTableAdapterManager.incid_ihs_complexTableAdapter,
                ihsComplexTable, ref incidIhsComplexRows);
            //---------------------------------------------------------------------

            // Store a copy of the table for the current incid
            HluDataSet.incid_bapDataTable bapTable = (HluDataSet.incid_bapDataTable)_viewModelMain.HluDataset.incid_bap.Copy();
            // Load the child rows for the bap table for the supplied incid
            _viewModelMain.GetIncidChildRowsDb(relValues,
                _viewModelMain.HluTableAdapterManager.incid_bapTableAdapter, ref bapTable);
            // Update the rows in the database
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            BulkUpdateBap(currIncid, ihsHabitat, bapTable, incidIhsMatrixRows, incidIhsFormationRows,
                incidIhsManagementRows, incidIhsComplexRows, bulkDeleteOrphanBapHabitats, bulkDeletePotentialBapHabitats);
            //---------------------------------------------------------------------

            // Store the rows from the user interface
            HluDataSet.incid_sourcesRow[] incidSourcesRows = _viewModelMain.IncidSourcesRows;
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            // Store a copy of the table for the current incid
            HluDataSet.incid_sourcesDataTable incidSourcesTable =
                (HluDataSet.incid_sourcesDataTable)_viewModelMain.HluDataset.incid_sources.Copy();
            // Count the non-blank rows from the user interface
            int newRows = (from nr in incidSourcesRows
                           where nr.source_id != Int32.MinValue
                           select nr).Count();
            // If there are new source rows then delete the old sources
            int deleteSources = newRows > 0 ? 1 : 3;
            // Update the rows in the database
            BulkUpdateDbMultiplexSourceTable(deleteSources, currIncid, relValues,
                _viewModelMain.HluTableAdapterManager.incid_sourcesTableAdapter,
                incidSourcesTable, ref incidSourcesRows);
            //---------------------------------------------------------------------
        }

        /// <summary>
        /// Bulk updates the multiplex and source tables using an ADO connection.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="R"></typeparam>
        /// <param name="deleteExistingRows">if set to <c>1</c> [delete existing rows].</param>
        /// <param name="currIncid">The current incid.</param>
        /// <param name="relValues">An array of the current incid.</param>
        /// <param name="adapter">The table adapter.</param>
        /// <param name="dbTable">The database table to update.</param>
        /// <param name="uiRows">The user interface rows.</param>
        private void BulkUpdateAdoMultiplexSourceTable<T, R>(int deleteExistingRows, string currIncid,
            object[] relValues, HluTableAdapter<T, R> adapter, T dbTable, ref R[] uiRows)
            where T : DataTable, new()
            where R : DataRow
        {
            if ((uiRows != null) && (uiRows.Count(r =>
                !r.IsNull(dbTable.PrimaryKey[0].Ordinal)) > 0))
            {
                // Create a cloned set of rows for the supplied incid
                T newRows = CloneUpdateRows<T, R>(uiRows, currIncid);
                
                // Load the child rows for the required data table for the supplied incid
                _viewModelMain.GetIncidChildRowsDb(relValues, adapter, ref dbTable);

                // Update the child rows for the supplied incid to match the cloned rows
                BulkUpdateAdoChildTable<T, R>(deleteExistingRows, newRows.AsEnumerable()
                    .Cast<R>().ToArray(), ref dbTable, adapter);

                uiRows = newRows.AsEnumerable().Cast<R>().ToArray();
            }
        }

        /// <summary>
        /// Bulk updates the multiplex and source tables using a DB connection.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="R"></typeparam>
        /// <param name="deleteExistingRows">if set to <c>1</c> [delete existing rows].</param>
        /// <param name="currIncid">The current incid.</param>
        /// <param name="relValues">An array of the current incid.</param>
        /// <param name="adapter">The table adapter.</param>
        /// <param name="dbTable">The database table to update.</param>
        /// <param name="uiRows">The user interface rows.</param>
        private void BulkUpdateDbMultiplexSourceTable<T, R>(int deleteExistingRows, string currIncid,
            object[] relValues, HluTableAdapter<T, R> adapter, T dbTable, ref R[] uiRows)
            where T : DataTable, new()
            where R : DataRow
        {
            if ((uiRows != null) && (uiRows.Count(rm =>
                !rm.IsNull(dbTable.PrimaryKey[0].Ordinal)) > 0))
            {
                // Create a cloned set of rows for the supplied incid
                T newRows = CloneUpdateRows<T, R>(uiRows, currIncid);

                // Load the child rows for the required data table for the supplied incid
                _viewModelMain.GetIncidChildRowsDb(relValues, adapter, ref dbTable);

                // Update the child rows for the supplied incid to match the cloned rows
                BulkUpdateDbChildTable<T, R>(deleteExistingRows, newRows.AsEnumerable()
                    .Cast<R>().ToArray(), ref dbTable);

                uiRows = newRows.AsEnumerable().Cast<R>().ToArray();
            }
        }

        /// <summary>
        /// Filters out any rows not set (because the maximum number of blank
        /// rows are created at the start of the bulk process so any not used
        /// need to be removed)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="R"></typeparam>
        /// <param name="rows">The rows.</param>
        /// <returns></returns>
        private R[] FilterUpdateRows<T, R>(R[] rows)
            where T : DataTable
            where R : DataRow
        {
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            List<R> newRows = new List<R>(rows.Length);

            if ((rows == null) || (rows.Length == 0) || (rows[0] == null)) return newRows.ToArray();
            //---------------------------------------------------------------------

            T table = (T)rows[0].Table;

            var q = from rel in table.ParentRelations.Cast<DataRelation>()
                    where rel.ParentTable.TableName.ToLower().StartsWith("lut_") &&
                          rel.ParentColumns.Length == 1
                    select rel;

            var lookup = (from c in table.Columns.Cast<DataColumn>()
                          let p = from rel in q
                                  where rel.ChildColumns.Length == 1 && rel.ChildColumns[0].Ordinal == c.Ordinal
                                  select rel.ParentColumns[0]
                          select new
                          {
                              ChildColumnOrdinal = c.Ordinal,
                              ParentColumnOrdinal = p.Count() != 0 ? p.ElementAt(0).Ordinal : -1,
                              ParentTable = p.Count() != 0 ? p.ElementAt(0).Table.TableName : String.Empty
                          }).ToArray();

            for (int i = 0; i < rows.Length; i++)
            {
                if (rows[i] == null) continue;
                bool add = true;
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    if ((lookup[j].ParentColumnOrdinal != -1) && (_viewModelMain.HluDataset.Tables[lookup[j].ParentTable]
                        .AsEnumerable().Count(r => r[lookup[j].ParentColumnOrdinal].Equals(rows[i][j])) == 0))
                    {
                        add = false;
                        break;
                    }
                }
                if (add) newRows.Add(rows[i]);
            }

            return newRows.ToArray();
        }

        /// <summary>
        /// Clones the rows in a data table, changing the incid column in the
        /// cloned rows to match the supplied incid value.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="R"></typeparam>
        /// <param name="rows">The source rows to be cloned.</param>
        /// <param name="incid">The incid value to be set in the cloned rows.</param>
        /// <returns></returns>
        private T CloneUpdateRows<T, R>(R[] rows, string incid)
            where T : DataTable, new()
            where R : DataRow
        {
            T newRows = new T();
            int incidOrdinal = newRows.Columns[_viewModelMain.HluDataset.incid.incidColumn.ColumnName].Ordinal;

            for (int i = 0; i < rows.Length; i++)
            {
                object[] itemArray = new object[rows[i].ItemArray.Length];
                Array.Copy(rows[i].ItemArray, itemArray, itemArray.Length);
                itemArray[incidOrdinal] = incid;
                R rm = (R)newRows.NewRow();
                rm.ItemArray = itemArray;
                newRows.Rows.Add(rm);
            }

            return newRows;
        }

        /// <summary>
        /// Uses the bap_habitat code to identify existing records and distinguish them from new ones to be inserted.
        /// </summary>
        /// <param name="currIncid">The current incid.</param>
        /// <param name="ihsHabitat">The ihs habitat.</param>
        /// <param name="incidBapTable">The current incid bap table in the database.</param>
        /// <param name="ihsMatrixRows">The ihs matrix rows from the user interface.</param>
        /// <param name="ihsFormationRows">The ihs formation rows from the user interface.</param>
        /// <param name="ihsManagementRows">The ihs management rows from the user interface.</param>
        /// <param name="ihsComplexRows">The ihs complex rows from the user interface.</param>
        /// <param name="deleteOrphanBapRows">if set to <c>true</c> delete any orphan primary bap rows.</param>
        /// <param name="deletePotentialBapRows">if set to <c>true</c> delete any potential bap rows.</param>
        /// <exception cref="Exception"></exception>
        private void BulkUpdateBap(string currIncid, string ihsHabitat,
            HluDataSet.incid_bapDataTable incidBapTable,
            HluDataSet.incid_ihs_matrixRow[] ihsMatrixRows,
            HluDataSet.incid_ihs_formationRow[] ihsFormationRows,
            HluDataSet.incid_ihs_managementRow[] ihsManagementRows,
            HluDataSet.incid_ihs_complexRow[] ihsComplexRows,
            bool deleteOrphanBapRows,
            bool deletePotentialBapRows)
        {
            var mx = ihsMatrixRows.Where(r => r.RowState != DataRowState.Deleted).Select(r => r.matrix);
            string[] ihsMatrixVals = mx.Concat(new string[3 - mx.Count()]).ToArray();
            var fo = ihsFormationRows.Where(r => r.RowState != DataRowState.Deleted).Select(r => r.formation);
            string[] ihsFormationVals = fo.Concat(new string[2 - fo.Count()]).ToArray();
            var mg = ihsManagementRows.Where(r => r.RowState != DataRowState.Deleted).Select(r => r.management);
            string[] ihsManagementVals = mg.Concat(new string[2 - mg.Count()]).ToArray();
            var cx = ihsComplexRows.Where(r => r.RowState != DataRowState.Deleted).Select(r => r.complex);
            string[] ihsComplexVals = cx.Concat(new string[2 - cx.Count()]).ToArray();

            // Get a list of all the primary (mandatory) bap habitats
            IEnumerable<string> primaryBap = _viewModelMain.PrimaryBapEnvironments(ihsHabitat, 
                ihsMatrixVals[0], ihsMatrixVals[1], ihsMatrixVals[2], ihsFormationVals[0], ihsFormationVals[1],
                ihsManagementVals[0], ihsManagementVals[1], ihsComplexVals[0], ihsComplexVals[1]);

            // Get an array of the columns not be updated in the table
            int[] skipOrdinals = new int[3] { 
                _viewModelMain.HluDataset.incid_bap.bap_idColumn.Ordinal, 
                _viewModelMain.HluDataset.incid_bap.incidColumn.Ordinal, 
                _viewModelMain.HluDataset.incid_bap.bap_habitatColumn.Ordinal };

            List<HluDataSet.incid_bapRow> updateRows = new List<HluDataSet.incid_bapRow>();
            HluDataSet.incid_bapRow updateRow;

            // Get all the BAP habitats from the user interface
            IEnumerable<BapEnvironment> beUI = from b in _viewModelMain.IncidBapRowsAuto.Concat(_viewModelMain.IncidBapRowsUser)
                                               group b by b.bap_habitat into habs
                                               select habs.First();

            // Loop through all the BAP habitats from the user interface
            foreach (BapEnvironment be in beUI)
            {
                // Check if the current BAP habitat is primary (mandatory) or secondary
                bool isSecondary = !primaryBap.Contains(be.bap_habitat);

                // Flag the current BAP habitat as secondary
                if (isSecondary) be.MakeSecondary();

                // Get any rows for the current BAP habitat already in the database
                IEnumerable<HluDataSet.incid_bapRow> dbRows =
                    incidBapTable.Where(r => r.bap_habitat == be.bap_habitat);

                switch (dbRows.Count())
                {
                    // If the current BAP habitat is not already in the database
                    case 0: // Insert the newly added BAP habitat
                        HluDataSet.incid_bapRow newRow = _viewModelMain.IncidBapTable.Newincid_bapRow();
                        newRow.ItemArray = be.ToItemArray(_viewModelMain.RecIDs.NextIncidBapId, currIncid);
                        if (be.IsValid(false, isSecondary, newRow)) // reset bulk update mode for full validation of a new row
                            _viewModelMain.HluTableAdapterManager.incid_bapTableAdapter.Insert(newRow);
                        break;
                    // If the current BAP habitat is already in the database
                    case 1: // Update the existing row
                        updateRow = dbRows.ElementAt(0);
                        object[] itemArray = be.ToItemArray();
                        for (int i = 0; i < itemArray.Length; i++)
                        {
                            if ((itemArray[i] != null) && (Array.IndexOf(skipOrdinals, i) == -1))
                                updateRow[i] = itemArray[i];
                        }
                        updateRows.Add(updateRow);
                        break;
                    default: // impossible if rules properly enforced
                        break;
                }
            }

            // Delete any previously primary BAP environments from the database
            if (deleteOrphanBapRows)
            {
                var delRows = incidBapTable.Where(r => !primaryBap.Contains(r.bap_habitat)
                    && BapEnvironment.IsSecondary(r) == false);
                foreach (HluDataSet.incid_bapRow r in delRows)
                    _viewModelMain.HluTableAdapterManager.incid_bapTableAdapter.Delete(r);
            }
            else
            {
                // Change any existing rows in the database that are no longer primary
                // (mandatory) BAP habitats to secondary
                incidBapTable.Where(r => !primaryBap.Contains(r.bap_habitat)
                    && BapEnvironment.IsSecondary(r) == false).ToList().ForEach(delegate(HluDataSet.incid_bapRow r)
                {
                    updateRows.Add(BapEnvironment.MakeSecondary(r));
                });
            }

            // Update the BAP habitat if there are any rows to update
            if (updateRows.Count > 0)
            {
                if (_viewModelMain.HluTableAdapterManager.incid_bapTableAdapter.Update(updateRows.ToArray()) == -1)
                    throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_bap.TableName));
            }

            // Delete any previously secondary BAP environments from the database
            // if they are not in the user interface
            if (deletePotentialBapRows)
            {
                var delRows = incidBapTable.Where(r => !primaryBap.Contains(r.bap_habitat) &&
                    BapEnvironment.IsSecondary(r) == true &&
                    _viewModelMain.IncidBapRowsUser.Count(be => be.bap_habitat == r.bap_habitat) == 0);
                foreach (HluDataSet.incid_bapRow r in delRows)
                    _viewModelMain.HluTableAdapterManager.incid_bapTableAdapter.Delete(r);
            }
        }

        /// <summary>
        /// Update the GIS layer.
        /// </summary>
        /// <param name="incidOrdinal">The incid ordinal.</param>
        /// <param name="incidSelection">The incid selection.</param>
        /// <exception cref="Exception">
        /// Failed to update GIS layer shadow copy
        /// or
        /// Failed to update GIS layer for incid
        /// </exception>
        private void BulkUpdateGis(int incidOrdinal, DataTable incidSelection, bool createHistory, ViewModelWindowMain.Operations operation, DateTime nowDtTm)
        {
            // Get the columns and values to be updated in GIS
            DataColumn[] updateColumns;
            object[] updateValues;
            BulkUpdateGisColumns(out updateColumns, out updateValues);

            string incidMMPolygonsUpdateCmdTemplate;
            List<List<SqlFilterCondition>> incidWhereClause;
            DataTable historyTable = null;

            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            // If there are no columns to update
            if ((updateColumns == null) || (updateColumns.Length == 0))
            {
                // if history is to be created
                if (createHistory)
                {
                    // Build a WHERE clause for all the incids in the DB shadow copy of GIS layer
                    incidWhereClause = ViewModelWindowMainHelpers.IncidSelectionToWhereClause(ViewModelWindowMain.IncidPageSize,
                        _viewModelMain.HluDataset.incid_mm_polygons.incidColumn.Ordinal, _viewModelMain.HluDataset.incid_mm_polygons,
                        incidSelection.AsEnumerable().Select(r => r.Field<string>(incidOrdinal)));

                    // Build a history table for every incid in the DB shadow copy of GIS layer
                    foreach (List<SqlFilterCondition> w in incidWhereClause)
                    {
                        // Retrieve all the GIS rows for the current incid
                        DataTable historyTmp = _viewModelMain.GISApplication.SqlSelect(false, true, _viewModelMain.HistoryColumns, w);

                        // Append history rows to the history table
                        if (historyTmp != null)
                        {
                            if (historyTable == null)
                            {
                                historyTable = historyTmp;
                            }
                            else
                            {
                                foreach (DataRow r in historyTmp.Rows)
                                    historyTable.ImportRow(r);
                            }
                        }
                    }
                }
            }
            else
            {
                // Check if the IHS Summary is one of the columns to update (it should always be)
                int ixIhsSummary = System.Array.IndexOf(updateColumns, _viewModelMain.HluDataset.incid_mm_polygons.ihs_summaryColumn);

                // If the IHS Summary is not one of the columns to update then use a
                // scratch database table to update the GIS features
                if (ixIhsSummary == -1)
                {
                    // Build an UPDATE statement for the DB shadow copy of GIS layer
                    incidMMPolygonsUpdateCmdTemplate = String.Format("UPDATE {0} SET {1} WHERE {2}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_mm_polygons.TableName),
                        updateColumns.Select((c, index) => new string[] { _viewModelMain.DataBase.QuoteIdentifier(c.ColumnName),
                    _viewModelMain.DataBase.QuoteValue(updateValues[index]) }).Aggregate(new StringBuilder(), (sb, a) =>
                                sb.Append(String.Format(", {0} = {1}", a[0], a[1]))).Remove(0, 2), "{0}");

                    // Build a scratch table of the GIS rows to update
                    ScratchDb.WriteSelectionScratchTable(_viewModelMain.GisIDColumns, _viewModelMain.IncidSelection);

                    // Update the GIS layer using the scratch table
                    historyTable = _viewModelMain.GISApplication.UpdateFeatures(updateColumns, updateValues,
                        _viewModelMain.HistoryColumns, ScratchDb.ScratchMdbPath, ScratchDb.ScratchSelectionTable);
                }
                // Otherwise, loop through each incid to individually update the GIS features
                else
                {
                    // Build an UPDATE statement for the DB shadow copy of GIS layer
                    incidMMPolygonsUpdateCmdTemplate = String.Format("UPDATE {0} SET {1} WHERE {2}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_mm_polygons.TableName),
                        updateColumns.Select((c, index) => new string[] { _viewModelMain.DataBase.QuoteIdentifier(c.ColumnName), 
                    c.ColumnName == _viewModelMain.HluDataset.incid_mm_polygons.ihs_summaryColumn.ColumnName ?
                    "{0}" : _viewModelMain.DataBase.QuoteValue(updateValues[index]) }).Aggregate(new StringBuilder(),
                            (sb, a) => sb.Append(String.Format(", {0} = {1}", a[0], a[1]))).Remove(0, 2), "{1}");

                    // Build a WHERE clause for the rows to update in the DB shadow copy of GIS layer
                    incidWhereClause = ViewModelWindowMainHelpers.IncidSelectionToWhereClause(1,
                        _viewModelMain.HluDataset.incid_mm_polygons.incidColumn.Ordinal, _viewModelMain.HluDataset.incid_mm_polygons,
                        incidSelection.AsEnumerable().Select(r => r.Field<string>(incidOrdinal)));

                    // Execute the UPDATE statement for each incid in the DB shadow copy of GIS layer
                    foreach (List<SqlFilterCondition> w in incidWhereClause)
                    {
                        updateValues[ixIhsSummary] = BuildIhsSummary(w[0].Value.ToString());

                        if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format(incidMMPolygonsUpdateCmdTemplate,
                            _viewModelMain.DataBase.QuoteValue(updateValues[ixIhsSummary]),
                            _viewModelMain.DataBase.WhereClause(false, true, true, w)),
                            _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                            throw new Exception("Failed to update GIS layer shadow copy");

                        // Update GIS layer row by row; no need for a joined scratch table
                        DataTable historyTmp = _viewModelMain.GISApplication.UpdateFeatures(updateColumns,
                            updateValues, _viewModelMain.HistoryColumns, w);

                        if (historyTmp == null)
                            throw new Exception(String.Format("Failed to update GIS layer for incid '{0}'", w[0].Value));

                        // Append history rows to the history table
                        if (createHistory)
                        {
                            if (historyTable == null)
                            {
                                historyTable = historyTmp;
                            }
                            else
                            {
                                foreach (DataRow r in historyTmp.Rows)
                                    historyTable.ImportRow(r);
                            }
                        }
                    }
                }
            }
            //---------------------------------------------------------------------

            // Write history for the affected incids
            if (createHistory && (historyTable != null))
            {
                ViewModelWindowMainHistory vmHist = new ViewModelWindowMainHistory(_viewModelMain);
                vmHist.HistoryWrite(null, historyTable, operation, nowDtTm);
            }
        }

        /// <summary>
        /// Build the IHS Summary string for the required incid.
        /// </summary>
        /// <param name="incid">The required incid.</param>
        /// <returns></returns>
        private string BuildIhsSummary(string incid)
        {
            try
            {
                var incidRow = _viewModelMain.IncidTable.Where(r => r.incid == incid);

                //---------------------------------------------------------------------
                // FIX: 078 Bulk update overhaul/improvements.
                // 
                //string ihsHabitat = incidRow.Count() == 1 ? incidRow.ElementAt(0).ihs_habitat : String.Empty;

                // Get the IHS habitat from the user interface
                string ihsHabitat = _viewModelMain.IncidCurrentRow.ihs_habitat;
                //---------------------------------------------------------------------

                // Get the multiplex rows from the database
                _viewModelMain.GetIncidChildRowsDb(new object[] { incid },
                    _viewModelMain.HluTableAdapterManager.incid_ihs_matrixTableAdapter, ref _ihsMatrixTable);

                _viewModelMain.GetIncidChildRowsDb(new object[] { incid },
                    _viewModelMain.HluTableAdapterManager.incid_ihs_formationTableAdapter, ref _ihsFormationTable);

                _viewModelMain.GetIncidChildRowsDb(new object[] { incid },
                    _viewModelMain.HluTableAdapterManager.incid_ihs_managementTableAdapter, ref _ihsManagementTable);

                _viewModelMain.GetIncidChildRowsDb(new object[] { incid },
                    _viewModelMain.HluTableAdapterManager.incid_ihs_complexTableAdapter, ref _ihsComplexTable);

                //---------------------------------------------------------------------
                // FIX: 078 Bulk update overhaul/improvements.
                // 
                // Build the IHS Summary string
                object ihsSummaryValue = ViewModelWindowMainHelpers.IhsSummary(new string[] { ihsHabitat }
                    .Concat(_ihsMatrixTable.Select((ihs, index) =>  
                        !ihs.IsNull(_viewModelMain.HluDataset.incid_ihs_matrix.matrixColumn.Ordinal) ? 
                        ihs[_viewModelMain.HluDataset.incid_ihs_matrix.matrixColumn.Ordinal].ToString() :
                        String.Empty))
                    .Concat(_ihsFormationTable.Select((ihs, index) =>
                        !ihs.IsNull(_viewModelMain.HluDataset.incid_ihs_formation.formationColumn.Ordinal) ?
                        ihs[_viewModelMain.HluDataset.incid_ihs_formation.formationColumn.Ordinal].ToString() :
                        String.Empty))
                    .Concat(_ihsManagementTable.Select((ihs, index) =>
                        !ihs.IsNull(_viewModelMain.HluDataset.incid_ihs_management.managementColumn.Ordinal) ?
                        ihs[_viewModelMain.HluDataset.incid_ihs_management.managementColumn.Ordinal].ToString() :
                        String.Empty))
                    .Concat(_ihsComplexTable.Select((ihs, index) =>
                        !ihs.IsNull(_viewModelMain.HluDataset.incid_ihs_complex.complexColumn.Ordinal) ?
                        ihs[_viewModelMain.HluDataset.incid_ihs_complex.complexColumn.Ordinal].ToString() :
                        String.Empty)).ToArray());
                //---------------------------------------------------------------------
                
                // Return the IHS summary string
                if (ihsSummaryValue != null)
                    return ihsSummaryValue.ToString();
                else
                    return null;
            }
            catch { throw; }
            finally
            {
                _ihsMatrixTable.Clear();
                _ihsFormationTable.Clear();
                _ihsManagementTable.Clear();
                _ihsComplexTable.Clear();
            }
        }

        /// <summary>
        /// Set the columns and values to be updated in GIS.
        /// </summary>
        /// <param name="updateColumns">The columns to update.</param>
        /// <param name="updateValues">The values to update.</param>
        private void BulkUpdateGisColumns(out DataColumn[] updateColumns, out object[] updateValues)
        {
            List<DataColumn> updateColumnList = new List<DataColumn>();
            List<object> updateValueList = new List<object>();
            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // FIX: 005 Always save all (both) of the history columns
            //
            // Check if a new IHS habitat has been set
            if (!_viewModelMain.IncidCurrentRow.Isihs_habitatNull())
            {
                // Add the IHS Category column and value
                updateColumnList.Add(_viewModelMain.HluDataset.incid_mm_polygons.ihs_categoryColumn);
                updateValueList.Add(_viewModelMain.HluDataset.lut_ihs_habitat
                    .Single(h => h.code == _viewModelMain.IncidCurrentRow.ihs_habitat).category);

                // Add the IHS Summary column and a null value (to be set later)
                updateColumnList.Add(_viewModelMain.HluDataset.incid_mm_polygons.ihs_summaryColumn);
                updateValueList.Add(null);
            }
            //---------------------------------------------------------------------

            // Add any other history columns to be updated
            var addCols = _viewModelMain.HistoryColumns.Where(h => 
                _viewModelMain.HluDataset.incid_mm_polygons.Columns.Contains(h.ColumnName) &&
                _viewModelMain.IncidCurrentRow.Table.Columns.Contains(h.ColumnName) && 
                !_viewModelMain.IncidCurrentRow.IsNull(h.ColumnName));

            if (addCols.Count() > 0)
            {
                updateColumnList.AddRange(addCols);
                updateValueList.AddRange(_viewModelMain.HistoryColumns.Where(h =>
                    _viewModelMain.HluDataset.incid_mm_polygons.Columns.Contains(h.ColumnName) &&
                    _viewModelMain.IncidCurrentRow.Table.Columns.Contains(h.ColumnName) &&
                    !_viewModelMain.IncidCurrentRow.IsNull(h.ColumnName))
                    .Select(h => _viewModelMain.IncidCurrentRow[h.ColumnName]));
            }

            updateColumns = updateColumnList.ToArray();
            updateValues = updateValueList.ToArray();
        }

        /// <summary>
        /// Bulk updates the required child table using a DB connection.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="R"></typeparam>
        /// <param name="deleteExistingRows">if set to <c>1</c> delete any existing rows.</param>
        /// <param name="newRows">The new rows to be inserted.</param>
        /// <param name="dbRows">The database rows.</param>
        /// <exception cref="ArgumentException">
        /// dbRows
        /// or
        /// newRows
        /// or
        /// Table must have a single column primary key of type Int32 - dbRows
        /// </exception>
        /// <exception cref="Exception">
        /// </exception>
        private void BulkUpdateDbChildTable<T, R>(int deleteExistingRows, R[] newRows, ref T dbRows)
            where T : DataTable, new()
            where R : DataRow
        {
            if (dbRows == null) throw new ArgumentException("dbRows");
            if (newRows == null) throw new ArgumentException("newRows");

            if ((dbRows.PrimaryKey.Length != 1) || (dbRows.PrimaryKey[0].DataType != typeof(Int32)))
                throw new ArgumentException("Table must have a single column primary key of type Int32", "dbRows");

            DataColumn[] pk = dbRows.PrimaryKey;
            int pkOrdinal = pk[0].Ordinal;

            // remove any new rows with no or duplicate data
            R[] dbRowsEnum = dbRows.AsEnumerable().Select(r => (R)r).ToArray();
            DataColumn[] cols = dbRows.Columns.Cast<DataColumn>().Where(c => c.Ordinal != pkOrdinal).ToArray();
            R[] newRowsNoDups = (from nr in newRows
                                 where cols.Count(col => !nr.IsNull(col.Ordinal)) > 0 &&
                                     dbRowsEnum.Count(dr => cols.Count(c => !dr[c.Ordinal].Equals(nr[c.Ordinal])) == 0) == 0
                                 select nr).OrderBy(r => r[pkOrdinal]).ToArray();

            // number of non-blank controls corresponding to child table dbRows
            int numRowsNew = newRowsNoDups.Count();

            if (numRowsNew == 0) return;

            // number of child rows in database for current incid
            int numRowsDb = dbRows.Rows.Count;

            // update existing rows matching controls to rows in order of PK values (same as UI display order)
            string updateCommand = String.Format("UPDATE {0} SET ", _viewModelMain.DataBase.QualifyTableName(dbRows.TableName));

            int limit = numRowsDb <= numRowsNew ? numRowsDb : numRowsNew;
            for (int i = 0; i < limit; i++)
            {
                R newRow = newRowsNoDups[i];
                R dbRow = dbRowsEnum[(int)newRowsNoDups[i][pkOrdinal]];
                if (_viewModelMain.DataBase.ExecuteNonQuery(dbRows.Columns.Cast<DataColumn>()
                    .Where(c => pk.Count(k => k.Ordinal == c.Ordinal) == 0 && !newRow.IsNull(c.Ordinal))
                    .Aggregate(new StringBuilder(), (sb, c) => sb.Append(String.Format(", {0} = {1}",
                        _viewModelMain.DataBase.QuoteIdentifier(c.ColumnName), 
                        _viewModelMain.DataBase.QuoteValue(newRow[c.Ordinal])))).Remove(0, 2)
                        .Append(String.Format(" WHERE {0} = {1}", dbRows.PrimaryKey.Aggregate(
                        new StringBuilder(), (sb, c) => sb.Append(String.Format("AND {0} = {1}",
                            _viewModelMain.DataBase.QuoteIdentifier(c.ColumnName), 
                            _viewModelMain.DataBase.QuoteValue(dbRow[c.Ordinal])))).Remove(0, 4)))
                            .Insert(0, updateCommand).ToString(), _viewModelMain.DataBase.Connection.ConnectionTimeout, 
                            CommandType.Text) == -1)
                    throw new Exception(String.Format("Failed to update table '{0}'.", dbRows.TableName));
            }

            if (numRowsNew > numRowsDb) // user entered new values
            {
                string recordIdPropertyName = dbRows.TableName.Split('_')
                    .Aggregate(new StringBuilder(), (sb, s) => sb.Append(char.ToUpper(s[0])).Append(s.Substring(1)))
                    .Insert(0, "Next").Append("Id").ToString();

                PropertyInfo recordIDPropInfo = typeof(RecordIds).GetProperty(recordIdPropertyName);

                string insertCommand = String.Format("INSERT INTO {0} ({1}) VALUES ({2})",
                    _viewModelMain.DataBase.QualifyTableName(dbRows.TableName), 
                    _viewModelMain.DataBase.QuoteValue("{0}"), 
                    _viewModelMain.DataBase.QuoteValue("{1}"));

                for (int i = numRowsDb; i < numRowsNew; i++)
                {
                    R newRow = newRows[i];

                    StringBuilder columnNames = new StringBuilder();
                    StringBuilder columnValues = new StringBuilder();
                    for (int j = 0; j < dbRows.Columns.Count; j++)
                    {
                        columnNames.Append(String.Format(", {0}", 
                            _viewModelMain.DataBase.QuoteIdentifier(dbRows.Columns[i].ColumnName)));
                        if (j == pkOrdinal)
                        {
                            int newPK = (int)recordIDPropInfo.GetValue(_viewModelMain.RecIDs, null);
                            columnValues.Append(String.Format(", {0}", newPK));
                        }
                        else
                        {
                            columnValues.Append(String.Format(", {0}", _viewModelMain.DataBase.QuoteValue(newRow[i])));
                        }
                    }

                    if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format(insertCommand, columnNames, columnValues),
                        _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                        throw new Exception(String.Format("Failed to insert into table {0}.", dbRows.TableName));
                }
            }
            else if ((deleteExistingRows == (int)DeleteMultiplexCodesAction.All) && (numRowsDb > numRowsNew))
            {
                StringBuilder deleteCommand = new StringBuilder(String.Format(
                    "DELETE FROM {0} WHERE ", _viewModelMain.DataBase.QualifyTableName(dbRows.TableName)));

                for (int i = numRowsNew; i < numRowsDb; i++)
                {
                    R dbRow = (R)dbRows.Rows[i];

                    deleteCommand.Append(String.Format("{0} = {1}", dbRows.PrimaryKey.Aggregate(
                        new StringBuilder(), (sb, c) => sb.Append(String.Format("AND {0} = {1}",
                            _viewModelMain.DataBase.QuoteIdentifier(c.ColumnName), 
                            _viewModelMain.DataBase.QuoteValue(dbRow[c.Ordinal])))))).Remove(0, 4);

                    if (_viewModelMain.DataBase.ExecuteNonQuery(deleteCommand.ToString(), 
                        _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                        throw new Exception(String.Format("Failed to delete from table {0}.", dbRows.TableName));
                }
            }

            if (newRowsNoDups.Length > 0)
            {
                newRowsNoDups[0].Table.AcceptChanges();
                dbRows = (T)newRowsNoDups[0].Table;
            }
        }

        /// <summary>
        /// Bulk updates the required child table using an ADO connection.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="R"></typeparam>
        /// <param name="deleteExistingRows">if set to <c>1</c> delete any existing rows.</param>
        /// <param name="newRows">The new rows to be inserted.</param>
        /// <param name="dbRows">The database rows.</param>
        /// <exception cref="ArgumentException">
        /// dbRows
        /// or
        /// newRows
        /// or
        /// Table must have a single column primary key of type Int32 - dbRows
        /// </exception>
        /// <exception cref="Exception">
        /// </exception>
        private void BulkUpdateAdoChildTable<T, R>(int deleteExistingRows, R[] newRows, ref T dbRows,
            HluTableAdapter<T, R> adapter)
            where T : DataTable, new()
            where R : DataRow
        {
            if (dbRows == null) throw new ArgumentException("dbRows");
            if (newRows == null) throw new ArgumentException("newRows");
            if (adapter == null) throw new ArgumentException("adapter");

            // Check the data table primary key is an integer
            if ((dbRows.PrimaryKey.Length != 1) || (dbRows.PrimaryKey[0].DataType != typeof(Int32)))
                throw new ArgumentException("Table must have a single column primary key of type Int32", "dbRows");

            // Get the primary key column ordinal
            int pkOrdinal = dbRows.PrimaryKey[0].Ordinal;

            // Create an enumerable of all the target data table rows
            R[] dbRowsEnum = dbRows.AsEnumerable().Select(r => (R)r).ToArray();

            // Get an array of all the columns in the target data table (except the primary key)
            DataColumn[] cols = dbRows.Columns.Cast<DataColumn>().Where(c => c.Ordinal != pkOrdinal).ToArray();

            // Select only new rows with non-blank data or non-duplicate data
            R[] newRowsNoDups = (from nr in newRows
                                 where cols.Count(col => !nr.IsNull(col.Ordinal)) > 0 &&
                                     dbRowsEnum.Count(dr => cols.Count(c => !dr[c.Ordinal].Equals(nr[c.Ordinal])) == 0) == 0
                                 select nr).OrderBy(r => r[pkOrdinal]).ToArray();

            // Get the number of non-blank row with non-duplicate data corresponding to child table dbRows
            int numRowsNew = newRowsNoDups.Count();

            // Exit if no existing rows are to be retained and there are no new rows to add
            if ((deleteExistingRows == (int)DeleteMultiplexCodesAction.All) && (numRowsNew == 0)) return;

            // Select only existing data table rows not in the new rows
            R[] oldRows = (from dr in dbRowsEnum
                                 where cols.Count(col => !dr.IsNull(col.Ordinal)) > 0 &&
                                     newRows.Count(nr => cols.Count(c => nr[c.Ordinal].Equals(dr[c.Ordinal])) == cols.Count()) == 0
                                 select dr).OrderBy(r => r[pkOrdinal]).ToArray();

            // Delete all the existing database rows
            System.Array.ForEach(dbRowsEnum.ToArray(),
                new Action<R>(r => adapter.Delete(r)));

            // Set the property name for the primary key
            string recordIdPropertyName = dbRows.TableName.Split('_')
                .Aggregate(new StringBuilder(), (sb, s) => sb.Append(char.ToUpper(s[0])).Append(s.Substring(1)))
                .Insert(0, "Next").Append("Id").ToString();

            // Set the property info for the primary key property name
            PropertyInfo recordIDPropInfo = typeof(RecordIds).GetProperty(recordIdPropertyName);

            // Insert any new rows into the data table
            int pkMax = -1;
            foreach (R newRow in newRows)
            {
                // Store the new row primary key
                int pkValue = (int)newRow[pkOrdinal];
                pkMax = pkValue;

                // Set the primary key to the next value
                if (!dbRows.Columns[pkOrdinal].AutoIncrement)
                    newRow[pkOrdinal] = (int)recordIDPropInfo.GetValue(_viewModelMain.RecIDs, null);

                // Insert the new row
                adapter.Insert(newRow);

                // Restore the new row primary key
                newRow[pkOrdinal] = pkValue;
            }

            //---------------------------------------------------------------------
            // FIX: 078 Bulk update overhaul/improvements.
            // 
            // Get the maximum number of child rows for the current table
            int maxRowsDb = 0;
            switch (dbRows.TableName.ToLower())
            {
                case "incid_ihs_matrix":
                    maxRowsDb = 3;
                    break;
                case "incid_ihs_formation":
                    maxRowsDb = 2;
                    break;
                case "incid_ihs_management":
                    maxRowsDb = 2;
                    break;
                case "incid_ihs_complex":
                    maxRowsDb = 2;
                    break;
                case "incid_sources":
                    maxRowsDb = 3;
                    break;
            }

            // Re-insert any old rows not in the new rows
            if (deleteExistingRows == (int)DeleteMultiplexCodesAction.None)
            {
                foreach (R oldRow in oldRows)
                {
                    // Increment the primary key integer
                    pkMax = pkMax + 1;

                    // If there are still spare rows
                    if (pkMax < maxRowsDb)
                    {
                        // Store the new row primary key
                        int pkValue = (int)oldRow[pkOrdinal];

                        // Set the primary key to the next value
                        if (!dbRows.Columns[pkOrdinal].AutoIncrement)
                            oldRow[pkOrdinal] = (int)recordIDPropInfo.GetValue(_viewModelMain.RecIDs, null);
                        else
                            oldRow[pkOrdinal] = pkMax;

                        // Insert the new row
                        adapter.Insert(oldRow);

                        // Restore the new row primary key
                        oldRow[pkOrdinal] = pkValue;
                    }
                }
            }

            // Accept any changes outstanding for the table
            if (pkMax >= 0)
            {
                if (newRows.Count() > 0 )
                    newRows[0].Table.AcceptChanges();
                else
                    oldRows[0].Table.AcceptChanges();
                //dbRows = (T)newRows[0].Table;
            }
            //---------------------------------------------------------------------
        }

        #endregion

    }

}
