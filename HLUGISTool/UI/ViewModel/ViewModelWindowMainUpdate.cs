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
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using HLU.Data;
using HLU.Data.Model;

namespace HLU.UI.ViewModel
{
    class ViewModelWindowMainUpdate
    {
        ViewModelWindowMain _viewModelMain;

        public ViewModelWindowMainUpdate(ViewModelWindowMain viewModelMain)
        {
            _viewModelMain = viewModelMain;
        }

        /// <summary>
        /// Writes changes made to current incid back to database and GIS layer.
        /// Also synchronizes shadow copy of GIS layer in DB and writes history.
        /// </summary>
        internal bool Update()
        {
            _viewModelMain.DataBase.BeginTransaction(true, IsolationLevel.ReadCommitted);

            try
            {
                _viewModelMain.ChangeCursor(Cursors.Wait, "Saving ...");

                int incidCurrRowIx = _viewModelMain.IncidCurrentRowIndex;

                //---------------------------------------------------------------------
                // FIXED: KI97 (Last modified date and user)
                // CHANGED: CR3 (IHS version)
                // Previously only changes to fields on the incid table triggered the
                // last modified date & user fields to be updated.
                // Update the last modified date & user fields and the ihs version on
                // the incid table regardless of which attributes have been changed.
                //if (_viewModelMain.IsDirtyIncid())
                //{
                IncidCurrentRowDerivedValuesUpdate(_viewModelMain);

                //---------------------------------------------------------------------
                // FIX: 028 Only update DateTime fields to whole seconds
                // Fractions of a second can cause rounding differences when
                // comparing DateTime fields later in some databases.
                DateTime currDtTm = DateTime.Now;
                DateTime nowDtTm = new DateTime(currDtTm.Year, currDtTm.Month, currDtTm.Day, currDtTm.Hour, currDtTm.Minute, currDtTm.Second, DateTimeKind.Local);
                //---------------------------------------------------------------------
                _viewModelMain.IncidCurrentRow.last_modified_date = nowDtTm;
                _viewModelMain.IncidCurrentRow.last_modified_user_id = _viewModelMain.UserID;
                _viewModelMain.IncidCurrentRow.ihs_version = _viewModelMain.IhsVersion;

                //---------------------------------------------------------------------
                // FIX: 075 Reset OSMM update flag after manual updates.
                // 
                if ((_viewModelMain.IncidOSMMUpdatesRows.Length > 0) &&
                   (_viewModelMain.ResetUpdatesFlag) &&
                   (_viewModelMain.IncidOSMMUpdatesRows[0].status_flag > 0))
                {
                    // Set the update flag to "Ignored"
                    _viewModelMain.IncidOSMMUpdatesRows[0].status_flag = -2;
                    _viewModelMain.IncidOSMMUpdatesRows[0].last_modified_date = nowDtTm;
                    _viewModelMain.IncidOSMMUpdatesRows[0].last_modified_user_id = _viewModelMain.UserID;
                }
                //---------------------------------------------------------------------

                if (_viewModelMain.HluTableAdapterManager.incidTableAdapter.Update(
                    (HluDataSet.incidDataTable)_viewModelMain.HluDataset.incid.GetChanges()) == -1)
                    throw new Exception(String.Format("Failed to update '{0}' table.",
                        _viewModelMain.HluDataset.incid.TableName));
                //}
                //---------------------------------------------------------------------

                if ((_viewModelMain.IncidIhsMatrixRows != null) && _viewModelMain.IsDirtyIncidIhsMatrix())
                {
                    if (_viewModelMain.HluTableAdapterManager.incid_ihs_matrixTableAdapter.Update(
                        (HluDataSet.incid_ihs_matrixDataTable)_viewModelMain.HluDataset.incid_ihs_matrix.GetChanges()) == -1)
                        throw new Exception(String.Format("Failed to update '{0}' table.",
                            _viewModelMain.HluDataset.incid_ihs_matrix.TableName));
                }

                if ((_viewModelMain.IncidIhsFormationRows != null) && _viewModelMain.IsDirtyIncidIhsFormation())
                {
                    if (_viewModelMain.HluTableAdapterManager.incid_ihs_formationTableAdapter.Update(
                        (HluDataSet.incid_ihs_formationDataTable)_viewModelMain.HluDataset.incid_ihs_formation.GetChanges()) == -1)
                        throw new Exception(String.Format("Failed to update '{0}' table.",
                            _viewModelMain.HluDataset.incid_ihs_formation.TableName));
                }

                if ((_viewModelMain.IncidIhsManagementRows != null) && _viewModelMain.IsDirtyIncidIhsManagement())
                {
                    if (_viewModelMain.HluTableAdapterManager.incid_ihs_managementTableAdapter.Update(
                        (HluDataSet.incid_ihs_managementDataTable)_viewModelMain.HluDataset.incid_ihs_management.GetChanges()) == -1)
                        throw new Exception(String.Format("Failed to update '{0}' table.",
                            _viewModelMain.HluDataset.incid_ihs_management.TableName));
                }

                if ((_viewModelMain.IncidIhsComplexRows != null) && _viewModelMain.IsDirtyIncidIhsComplex())
                {
                    if (_viewModelMain.HluTableAdapterManager.incid_ihs_complexTableAdapter.Update(
                        (HluDataSet.incid_ihs_complexDataTable)_viewModelMain.HluDataset.incid_ihs_complex.GetChanges()) == -1)
                        throw new Exception(String.Format("Failed to update '{0}' table.",
                            _viewModelMain.HluDataset.incid_ihs_complex.TableName));
                }

                if (_viewModelMain.IsDirtyIncidBap()) UpdateBap();

                if (_viewModelMain.IncidSourcesRows != null)
                {
                    int j = 0;
                    for (int i = 0; i < _viewModelMain.IncidSourcesRows.Length; i++)
                        if (_viewModelMain.IncidSourcesRows[i] != null)
                            _viewModelMain.IncidSourcesRows[i].sort_order = ++j;

                    if (_viewModelMain.HluTableAdapterManager.incid_sourcesTableAdapter.Update(
                        (HluDataSet.incid_sourcesDataTable)_viewModelMain.IncidSourcesTable.GetChanges()) == -1)
                        throw new Exception(String.Format("Failed to update {0} table.", 
                            _viewModelMain.HluDataset.incid_sources.TableName));
                }

                if ((_viewModelMain.IncidOSMMUpdatesRows != null) && _viewModelMain.IsDirtyIncidOSMMUpdates())
                {
                    if (_viewModelMain.HluTableAdapterManager.incid_osmm_updatesTableAdapter.Update(
                        (HluDataSet.incid_osmm_updatesDataTable)_viewModelMain.HluDataset.incid_osmm_updates.GetChanges()) == -1)
                        throw new Exception(String.Format("Failed to update '{0}' table.",
                            _viewModelMain.HluDataset.incid_osmm_updates.TableName));
                }

                // update all GIS rows corresponding to this incid
                List<SqlFilterCondition> incidCond = new List<SqlFilterCondition>(new SqlFilterCondition[] { 
                    new SqlFilterCondition(_viewModelMain.HluDataset.incid_mm_polygons, 
                        _viewModelMain.HluDataset.incid_mm_polygons.incidColumn, _viewModelMain.Incid) });

                var q = _viewModelMain.HluDataset.lut_ihs_habitat
                    .Where(r => r.code == _viewModelMain.IncidCurrentRow.ihs_habitat);
                string ihsHabitatCategory = q.Count() == 1 ? q.ElementAt(0).category : null;

                DataTable historyTable = _viewModelMain.GISApplication.UpdateFeatures(new DataColumn[] { 
                    _viewModelMain.HluDataset.incid_mm_polygons.ihs_categoryColumn,
                    _viewModelMain.HluDataset.incid_mm_polygons.ihs_summaryColumn },
                    new object[] { ihsHabitatCategory, _viewModelMain.IncidIhsSummary },
                    _viewModelMain.HistoryColumns, incidCond);

                if (historyTable == null)
                    throw new Exception("Error updating GIS layer.");
                else if (historyTable.Rows.Count == 0)
                    throw new Exception("No GIS features to update.");

                // likewise update DB shadow copy of GIS layer
                if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format("UPDATE {0} SET {1} = {3}, {2} = {4} WHERE {5}",
                    _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_mm_polygons.TableName),
                    _viewModelMain.DataBase.QuoteIdentifier(
                        _viewModelMain.HluDataset.incid_mm_polygons.ihs_categoryColumn.ColumnName),
                    _viewModelMain.DataBase.QuoteIdentifier(
                        _viewModelMain.HluDataset.incid_mm_polygons.ihs_summaryColumn.ColumnName),
                    _viewModelMain.DataBase.QuoteValue(ihsHabitatCategory),
                    _viewModelMain.DataBase.QuoteValue(_viewModelMain.IncidIhsSummary),
                    _viewModelMain.DataBase.WhereClause(false, true, true, incidCond)),
                    _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                    throw new Exception("Failed to update database copy of GIS layer.");

                // save history returned from GIS
                Dictionary<int, string> fixedValues = new Dictionary<int, string>();
                fixedValues.Add(_viewModelMain.HluDataset.history.incidColumn.Ordinal, _viewModelMain.Incid);
                ViewModelWindowMainHistory vmHist = new ViewModelWindowMainHistory(_viewModelMain);
                vmHist.HistoryWrite(fixedValues, historyTable, ViewModelWindowMain.Operations.AttributeUpdate);

                _viewModelMain.DataBase.CommitTransaction();
                _viewModelMain.HluDataset.AcceptChanges();
                _viewModelMain.Saved = true;

                _viewModelMain.IncidRowCount(true);
                _viewModelMain.IncidCurrentRowIndex = incidCurrRowIx;

                return true;
            }
            catch (Exception ex)
            {
                _viewModelMain.DataBase.RollbackTransaction();
                if (_viewModelMain.HaveGisApp)
                {
                    _viewModelMain.Saved = false;
                    MessageBox.Show("Your changes could not be saved. The error message returned was:\n\n" +
                        ex.Message, "HLU: Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
                else
                {
                    return true;
                }
            }
            finally
            {
                _viewModelMain.SavingAttempted = true;
                _viewModelMain.Saving = false;
                _viewModelMain.ChangeCursor(Cursors.Arrow, null);
            }
        }

        /// <summary>
        /// Updates BAP environment rows corresponding to current incid.
        /// </summary>
        private void UpdateBap()
        {
            if (_viewModelMain.IncidBapRowsAuto == null)
                _viewModelMain.IncidBapRowsAuto = new ObservableCollection<BapEnvironment>();
            if (_viewModelMain.IncidBapHabitatsUser == null)
                _viewModelMain.IncidBapHabitatsUser = new ObservableCollection<BapEnvironment>();

            // remove duplicate codes
            IEnumerable<BapEnvironment> beAuto = from b in _viewModelMain.IncidBapRowsAuto
                                                 group b by b.bap_habitat into habs
                                                 select habs.First();

            IEnumerable<BapEnvironment> beUser = from b in _viewModelMain.IncidBapHabitatsUser
                                                 where beAuto.Count(a => a.bap_habitat == b.bap_habitat) == 0
                                                 group b by b.bap_habitat into habs
                                                 select habs.First();

            var currentBapRows = beAuto.Concat(beUser);

            List<HluDataSet.incid_bapRow> newRows = new List<HluDataSet.incid_bapRow>();
            List<HluDataSet.incid_bapRow> updateRows = new List<HluDataSet.incid_bapRow>();
            HluDataSet.incid_bapRow updateRow;

            foreach (BapEnvironment be in currentBapRows)
            {
                if (be.bap_id == -1) // new BAP environment
                {
                    be.bap_id = _viewModelMain.RecIDs.NextIncidBapId;
                    be.incid = _viewModelMain.Incid;
                    HluDataSet.incid_bapRow newRow = _viewModelMain.IncidBapTable.Newincid_bapRow();
                    newRow.ItemArray = be.ToItemArray();
                    newRows.Add(newRow);
                }
                // Get the new values for every updated bap row from the bap data grid.
                else if ((updateRow = UpdateIncidBapRow(be)) != null)
                {
                    // If a row is returned from the data grid add it to the list
                    // of updated rows.
                    updateRows.Add(updateRow);
                }
            }

            // Delete any rows that haven't been marked as deleted but are
            // no longer in the current rows.
            _viewModelMain.IncidBapRows.Where(r => r.RowState != DataRowState.Deleted &&
                currentBapRows.Count(g => g.bap_id == r.bap_id) == 0).ToList()
                .ForEach(delegate(HluDataSet.incid_bapRow row) { row.Delete(); });
            
            // Update the table to remove the deleted rows.
            if (_viewModelMain.HluTableAdapterManager.incid_bapTableAdapter.Update(
                _viewModelMain.IncidBapRows.Where(r => r.RowState == DataRowState.Deleted).ToArray()) == -1)
                throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_bap.TableName));

            // If there are any rows that have been updated.
            if (updateRows.Count > 0)
            {
                // Update the table to update the updated rows.
                if (_viewModelMain.HluTableAdapterManager.incid_bapTableAdapter.Update(updateRows.ToArray()) == -1)
                    throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_bap.TableName));
            }

            // Insert the new rows into the table.
            foreach (HluDataSet.incid_bapRow r in newRows)
                _viewModelMain.HluTableAdapterManager.incid_bapTableAdapter.Insert(r);
        }

        /// <summary>
        /// Writes the values from a BapEnvironment object bound to the BAP data grids into the corresponding incid_bap DataRow.
        /// </summary>
        /// <param name="be">BapEnvironment object bound to data grid on form.</param>
        /// <returns>Updated incid_bap row, or null if no corresponding row was found.</returns>
        private HluDataSet.incid_bapRow UpdateIncidBapRow(BapEnvironment be)
        {
            var q = _viewModelMain.IncidBapRows.Where(r => r.RowState != DataRowState.Deleted && r.bap_id == be.bap_id);
            if (q.Count() == 1)
            {
                if (!be.IsValid()) return null;
                HluDataSet.incid_bapRow oldRow = q.ElementAt(0);
                object[] itemArray = be.ToItemArray();
                for (int i = 0; i < itemArray.Length; i++)
                    oldRow[i] = itemArray[i];
                return oldRow;
            }
            return null;
        }

        /// <summary>
        /// Updates those columns of IncidCurrentRow in main view model that are not directly updated 
        /// by properties (to enable undo if update cancelled).
        /// Called here and by bulk update.
        /// </summary>
        /// <param name="viewModelMain">Reference to main window view model.</param>
        internal static void IncidCurrentRowDerivedValuesUpdate(ViewModelWindowMain viewModelMain)
        {
            viewModelMain.IncidCurrentRow.ihs_habitat = viewModelMain.IncidIhsHabitat;
            viewModelMain.IncidCurrentRow.last_modified_user_id = viewModelMain.IncidLastModifiedUserId;
            viewModelMain.IncidCurrentRow.last_modified_date = viewModelMain.IncidLastModifiedDateVal;
        }

        internal void UpdateIncidModifiedColumns(string incid)
        {
            //---------------------------------------------------------------------
            // FIX: 028 Only update DateTime fields to whole seconds
            // Fractions of a second can cause rounding differences when
            // comparing DateTime fields later in some databases.
            DateTime currDtTm = DateTime.Now;
            DateTime nowDtTm = new DateTime(currDtTm.Year, currDtTm.Month, currDtTm.Day, currDtTm.Hour, currDtTm.Minute, currDtTm.Second, DateTimeKind.Local);
            //---------------------------------------------------------------------
            if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format("UPDATE {0} SET {1} = {2}, {3} = {4} WHERE {5} = {6}",
                _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid.TableName),
                _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid.last_modified_dateColumn.ColumnName),
                _viewModelMain.DataBase.QuoteValue(nowDtTm),
                _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid.last_modified_user_idColumn.ColumnName),
                _viewModelMain.DataBase.QuoteValue(_viewModelMain.UserID),
                _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid.incidColumn.ColumnName),
                _viewModelMain.DataBase.QuoteValue(incid)),
                _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                throw new Exception("Failed to update incid table.");
        }
    }
}
