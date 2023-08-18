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

        /// <summary>
        /// Initializes a new instance of the <see cref="ViewModelWindowMainUpdate"/> class.
        /// </summary>
        /// <param name="viewModelMain">The view model main.</param>
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

                // Store row index for reloading the row after the update
                int incidCurrRowIx = _viewModelMain.IncidCurrentRowIndex;

                // Previously only changes to fields on the incid table triggered the
                // last modified date & user fields to be updated.
                // Update the incid table regardless of which attributes have changed.
                IncidCurrentRowDerivedValuesUpdate(_viewModelMain);

                // Update the DateTime fields to whole seconds.
                // Fractions of a second can cause rounding differences when
                // comparing DateTime fields later in some databases.
                DateTime currDtTm = DateTime.Now;
                DateTime nowDtTm = new DateTime(currDtTm.Year, currDtTm.Month, currDtTm.Day, currDtTm.Hour, currDtTm.Minute, currDtTm.Second, DateTimeKind.Local);
                _viewModelMain.IncidCurrentRow.last_modified_date = nowDtTm;
                _viewModelMain.IncidCurrentRow.last_modified_user_id = _viewModelMain.UserID;

                // Update the incid row
                if (_viewModelMain.HluTableAdapterManager.incidTableAdapter.Update(
                    (HluDataSet.incidDataTable)_viewModelMain.HluDataset.incid.GetChanges()) == -1)
                    throw new Exception(String.Format("Failed to update '{0}' table.",
                        _viewModelMain.HluDataset.incid.TableName));

                //TODO: Update - Check incid_condition update
                // Update condition rows
                if ((_viewModelMain.IncidConditionRows != null) && _viewModelMain.IsDirtyIncidCondition())
                {
                    if (_viewModelMain.HluTableAdapterManager.incid_conditionTableAdapter.Update(
                        (HluDataSet.incid_conditionDataTable)_viewModelMain.HluDataset.incid_condition.GetChanges()) == -1)
                        throw new Exception(String.Format("Failed to update '{0}' table.",
                            _viewModelMain.HluDataset.incid_condition.TableName));
                }

                // Update the secondary rows
                if (_viewModelMain.IsDirtyIncidSecondary()) UpdateSecondary();

                // Update the BAP rows
                if (_viewModelMain.IsDirtyIncidBap()) UpdateBap();

                // Update the source rows
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

                //---------------------------------------------------------------------
                // CHANGED: CR49 Process proposed OSMM Updates
                // 
                // If there are OSMM update rows for this incid, and
                // if the OSMM update status is to be reset after manual
                // updates, and if the OSMM update status > 0 (proposed)
                // or status = 0 (pending) ...
                if ((_viewModelMain.IncidOSMMUpdatesRows.Length > 0) &&
                   (_viewModelMain.ResetOSMMUpdatesStatus) &&
                   (_viewModelMain.IncidOSMMUpdatesRows[0].status >= 0))
                {
                    // Set the update flag to "Ignored"
                    _viewModelMain.IncidOSMMUpdatesRows[0].status = -2;
                    _viewModelMain.IncidOSMMUpdatesRows[0].last_modified_date = nowDtTm;
                    _viewModelMain.IncidOSMMUpdatesRows[0].last_modified_user_id = _viewModelMain.UserID;
                }
                //---------------------------------------------------------------------

                // Update the OSMM Update rows
                if ((_viewModelMain.IncidOSMMUpdatesRows != null) && _viewModelMain.IsDirtyIncidOSMMUpdates())
                {
                    if (_viewModelMain.HluTableAdapterManager.incid_osmm_updatesTableAdapter.Update(
                        (HluDataSet.incid_osmm_updatesDataTable)_viewModelMain.HluDataset.incid_osmm_updates.GetChanges()) == -1)
                        throw new Exception(String.Format("Failed to update '{0}' table.",
                            _viewModelMain.HluDataset.incid_osmm_updates.TableName));
                }

                //---------------------------------------------------------------------
                // Update all of the GIS rows corresponding to this incid
                //---------------------------------------------------------------------

                // Set the SQL condition for the update
                List<SqlFilterCondition> incidCond = new List<SqlFilterCondition>(new SqlFilterCondition[] { 
                    new SqlFilterCondition(_viewModelMain.HluDataset.incid_mm_polygons, 
                        _viewModelMain.HluDataset.incid_mm_polygons.incidColumn, _viewModelMain.Incid) });

                // Update the GIS layer
                DataTable historyTable = _viewModelMain.GISApplication.UpdateFeatures(new DataColumn[] { 
                    _viewModelMain.HluDataset.incid_mm_polygons.habprimaryColumn,
                    _viewModelMain.HluDataset.incid_mm_polygons.habsecondColumn,
                    _viewModelMain.HluDataset.incid_mm_polygons.determqtyColumn,
                    _viewModelMain.HluDataset.incid_mm_polygons.interpqtyColumn,
                    _viewModelMain.HluDataset.incid_mm_polygons.interpcomColumn },
                    new object[] { _viewModelMain.IncidPrimary != null ? _viewModelMain.IncidPrimary : "",
                        _viewModelMain.IncidSecondarySummary != null ? _viewModelMain.IncidSecondarySummary : "",
                        _viewModelMain.IncidQualityDetermination != null ? _viewModelMain.IncidQualityDetermination : "",
                        _viewModelMain.IncidQualityInterpretation != null ? _viewModelMain.IncidQualityInterpretation : "",
                        _viewModelMain.IncidQualityComments != null ? _viewModelMain.IncidQualityComments : "" },
                    _viewModelMain.HistoryColumns, incidCond);

                // Check if a history table was returned from updating
                // the GIS rows
                if (historyTable == null)
                    throw new Exception("Error updating GIS layer.");
                else if (historyTable.Rows.Count == 0)
                    throw new Exception("No GIS features were updated.");

                // TODO: Update length and area for each polygon (if possible)?
                // Likewise update the DB shadow copy of the GIS layer
                String updateStatement = String.Format("UPDATE {0} SET {1}={2}, {3}={4}, {5}={6}, {7}={8}, {9}={10} WHERE {11}",
                    _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_mm_polygons.TableName),
                    _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_mm_polygons.habprimaryColumn.ColumnName),
                    _viewModelMain.DataBase.QuoteValue(_viewModelMain.IncidPrimary),
                    _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_mm_polygons.habsecondColumn.ColumnName),
                    _viewModelMain.DataBase.QuoteValue(_viewModelMain.IncidSecondarySummary),
                    _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_mm_polygons.determqtyColumn.ColumnName),
                    _viewModelMain.DataBase.QuoteValue(_viewModelMain.IncidQualityDetermination),
                    _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_mm_polygons.interpqtyColumn.ColumnName),
                    _viewModelMain.DataBase.QuoteValue(_viewModelMain.IncidQualityInterpretation),
                    _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_mm_polygons.interpcomColumn.ColumnName),
                    _viewModelMain.DataBase.QuoteValue(_viewModelMain.IncidQualityComments),
                    _viewModelMain.DataBase.WhereClause(false, true, true, incidCond));

                if (_viewModelMain.DataBase.ExecuteNonQuery(updateStatement,
                    _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                    throw new Exception("Failed to update database copy of GIS layer.");

                // Save the history returned from GIS
                Dictionary<int, string> fixedValues = new Dictionary<int, string>();
                fixedValues.Add(_viewModelMain.HluDataset.history.incidColumn.Ordinal, _viewModelMain.Incid);
                ViewModelWindowMainHistory vmHist = new ViewModelWindowMainHistory(_viewModelMain);
                vmHist.HistoryWrite(fixedValues, historyTable, ViewModelWindowMain.Operations.AttributeUpdate, nowDtTm);

                // Commit the transation and accept the changes
                _viewModelMain.DataBase.CommitTransaction();
                _viewModelMain.HluDataset.AcceptChanges();
                _viewModelMain.Saved = true;

                // Recount the incid rows in the database
                _viewModelMain.IncidRowCount(true);

                // Reload the current row index
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
        /// Updates secondary habitat rows corresponding to current incid.
        /// </summary>
        private void UpdateSecondary()
        {
            if (_viewModelMain.IncidSecondaryHabitats == null)
                _viewModelMain.IncidSecondaryHabitats = new ObservableCollection<SecondaryHabitat>();

            // remove duplicate codes
            IEnumerable<SecondaryHabitat> currSecondaryRows = from s in _viewModelMain.IncidSecondaryHabitats
                                                        group s by new { s.secondary_group, s.secondary_habitat } into secs
                                                        select secs.First();

            List<HluDataSet.incid_secondaryRow> newRows = new List<HluDataSet.incid_secondaryRow>();
            List<HluDataSet.incid_secondaryRow> updateRows = new List<HluDataSet.incid_secondaryRow>();
            HluDataSet.incid_secondaryRow updateRow;

            foreach (SecondaryHabitat sh in currSecondaryRows)
            {
                if (sh.secondary_id == -1) // new secondary habitat environment
                {
                    sh.secondary_id = _viewModelMain.RecIDs.NextIncidSecondaryId;
                    sh.incid = _viewModelMain.Incid;
                    HluDataSet.incid_secondaryRow newRow = _viewModelMain.IncidSecondaryTable.Newincid_secondaryRow();
                    newRow.ItemArray = sh.ToItemArray();
                    newRows.Add(newRow);
                }
                // Get the new values for every updated secondary habitat row from the
                // secondary habitat data grid.
                else if ((updateRow = UpdateIncidSecondaryRow(sh)) != null)
                {
                    // If a row is returned from the data grid add it to the list
                    // of updated rows.
                    updateRows.Add(updateRow);
                }
            }

            // Delete any rows that haven't been marked as deleted but are
            // no longer in the current rows.
            List<HluDataSet.incid_secondaryRow> temp = _viewModelMain.IncidSecondaryRows.Where(r => r.RowState != DataRowState.Deleted &&
                currSecondaryRows.Count(g => g.secondary_id == r.secondary_id) == 0).ToList();

            _viewModelMain.IncidSecondaryRows.Where(r => r.RowState != DataRowState.Deleted &&
                currSecondaryRows.Count(g => g.secondary_id == r.secondary_id) == 0).ToList()
                .ForEach(delegate(HluDataSet.incid_secondaryRow row) { row.Delete(); });

            // Update the table to remove the deleted rows.
            if (_viewModelMain.HluTableAdapterManager.incid_secondaryTableAdapter.Update(
                _viewModelMain.IncidSecondaryRows.Where(r => r.RowState == DataRowState.Deleted).ToArray()) == -1)
                throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_secondary.TableName));

            // If there are any rows that have been updated.
            if (updateRows.Count > 0)
            {
                // Update the table to update the updated rows.
                if (_viewModelMain.HluTableAdapterManager.incid_secondaryTableAdapter.Update(updateRows.ToArray()) == -1)
                    throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_secondary.TableName));
            }

            // Insert the new rows into the table.
            foreach (HluDataSet.incid_secondaryRow r in newRows)
                _viewModelMain.HluTableAdapterManager.incid_secondaryTableAdapter.Insert(r);

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
        /// Writes the values from a SecondaryHabitat object bound to the secondaries data grids
        /// into the corresponding incid_secondary DataRow.
        /// </summary>
        /// <param name="sh">SecondaryHabitat object bound to data grid on form.</param>
        /// <returns>Updated incid_secondary row, or null if no corresponding row was found.</returns>
        private HluDataSet.incid_secondaryRow UpdateIncidSecondaryRow(SecondaryHabitat sh)
        {
            var q = _viewModelMain.IncidSecondaryRows.Where(r => r.RowState != DataRowState.Deleted && r.secondary_id == sh.secondary_id);
            if (q.Count() == 1)
            {
                if (!sh.IsValid()) return null;
                HluDataSet.incid_secondaryRow oldRow = q.ElementAt(0);
                object[] itemArray = sh.ToItemArray();
                for (int i = 0; i < itemArray.Length; i++)
                    oldRow[i] = itemArray[i];
                return oldRow;
            }
            return null;
        }

        /// <summary>
        /// Updates those columns of IncidCurrentRow in main view model that are not directly updated 
        /// by properties (to enable undo if update cancelled).
        /// </summary>
        /// <param name="viewModelMain">Reference to main window view model.</param>
        internal static void IncidCurrentRowDerivedValuesUpdate(ViewModelWindowMain viewModelMain)
        {
            // Update other incid vales
            viewModelMain.IncidCurrentRow.habitat_primary = viewModelMain.IncidPrimary;
            viewModelMain.IncidCurrentRow.habitat_secondaries = viewModelMain.IncidSecondarySummary;
            viewModelMain.IncidCurrentRow.habitat_version = viewModelMain.HabitatVersion;

            //TODO: Update - Needed (e.g. if clearing IHS values on update)?
            viewModelMain.IncidCurrentRow.ihs_habitat = viewModelMain.IncidIhsHabitat;
        }

        /// <summary>
        /// Updates the incid modified columns following a physical or
        /// logical split or merge.
        /// </summary>
        /// <param name="incid">The incid.</param>
        /// <param name="nowDtTm">The current date and time.</param>
        /// <exception cref="Exception">Failed to update incid table modified details.</exception>
        internal void UpdateIncidModifiedColumns(string incid, DateTime nowDtTm)
        {
            String updateStatement = String.Format("UPDATE {0} SET {1} = {2}, {3} = {4} WHERE {5} = {6}",
                _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid.TableName),
                _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid.last_modified_dateColumn.ColumnName),
                _viewModelMain.DataBase.QuoteValue(nowDtTm),
                _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid.last_modified_user_idColumn.ColumnName),
                _viewModelMain.DataBase.QuoteValue(_viewModelMain.UserID),
                _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid.incidColumn.ColumnName),
                _viewModelMain.DataBase.QuoteValue(incid));

            if (_viewModelMain.DataBase.ExecuteNonQuery(updateStatement,
                _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                throw new Exception("Failed to update incid table modified details.");
        }
    }
}
