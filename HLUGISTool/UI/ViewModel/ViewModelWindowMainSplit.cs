// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2013 Andy Foy
// Copyright © 2014 Sussex Biodiversity Record Centre
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
using System.Windows;
using HLU.Data;
using HLU.Data.Model;

namespace HLU.UI.ViewModel
{
    class ViewModelWindowMainSplit
    {
        ViewModelWindowMain _viewModelMain;

        public ViewModelWindowMainSplit(ViewModelWindowMain viewModelMain)
        {
            _viewModelMain = viewModelMain;
        }

        internal bool Split()
        {
            if ((_viewModelMain.GisSelection == null) || (_viewModelMain.GisSelection.Rows.Count == 0))
            {
                MessageBox.Show("Cannot split: nothing is selected on the map.", "HLU: Split",
                    MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return false;
            }
            if ((_viewModelMain.GisSelection.Rows.Count > 1) && (_viewModelMain.IncidsSelectedMapCount == 1) &&
                (_viewModelMain.ToidsSelectedMapCount == 1) && (_viewModelMain.FragsSelectedMapCount == 1))
            {
                // all features in selection share same incid, toid and toid_fragment_id
                return PerformPhysicalSplit();
            }
            else if ((_viewModelMain.GisSelection.Rows.Count > 0) && 
                ((_viewModelMain.ToidsSelectedMapCount > 1) || (_viewModelMain.FragsSelectedMapCount > 1) || 
                (_viewModelMain.FragsSelectedMapCount == 1)))
            {
                // all features in selection share same incid, but *not* toid and toid_fragment_id
                return PerformLogicalSplit();
            }
            else if (_viewModelMain.GisSelection.Rows.Count < 1)
            {
                MessageBox.Show("Cannot split: map selection set must contain at least one feature for a split.",
                    "HLU: Split", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return false;
            }
            else
            {
                MessageBox.Show("Cannot split: map selection set contains features belonging to more than one INCID.",
                    "HLU: Split", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return false;
            }
        }

        /// <summary>
        /// Physically split features on the GIS layer.
        /// </summary>
        private bool PerformPhysicalSplit()
        {
            bool success = true;
            _viewModelMain.DataBase.BeginTransaction(true, IsolationLevel.ReadCommitted);

            try
            {
                // find the last used toid_fragment_id for the selected toid
                string lastToidFragmentID = _viewModelMain.RecIDs.MaxToidFragmentId(_viewModelMain.ToidsSelectedMap.ElementAt(0));

                //---------------------------------------------------------------------
                // FIXED: KI110 (Physical split)
                // Skip all but one of the GIS select criteria as they are all the same in the case of a physical split anyway
                int skipCount = _viewModelMain.GisSelection.Rows.Count - 1;

                // get a filter from the GIS selection
                List<List<SqlFilterCondition>> featuresFilter = ViewModelWindowMainHelpers.GisSelectionToWhereClause(
                    _viewModelMain.GisSelection.AsEnumerable().Skip(skipCount).ToArray(), _viewModelMain.GisIDColumnOrdinals,
                    ViewModelWindowMain.IncidPageSize, _viewModelMain.HluDataset.incid_mm_polygons);
                //---------------------------------------------------------------------

                if (featuresFilter.Count != 1)
                    throw new Exception("Error finding features in database.");

                // Find the current toid_fragment_id for the selected toid (there should be only one)
                string currentToidFragmentID = _viewModelMain.FragsSelectedMap.ElementAt(0);

                // update records in GIS and collect new features resulting from split
                DataTable newFeatures = _viewModelMain.GISApplication.SplitFeature(currentToidFragmentID,
                    lastToidFragmentID, featuresFilter[0],
                    _viewModelMain.HluDataset.incid_mm_polygons.Columns.Cast<DataColumn>().Where(c =>
                        c.ColumnName != _viewModelMain.HluDataset.incid_mm_polygons.shape_lengthColumn.ColumnName &&
                        c.ColumnName != _viewModelMain.HluDataset.incid_mm_polygons.shape_areaColumn.ColumnName).ToArray());

                if ((newFeatures == null) || (newFeatures.Rows.Count < 2))
                    throw new Exception("Failed to update GIS layer.");

                ViewModelWindowMainHistory vmHist = new ViewModelWindowMainHistory(_viewModelMain);
                vmHist.HistoryRenameGeometryPropertyColumns(
                    _viewModelMain.HluDataset.incid_mm_polygons.shape_lengthColumn.ColumnName,
                    _viewModelMain.HluDataset.incid_mm_polygons.shape_areaColumn.ColumnName, ref newFeatures);

                // get a where clause for the original split feature
                List<List<SqlFilterCondition>> originalFeatureWhereClause = ViewModelWindowMainHelpers.GisSelectionToWhereClause(
                    _viewModelMain.GisSelection.AsEnumerable().Take(1).ToArray(), _viewModelMain.GisIDColumnOrdinals, 
                    ViewModelWindowMain.IncidPageSize, _viewModelMain.HluDataset.incid_mm_polygons);

                if (originalFeatureWhereClause.Count != 1)
                    throw new Exception("Error finding features in database.");

                // get the attributes of the split feature
                HluDataSet.incid_mm_polygonsDataTable updTable = new HluDataSet.incid_mm_polygonsDataTable();
                _viewModelMain.GetIncidMMPolygonRows(originalFeatureWhereClause, ref updTable);

                if ((updTable == null) || (updTable.Rows.Count != 1))
                    throw new Exception("Failed to fetch incid_mm_polygon rows.");

                // insert attributes of original split feature into history
                DataTable history = updTable.Copy();
                history.Columns[updTable.shape_lengthColumn.ColumnName].ColumnName = ViewModelWindowMain.HistoryGeometry1ColumnName;
                history.Columns[updTable.shape_areaColumn.ColumnName].ColumnName = ViewModelWindowMain.HistoryGeometry2ColumnName;
                string[] historyColNames = (new string[] { ViewModelWindowMain.HistoryGeometry1ColumnName, ViewModelWindowMain.HistoryGeometry2ColumnName }
                    .Concat(_viewModelMain.HistoryColumns.Select(c => c.ColumnName)).ToArray());
                DataColumn[] delCols = history.Columns.Cast<DataColumn>().Where(c => !historyColNames.Contains(c.ColumnName)).ToArray();
                foreach (DataColumn c in delCols)
                    history.Columns.Remove(c);

                vmHist.HistoryWrite(null, history, ViewModelWindowMain.Operations.PhysicalSplit);

                // update the original row
                if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format("UPDATE {0} SET {1} WHERE {2}",
                    _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_mm_polygons.TableName),
                    String.Join(",", newFeatures.Rows[0].ItemArray.Select((i, index) =>
                        new
                        {
                            ColumnName = newFeatures.Columns[index].ColumnName,
                            value = i
                        })
                    .Where(a => _viewModelMain.GisIDColumns.Count(c => c.ColumnName == a.ColumnName) == 0)
                    .Select(a => String.Format("{0} = {1}", _viewModelMain.DataBase.QuoteIdentifier(a.ColumnName),
                        _viewModelMain.DataBase.QuoteValue(a.value))).ToArray()),
                    _viewModelMain.DataBase.WhereClause(false, true, true, originalFeatureWhereClause[0])),
                    _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                    throw new Exception("Failed to update original row in database copy of GIS layer.");

                // build an insert statement for DB shadow copy of GIS layer
                string insertStatement = String.Format("INSERT INTO {0} ({1}) VALUES (", 
                    _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_mm_polygons.TableName),
                    String.Join(",", newFeatures.Columns.Cast<DataColumn>().Select(c => 
                    _viewModelMain.DataBase.QuoteIdentifier(c.ColumnName)).ToArray())) + "{0})";

                int toidFragID = Int32.Parse(lastToidFragmentID);
                string numFormat = String.Format("D{0}", updTable.toid_fragment_idColumn.MaxLength);

                // insert new features returned from GIS into DB shadow copy of GIS layer
                for (int i = 1; i < newFeatures.Rows.Count; i++)
                {
                    if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format(insertStatement, String.Join(",",
                        newFeatures.Rows[i].ItemArray.Select((item, index) =>
                            _viewModelMain.DataBase.QuoteValue(newFeatures.Columns[index].ColumnName ==
                            updTable.toid_fragment_idColumn.ColumnName ?
                            (toidFragID + i).ToString(numFormat) : item)).ToArray())),
                        _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                        throw new Exception("Failed to insert new rows into database copy of GIS layer.");
                }

                _viewModelMain.ViewModelUpdate.UpdateIncidModifiedColumns(_viewModelMain.IncidsSelectedMap.ElementAt(0));

                _viewModelMain.DataBase.CommitTransaction();
                _viewModelMain.HluDataset.AcceptChanges();
            }
            catch (Exception ex)
            {
                _viewModelMain.DataBase.RollbackTransaction();
                success = false;
                MessageBox.Show("Split operation failed. The error message returned was:\n\n" +
                    ex.Message, "HLU Split Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (success)
                {
                    // Re-count the incid records in the database.
                    _viewModelMain.IncidRowCount(true);

                    // Reset the incid and map selections but don't move
                    // to the first incid in the database.
                    _viewModelMain.ClearFilter(false);

                    // Synch with the GIS selection.
                    _viewModelMain.ReadMapSelection(true);
                }
            }
            return success;
        }

        private bool PerformLogicalSplit()
        {
            //if ((_viewModelMain.GisSelection.Rows.Count > 1) &&
            //    (_viewModelMain.HluDataset.incid_mm_polygons.Rows.Count == _viewModelMain.GisSelection.Rows.Count))
            //{
            //    MessageBox.Show(String.Format("Cannot split: all features in map " +
            //        "selection set correspond to INCID '{0}'.", _viewModelMain.Incid), "HLU: Split",
            //        MessageBoxButton.OK, MessageBoxImage.Exclamation);
            //    return false;
            //}
            //else
            //{
            // check if selected feature is the only one pertaining to its incid
            if (_viewModelMain.GisSelection.Rows.Count == 1)
            {
                int featCount = (int)_viewModelMain.DataBase.ExecuteScalar(String.Format(
                    "SELECT COUNT(*) FROM {0} WHERE {1} = {2}",
                    _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_mm_polygons.TableName),
                    _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid_mm_polygons.incidColumn.ColumnName),
                    _viewModelMain.DataBase.QuoteValue(_viewModelMain.Incid)),
                    _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text);
                if (featCount == 1)
                {
                    MessageBox.Show(String.Format("Cannot split: feature selected in map is the only" +
                        " feature corresponding to INCID {0}", _viewModelMain.Incid), "HLU: Split",
                        MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    return false;
                }
            }

            bool success = true;
            _viewModelMain.DataBase.BeginTransaction(true, IsolationLevel.ReadCommitted);

            try
            {
                //---------------------------------------------------------------------
                // FIX: Update the incid modified columns for the current incid
                // The incid modified columns (i.e. last modified user and date)
                // should be updated for the active incid not the 'highest' incid
                // which is what "_viewModelMain.CurrentIncid" represents.
                //
                //_viewModelMain.ViewModelUpdate.UpdateIncidModifiedColumns(_viewModelMain.CurrentIncid);
                _viewModelMain.ViewModelUpdate.UpdateIncidModifiedColumns(_viewModelMain.Incid);
                //---------------------------------------------------------------------

                // create new incid by cloning the current one
                string msg;
                if (!CloneCurrentIncid(false, out msg)) throw new Exception(msg);
                string newIncid = _viewModelMain.RecIDs.CurrentIncid;

                //---------------------------------------------------------------------
                // CHANGED: CR10 (Attribute updates for incid subsets)
                // Pass the old incid number together with the new incid number
                // so that only features belonging to the old incid are
                // updated.
                //
                // update GIS layer
                DataTable historyTable = _viewModelMain.GISApplication.SplitFeaturesLogically(_viewModelMain.Incid, newIncid,
                    _viewModelMain.HistoryColumns.Concat(new DataColumn[] { new DataColumn(
                            _viewModelMain.HluDataset.history.modified_toid_fragment_idColumn.ColumnName.Replace(
                            _viewModelMain.HluDataset.incid_mm_polygons.toid_fragment_idColumn.ColumnName, String.Empty) + 
                            GISApplication.GISApp.HistoryAdditionalFieldsDelimiter + 
                            _viewModelMain.HluDataset.incid_mm_polygons.toid_fragment_idColumn.ColumnName, 
                            _viewModelMain.HluDataset.history.modified_toid_fragment_idColumn.DataType)}).ToArray());
                //---------------------------------------------------------------------
                if (historyTable == null) throw new Exception("Failed to update GIS layer.");

                // update DB shadow copy of GIS layer
                HluDataSet.incid_mm_polygonsDataTable polygons = new HluDataSet.incid_mm_polygonsDataTable();
                _viewModelMain.GetIncidMMPolygonRows(ViewModelWindowMainHelpers.GisSelectionToWhereClause(
                    _viewModelMain.GisSelection.Select(), _viewModelMain.GisIDColumnOrdinals,
                    ViewModelWindowMain.IncidPageSize, polygons), ref polygons);

                historyTable.PrimaryKey = historyTable.Columns.Cast<DataColumn>()
                    .Where(c => _viewModelMain.GisIDColumnOrdinals.Contains(c.Ordinal)).ToArray();

                foreach (HluDataSet.incid_mm_polygonsRow r in polygons)
                {
                    // If the feature in GIS belongs to the current incid then update the
                    // toid_fragment_id and incid.
                    if (r.incid == _viewModelMain.Incid)
                    {
                        DataRow historyRow = historyTable.Rows.Find(r.ItemArray.Where((i, index) =>
                            _viewModelMain.GisIDColumnOrdinals.Contains(index)).ToArray());
                        r.toid_fragment_id = historyRow.Field<string>(
                            _viewModelMain.HluDataset.history.modified_toid_fragment_idColumn.ColumnName);
                        r.incid = newIncid;
                    }
                }
                if (_viewModelMain.HluTableAdapterManager.incid_mm_polygonsTableAdapter.Update(polygons) == -1)
                    throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_mm_polygons.TableName));

                //---------------------------------------------------------------------
                // FIX: Don't update the incid modified columns again
                // The incid modified columns (i.e. last modified user and date)
                // have already been update above for the current incid.
                //
                //_viewModelMain.ViewModelUpdate.UpdateIncidModifiedColumns(
                //    historyTable.Rows[0][_viewModelMain.HluDataset.history.incidColumn.ColumnName].ToString());
                //---------------------------------------------------------------------

                // write history
                Dictionary<int, string> fixedValues = new Dictionary<int, string>();
                fixedValues.Add(_viewModelMain.HluDataset.history.incidColumn.Ordinal, newIncid);
                historyTable.Columns[_viewModelMain.HluDataset.history.incidColumn.ColumnName].ColumnName =
                    _viewModelMain.HluDataset.history.modified_incidColumn.ColumnName;
                ViewModelWindowMainHistory vmHist = new ViewModelWindowMainHistory(_viewModelMain);
                vmHist.HistoryWrite(fixedValues, historyTable, ViewModelWindowMain.Operations.LogicalSplit);

                _viewModelMain.DataBase.CommitTransaction();
                _viewModelMain.HluDataset.AcceptChanges();

            }
            catch (Exception ex)
            {
                _viewModelMain.DataBase.RollbackTransaction();
                success = false;
                MessageBox.Show("Split operation failed. The error message returned was:\n\n" +
                    ex.Message, "HLU Split Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (success)
                {
                    // Re-count the incid records in the database.
                    _viewModelMain.IncidRowCount(true);

                    // Reset the incid and map selections but don't move
                    // to the first incid in the database.
                    _viewModelMain.ClearFilter(false);

                    // Synch with the GIS selection.
                    _viewModelMain.ReadMapSelection(true);
                }
            }
            //}
            return success;
        }

        private bool CloneCurrentIncid(bool startTransaction, out string errorMessage)
        {
            errorMessage = null;

            if (startTransaction && !_viewModelMain.DataBase.BeginTransaction(true, IsolationLevel.ReadCommitted))
                throw new Exception("Failed to start a database transaction.");

            try
            {
                // Create a local copy of the Incid row.
                HluDataSet.incidRow newIncidRow = _viewModelMain.IncidTable.NewincidRow();
                for (int i = 0; i < _viewModelMain.IncidTable.Columns.Count; i++)
                    if (!_viewModelMain.IncidCurrentRow.IsNull(i)) newIncidRow[i] = _viewModelMain.IncidCurrentRow[i];

                // Get the next number for the new incid.
                string newIncid = _viewModelMain.NextIncid;
                newIncidRow.incid = newIncid;

                //---------------------------------------------------------------------
                // CHANGED: CR10 (Attribute updates for incid subsets)
                // Get the current value of the IHS Habitat which isn't stored
                // in the current Incid row.
                newIncidRow.ihs_habitat = _viewModelMain.IncidIhsHabitat;

                // Discard any changes to the Incid table once a copy has been
                // made.
                _viewModelMain.IncidTable.RejectChanges();

                //---------------------------------------------------------------------

                //---------------------------------------------------------------------
                // FIX: Don't update the ihs version when splitting features
                // Previously the ihs version was set to the latest version for
                // new (cloned) incids which meant that a feature split into 2
                // parts could result in one with the current 'old' ihs version
                // and one with the latest ihs version.
                // The ihs version should only be updated when attribute updates
                // are applied to an incid (as the ihs attributes have then been
                // validated by the tool and the user) and not as a result of a
                // logical split.
                //
                //newIncidRow.ihs_version = _viewModelMain.RecIDs.IhsVersion;
                //---------------------------------------------------------------------

                // Update the created and last modified date and user fields.
                newIncidRow.created_date = DateTime.Now;
                newIncidRow.created_user_id = _viewModelMain.UserID;
                newIncidRow.last_modified_date = newIncidRow.created_date;
                newIncidRow.last_modified_user_id = newIncidRow.created_user_id;

                // Add the local copy of the row to the table.
                _viewModelMain.IncidTable.AddincidRow(newIncidRow);

                // Update the table with the new row.
                if (_viewModelMain.HluTableAdapterManager.incidTableAdapter.Update(_viewModelMain.HluDataset.incid) == -1)
                    throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid.TableName));

                if ((_viewModelMain.IncidIhsMatrixRows != null) && (_viewModelMain.IncidIhsMatrixRows.Length > 0))
                {
                    //---------------------------------------------------------------------
                    // CHANGED: CR10 (Attribute updates for incid subsets)
                    // Copy the values of any IncidIhsMatrix rows rather than
                    // copying the rows themselves so that any pending changes
                    // to the rows can be discarded afterwards.
                    //
                    //HluDataSet.incid_ihs_matrixRow[] ihsMatrixRows =
                    //    _viewModelMain.IncidIhsMatrixRows.Where(r => r != null).ToArray();

                    // Create a local copy of the IncidIhsMatrix rows.
                    List<HluDataSet.incid_ihs_matrixRow> ihsMatrixRows = new List<HluDataSet.incid_ihs_matrixRow>();

                    // Copy the column values for each row in the IncidIhsMatrix table.
                    foreach (HluDataSet.incid_ihs_matrixRow row in _viewModelMain.IncidIhsMatrixRows)
                    {
                        if (row != null)
                        {
                            HluDataSet.incid_ihs_matrixRow newIncidIhsMatrixRow = _viewModelMain.IncidIhsMatrixTable.Newincid_ihs_matrixRow();
                            for (int i = 0; i < _viewModelMain.IncidIhsMatrixTable.Columns.Count; i++)
                                if (!row.IsNull(i)) newIncidIhsMatrixRow[i] = row[i];

                            ihsMatrixRows.Add(newIncidIhsMatrixRow);
                        }
                    }

                    // Discard any changes to the IncidIhsMatrix table once a copy has been
                    // made.
                    _viewModelMain.IncidIhsMatrixTable.RejectChanges();

                    // Remove any rows added by the edit that have been discarded but
                    // are still in the rows array.
                    for (int i = 0; i < _viewModelMain.IncidIhsMatrixRows.Count(); i++)
                    {
                        if ((_viewModelMain.IncidIhsMatrixRows[i] != null) && (_viewModelMain.IncidIhsMatrixRows[i].incidRow == null))
                        {
                            if (_viewModelMain.IncidIhsMatrixRows[i].RowState != DataRowState.Detached)
                                _viewModelMain.IncidIhsMatrixRows[i].Delete();
                            _viewModelMain.IncidIhsMatrixRows[i] = null;
                        }
                    }
                    //---------------------------------------------------------------------

                    // Clone the temporary rows, replacing the original incid with the
                    // new incid.
                    foreach (HluDataSet.incid_ihs_matrixRow row in ihsMatrixRows)
                        _viewModelMain.HluDataset.incid_ihs_matrix.Addincid_ihs_matrixRow(
                            CloneRow<HluDataSet.incid_ihs_matrixRow, HluDataSet.incid_ihs_matrixDataTable>(
                            row, _viewModelMain.HluDataset.incid_ihs_matrix.incidColumn.Ordinal, newIncid));

                    // Update the table with the new rows.
                    if (_viewModelMain.HluTableAdapterManager.incid_ihs_matrixTableAdapter.Update(_viewModelMain.HluDataset.incid_ihs_matrix) == -1)
                        throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_ihs_matrix.TableName));
                }

                if ((_viewModelMain.IncidIhsFormationRows != null) && (_viewModelMain.IncidIhsFormationRows.Length > 0))
                {
                    //---------------------------------------------------------------------
                    // CHANGED: CR10 (Attribute updates for incid subsets)
                    // Copy the values of any IncidIhsFormation rows rather than
                    // copying the rows themselves so that any pending changes
                    // to the rows can be discarded afterwards.
                    //
                    //HluDataSet.incid_ihs_formationRow[] ihsFormationRows =
                    //    _viewModelMain.IncidIhsFormationRows.Where(r => r != null).ToArray();

                    // Create a local copy of the IncidIhsFormation rows.
                    List<HluDataSet.incid_ihs_formationRow> ihsFormationRows = new List<HluDataSet.incid_ihs_formationRow>();

                    // Copy the column values for each row in the IncidIhsFormation table.
                    foreach (HluDataSet.incid_ihs_formationRow row in _viewModelMain.IncidIhsFormationRows)
                    {
                        if (row != null)
                        {
                            HluDataSet.incid_ihs_formationRow newIncidIhsFormationRow = _viewModelMain.IncidIhsFormationTable.Newincid_ihs_formationRow();
                            for (int i = 0; i < _viewModelMain.IncidIhsFormationTable.Columns.Count; i++)
                                if (!row.IsNull(i)) newIncidIhsFormationRow[i] = row[i];

                            ihsFormationRows.Add(newIncidIhsFormationRow);
                        }
                    }

                    // Discard any changes to the IncidIhsFormation table once a copy has been
                    // made.
                    _viewModelMain.IncidIhsFormationTable.RejectChanges();

                    // Remove any rows added by the edit that have been discarded but
                    // are still in the rows array.
                    for (int i = 0; i < _viewModelMain.IncidIhsFormationRows.Count(); i++)
                    {
                        if ((_viewModelMain.IncidIhsFormationRows[i] != null) && (_viewModelMain.IncidIhsFormationRows[i].incidRow == null))
                        {
                            if (_viewModelMain.IncidIhsFormationRows[i].RowState != DataRowState.Detached)
                                _viewModelMain.IncidIhsFormationRows[i].Delete();
                            _viewModelMain.IncidIhsFormationRows[i] = null;
                        }
                    }
                    //---------------------------------------------------------------------

                    // Clone the temporary rows, replacing the original incid with the
                    // new incid.
                    foreach (HluDataSet.incid_ihs_formationRow row in ihsFormationRows)
                        _viewModelMain.HluDataset.incid_ihs_formation.Addincid_ihs_formationRow(
                            CloneRow<HluDataSet.incid_ihs_formationRow, HluDataSet.incid_ihs_formationDataTable>(
                            row, _viewModelMain.HluDataset.incid_ihs_formation.incidColumn.Ordinal, newIncid));

                    // Update the table with the new rows.
                    if (_viewModelMain.HluTableAdapterManager.incid_ihs_formationTableAdapter.Update(_viewModelMain.HluDataset.incid_ihs_formation) == -1)
                            throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_ihs_formation.TableName));
                }

                if ((_viewModelMain.IncidIhsManagementRows != null) && (_viewModelMain.IncidIhsManagementRows.Length > 0))
                {
                    //---------------------------------------------------------------------
                    // CHANGED: CR10 (Attribute updates for incid subsets)
                    // Copy the values of any IncidIhsManagement rows rather than
                    // copying the rows themselves so that any pending changes
                    // to the rows can be discarded afterwards.
                    //
                    //HluDataSet.incid_ihs_managementRow[] ihsManagementRows =
                    //    _viewModelMain.IncidIhsManagementRows.Where(r => r != null).ToArray();

                    // Create a local copy of the IncidIhsManagement rows.
                    List<HluDataSet.incid_ihs_managementRow> ihsManagementRows = new List<HluDataSet.incid_ihs_managementRow>();

                    // Copy the column values for each row in the IncidIhsManagement table.
                    foreach (HluDataSet.incid_ihs_managementRow row in _viewModelMain.IncidIhsManagementRows)
                    {
                        if (row != null)
                        {
                            HluDataSet.incid_ihs_managementRow newIncidIhsManagementRow = _viewModelMain.IncidIhsManagementTable.Newincid_ihs_managementRow();
                            for (int i = 0; i < _viewModelMain.IncidIhsManagementTable.Columns.Count; i++)
                                if (!row.IsNull(i)) newIncidIhsManagementRow[i] = row[i];

                            ihsManagementRows.Add(newIncidIhsManagementRow);
                        }
                    }

                    // Discard any changes to the IncidIhsManagement table once a copy has been
                    // made.
                    _viewModelMain.IncidIhsManagementTable.RejectChanges();

                    // Remove any rows added by the edit that have been discarded but
                    // are still in the rows array.
                    for (int i = 0; i < _viewModelMain.IncidIhsManagementRows.Count(); i++)
                    {
                        if ((_viewModelMain.IncidIhsManagementRows[i] != null) && (_viewModelMain.IncidIhsManagementRows[i].incidRow == null))
                        {
                            if (_viewModelMain.IncidIhsManagementRows[i].RowState != DataRowState.Detached)
                                _viewModelMain.IncidIhsManagementRows[i].Delete();
                            _viewModelMain.IncidIhsManagementRows[i] = null;
                        }
                    }
                    //---------------------------------------------------------------------

                    // Clone the temporary rows, replacing the original incid with the
                    // new incid.
                    foreach (HluDataSet.incid_ihs_managementRow row in ihsManagementRows)
                        _viewModelMain.HluDataset.incid_ihs_management.Addincid_ihs_managementRow(
                            CloneRow<HluDataSet.incid_ihs_managementRow, HluDataSet.incid_ihs_managementDataTable>(
                            row, _viewModelMain.HluDataset.incid_ihs_management.incidColumn.Ordinal, newIncid));

                    // Update the table with the new rows.
                    if (_viewModelMain.HluTableAdapterManager.incid_ihs_managementTableAdapter.Update(_viewModelMain.HluDataset.incid_ihs_management) == -1)
                        throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_ihs_management.TableName));
                }

                if ((_viewModelMain.IncidIhsComplexRows != null) && (_viewModelMain.IncidIhsComplexRows.Length > 0))
                {
                    //---------------------------------------------------------------------
                    // CHANGED: CR10 (Attribute updates for incid subsets)
                    // Copy the values of any IncidIhsComplex rows rather than
                    // copying the rows themselves so that any pending changes
                    // to the rows can be discarded afterwards.
                    //
                    //HluDataSet.incid_ihs_complexRow[] ihsComplexRows =
                    //    _viewModelMain.IncidIhsComplexRows.Where(r => r != null).ToArray();

                    // Create a local copy of the IncidIhsComplex rows.
                    List<HluDataSet.incid_ihs_complexRow> ihsComplexRows = new List<HluDataSet.incid_ihs_complexRow>();

                    // Copy the column values for each row in the IncidIhsComplex table.
                    foreach (HluDataSet.incid_ihs_complexRow row in _viewModelMain.IncidIhsComplexRows)
                    {
                        if (row != null)
                        {
                            HluDataSet.incid_ihs_complexRow newIncidIhsComplexRow = _viewModelMain.IncidIhsComplexTable.Newincid_ihs_complexRow();
                            for (int i = 0; i < _viewModelMain.IncidIhsComplexTable.Columns.Count; i++)
                                if (!row.IsNull(i)) newIncidIhsComplexRow[i] = row[i];

                            ihsComplexRows.Add(newIncidIhsComplexRow);
                        }
                    }

                    // Discard any changes to the IncidIhsComplex table once a copy has been
                    // made.
                    _viewModelMain.IncidIhsComplexTable.RejectChanges();

                    // Remove any rows added by the edit that have been discarded but
                    // are still in the rows array.
                    for (int i = 0; i < _viewModelMain.IncidIhsComplexRows.Count(); i++)
                    {
                        if ((_viewModelMain.IncidIhsComplexRows[i] != null) && (_viewModelMain.IncidIhsComplexRows[i].incidRow == null))
                        {
                            if (_viewModelMain.IncidIhsComplexRows[i].RowState != DataRowState.Detached)
                                _viewModelMain.IncidIhsComplexRows[i].Delete();
                            _viewModelMain.IncidIhsComplexRows[i] = null;
                        }
                    }
                    //---------------------------------------------------------------------

                    // Clone the temporary rows, replacing the original incid with the
                    // new incid.
                    foreach (HluDataSet.incid_ihs_complexRow row in ihsComplexRows)
                        _viewModelMain.HluDataset.incid_ihs_complex.Addincid_ihs_complexRow(
                            CloneRow<HluDataSet.incid_ihs_complexRow, HluDataSet.incid_ihs_complexDataTable>(
                            row, _viewModelMain.HluDataset.incid_ihs_complex.incidColumn.Ordinal, newIncid));

                    // Update the table with the new rows.
                    if (_viewModelMain.HluTableAdapterManager.incid_ihs_complexTableAdapter.Update(_viewModelMain.HluDataset.incid_ihs_complex) == -1)
                        throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_ihs_complex.TableName));
                }

                //---------------------------------------------------------------------
                // CHANGED: CR10 (Attribute updates for incid subsets)
                // Copy the values of any IncidBap rows rather than
                // copying the rows themselves so that any pending changes
                // to the rows can be discarded afterwards.
                //
                //HluDataSet.incid_bapRow[] bapRows = _viewModelMain.IncidBapRows.Where(r => r != null).ToArray();

                // Create a local copy of the IncidBap rows.
                List<HluDataSet.incid_bapRow> bapRows = new List<HluDataSet.incid_bapRow>();

                IEnumerable<BapEnvironment> beAuto = null;
                IEnumerable<BapEnvironment> beUser = null;
                //beAuto = new BapEnvironment[0].AsEnumerable();
                //beUser = new BapEnvironment[0].AsEnumerable();

                // Get a list of the auto BAP rows.
                if (_viewModelMain.IncidBapRowsAuto != null)
                {
                    beAuto = from b in _viewModelMain.IncidBapRowsAuto
                                                            group b by b.bap_habitat into habs
                                                            select habs.First();
                }

                // Get a list of the user BAP rows, removing any duplicate
                // codes already in the auto list as it goes.
                if (_viewModelMain.IncidBapHabitatsUser != null)
                {
                    beUser = from b in _viewModelMain.IncidBapHabitatsUser
                                                            where beAuto.Count(a => a.bap_habitat == b.bap_habitat) == 0
                                                            group b by b.bap_habitat into habs
                                                            select habs.First();
                }

                // Concatenate the two BAP lists together.
                var currentBapRows = beAuto.Concat(beUser);

                // Iterate through the concatenated BAP list, adding each
                // one to the local BAP table.
                foreach (BapEnvironment be in currentBapRows)
                {
                    HluDataSet.incid_bapRow newRow;
                    // If the row is new get a new bap_id and then add it to
                    // the local copy of rows.
                    if (be.bap_id == -1)
                    {
                        be.bap_id = _viewModelMain.RecIDs.NextIncidBapId;
                        be.incid = _viewModelMain.Incid;
                        newRow = _viewModelMain.IncidBapTable.Newincid_bapRow();
                        newRow.ItemArray = be.ToItemArray();
                        bapRows.Add(newRow);
                    }
                    // If the row is updated get the new values from the bap
                    // data grid and then add it to the local copy of rows.
                    else if ((newRow = UpdateIncidBapRow(be)) != null)
                    {
                        bapRows.Add(newRow);
                    }
                }

                // Restore the IHS Habitat back to it's original value. This
                // will retrieve the auto and user bap rows from the database
                // again so that they will no longer be flagged as changed.
                _viewModelMain.IncidBapRowsUser = null;
                _viewModelMain.IncidIhsHabitat = _viewModelMain.IncidCurrentRow[_viewModelMain.IncidTable.ihs_habitatColumn.ColumnName].ToString();

                // If there are any local rows ...
                if ((bapRows != null) && (bapRows.Count > 0))
                {
                    // Clone the temporary rows, replacing the original incid with the
                    // new incid.
                    foreach (HluDataSet.incid_bapRow row in bapRows)
                    {
                        _viewModelMain.HluDataset.incid_bap.Addincid_bapRow(
                            CloneRow<HluDataSet.incid_bapRow, HluDataSet.incid_bapDataTable>(
                            row, _viewModelMain.HluDataset.incid_bap.incidColumn.Ordinal, newIncid));
                    }

                    // Update the table with the new rows.
                    if (_viewModelMain.HluTableAdapterManager.incid_bapTableAdapter.Update(_viewModelMain.HluDataset.incid_bap) == -1)
                        throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_bap.TableName));
                }
                //---------------------------------------------------------------------

                if ((_viewModelMain.IncidSourcesRows != null) && (_viewModelMain.IncidSourcesRows.Length > 0))
                {
                    //---------------------------------------------------------------------
                    // CHANGED: CR10 (Attribute updates for incid subsets)
                    // Copy the values of any IncidSources rows rather than
                    // copying the rows themselves so that any pending changes
                    // to the rows can be discarded afterwards.
                    //
                    //HluDataSet.incid_sourcesRow[] sourcesRows = _viewModelMain.IncidSourcesRows.Where(r => r != null).ToArray();

                    // Create a local copy of the IncidSources rows.
                    List<HluDataSet.incid_sourcesRow> sourcesRows = new List<HluDataSet.incid_sourcesRow>();

                    // Copy the column values for each row in the IncidSources table.
                    foreach (HluDataSet.incid_sourcesRow row in _viewModelMain.IncidSourcesRows)
                    {
                        if (row != null)
                        {
                            HluDataSet.incid_sourcesRow newIncidSourcesRow = _viewModelMain.IncidSourcesTable.Newincid_sourcesRow();
                            for (int i = 0; i < _viewModelMain.IncidSourcesTable.Columns.Count; i++)
                                if (!row.IsNull(i)) newIncidSourcesRow[i] = row[i];

                            sourcesRows.Add(newIncidSourcesRow);
                        }
                    }

                    // Discard any changes to the IncidSources table once a copy has been
                    // made.
                    _viewModelMain.IncidSourcesTable.RejectChanges();

                    // Remove any rows added by the edit that have been discarded but
                    // are still in the rows array.
                    for (int i = 0; i < _viewModelMain.IncidSourcesRows.Count(); i++)
                    {
                        if ((_viewModelMain.IncidSourcesRows[i] != null) && (_viewModelMain.IncidSourcesRows[i].incidRow == null))
                        {
                            if (_viewModelMain.IncidSourcesRows[i].RowState != DataRowState.Detached)
                                _viewModelMain.IncidSourcesRows[i].Delete();
                            _viewModelMain.IncidSourcesRows[i] = null;
                        }
                    }
                    //---------------------------------------------------------------------

                    // Clone the temporary rows, replacing the original incid with the
                    // new incid.
                    foreach (HluDataSet.incid_sourcesRow row in sourcesRows)
                        _viewModelMain.HluDataset.incid_sources.Addincid_sourcesRow(
                            CloneRow<HluDataSet.incid_sourcesRow, HluDataSet.incid_sourcesDataTable>(
                            row, _viewModelMain.HluDataset.incid_sources.incidColumn.Ordinal, newIncid));

                    // Update the table with the new rows.
                    if (_viewModelMain.HluTableAdapterManager.incid_sourcesTableAdapter.Update(_viewModelMain.HluDataset.incid_sources) == -1)
                        throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_sources.TableName));
                }

                if (startTransaction) _viewModelMain.DataBase.CommitTransaction();

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                if (startTransaction)
                {
                    _viewModelMain.DataBase.RollbackTransaction();
                    MessageBox.Show("Changes could not be written to the database. The error message was:\n\n" + ex.Message,
                        "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                return false;
            }
        }

        private R CloneRow<R, T>(R row, int incidColumnOrdinal, string incid)
            where R : DataRow
            where T : DataTable
        {
            if (row == null) return null;
            T table = (T)((R)row).Table;
            R newRow = (R)table.NewRow();
            if ((table.PrimaryKey == null) || (table.PrimaryKey.Length != 1)) return null;
            int pkOrdinal = table.PrimaryKey[0].Ordinal;
            if (table.Columns[pkOrdinal].DataType != typeof(Int32)) return null;
            for (int i = 0; i < table.Columns.Count; i++)
            {
                if (i == incidColumnOrdinal)
                    newRow[i] = incid;
                else if (i == pkOrdinal)
                    newRow[i] = _viewModelMain.RecIDs.NextID<T>(-1, table, pkOrdinal);
                else if (!row.IsNull(i))
                    newRow[i] = row[i];
            }
            return newRow;
        }

        //---------------------------------------------------------------------
        // CHANGED: CR10 (Attribute updates for incid subsets)
        // Used when cloning an incid to retrieve the current values for
        // a given BAP row.
        //
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
        //---------------------------------------------------------------------
    }
}
