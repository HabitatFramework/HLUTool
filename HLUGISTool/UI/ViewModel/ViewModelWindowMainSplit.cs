using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;
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
            else if ((_viewModelMain.GisSelection.Rows.Count > 0) && (_viewModelMain.IncidsSelectedMapCount == 1) &&
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
                // find next available toid_fragment_id for selected toid
                string lastToidFragmentID = _viewModelMain.RecIDs.MaxToidFragmentId(_viewModelMain.ToidsSelectedMap.ElementAt(0));

                // get a filter from the GIS selection
                List<List<SqlFilterCondition>> featuresFilter = ViewModelWindowMainHelpers.GisSelectionToWhereClause(
                    _viewModelMain.GisSelection.AsEnumerable().Skip(1).ToArray(), _viewModelMain.GisIDColumnOrdinals,
                    ViewModelWindowMain.IncidPageSize, _viewModelMain.HluDataset.incid_mm_polygons);

                if (featuresFilter.Count != 1)
                    throw new Exception("Error finding features in database.");

                // update records in GIS and collect new features resulting from split
                DataTable newFeatures = _viewModelMain.GISApplication.SplitFeature(lastToidFragmentID, featuresFilter[0],
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
                    _viewModelMain.IncidRowCount(true);
                    _viewModelMain.ClearFilter();
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
                _viewModelMain.ViewModelUpdate.UpdateIncidModifiedColumns(_viewModelMain.CurrentIncid);

                // create new incid by cloning the current one
                string msg;
                if (!CloneCurrentIncid(false, out msg)) throw new Exception(msg);
                string newIncid = _viewModelMain.RecIDs.CurrentIncid;

                // update GIS layer
                DataTable historyTable = _viewModelMain.GISApplication.SplitFeaturesLogically(newIncid,
                    _viewModelMain.HistoryColumns.Concat(new DataColumn[] { new DataColumn(
                            _viewModelMain.HluDataset.history.modified_toid_fragment_idColumn.ColumnName.Replace(
                            _viewModelMain.HluDataset.incid_mm_polygons.toid_fragment_idColumn.ColumnName, String.Empty) + 
                            GISApplication.GISApp.HistoryAdditionalFieldsDelimiter + 
                            _viewModelMain.HluDataset.incid_mm_polygons.toid_fragment_idColumn.ColumnName, 
                            _viewModelMain.HluDataset.history.modified_toid_fragment_idColumn.DataType)}).ToArray());
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
                    DataRow historyRow = historyTable.Rows.Find(r.ItemArray.Where((i, index) =>
                        _viewModelMain.GisIDColumnOrdinals.Contains(index)).ToArray());
                    r.toid_fragment_id = historyRow.Field<string>(
                        _viewModelMain.HluDataset.history.modified_toid_fragment_idColumn.ColumnName);
                    r.incid = newIncid;
                }
                if (_viewModelMain.HluTableAdapterManager.incid_mm_polygonsTableAdapter.Update(polygons) == -1)
                    throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_mm_polygons.TableName));

                _viewModelMain.ViewModelUpdate.UpdateIncidModifiedColumns(
                    historyTable.Rows[0][_viewModelMain.HluDataset.history.incidColumn.ColumnName].ToString());

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
                    _viewModelMain.ClearFilter();

                    // re-count records
                    _viewModelMain.IncidRowCount(true);

                    // synch with GIS selection
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
                HluDataSet.incidRow newIncidRow = _viewModelMain.IncidTable.NewincidRow();
                for (int i = 0; i < _viewModelMain.IncidTable.Columns.Count; i++)
                    if (!_viewModelMain.IncidCurrentRow.IsNull(i)) newIncidRow[i] = _viewModelMain.IncidCurrentRow[i];
                string newIncid = _viewModelMain.NextIncid;
                newIncidRow.incid = newIncid;
                newIncidRow.ihs_version = _viewModelMain.RecIDs.IhsVersion;
                newIncidRow.created_date = DateTime.Today;
                newIncidRow.created_user_id = _viewModelMain.UserID;
                newIncidRow.last_modified_date = newIncidRow.created_date;
                newIncidRow.last_modified_user_id = newIncidRow.created_user_id;

                _viewModelMain.IncidTable.AddincidRow(newIncidRow);
                if (_viewModelMain.HluTableAdapterManager.incidTableAdapter.Update(_viewModelMain.HluDataset.incid) == -1)
                    throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid.TableName));

                if ((_viewModelMain.IncidIhsMatrixRows != null) && (_viewModelMain.IncidIhsMatrixRows.Length > 0))
                {
                    HluDataSet.incid_ihs_matrixRow[] ihsMatrixRows =
                        _viewModelMain.IncidIhsMatrixRows.Where(r => r != null).ToArray();
                    foreach (HluDataSet.incid_ihs_matrixRow row in ihsMatrixRows)
                        _viewModelMain.HluDataset.incid_ihs_matrix.Addincid_ihs_matrixRow(
                            CloneRow<HluDataSet.incid_ihs_matrixRow, HluDataSet.incid_ihs_matrixDataTable>(
                            row, _viewModelMain.HluDataset.incid_ihs_matrix.incidColumn.Ordinal, newIncid));
                    if (_viewModelMain.HluTableAdapterManager.incid_ihs_matrixTableAdapter.Update(_viewModelMain.HluDataset.incid_ihs_matrix) == -1)
                        throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_ihs_matrix.TableName));
                }

                if ((_viewModelMain.IncidIhsFormationRows != null) && (_viewModelMain.IncidIhsFormationRows.Length > 0))
                {
                    HluDataSet.incid_ihs_formationRow[] ihsFormationRows =
                        _viewModelMain.IncidIhsFormationRows.Where(r => r != null).ToArray();
                    foreach (HluDataSet.incid_ihs_formationRow row in ihsFormationRows)
                        _viewModelMain.HluDataset.incid_ihs_formation.Addincid_ihs_formationRow(
                            CloneRow<HluDataSet.incid_ihs_formationRow, HluDataSet.incid_ihs_formationDataTable>(
                            row, _viewModelMain.HluDataset.incid_ihs_formation.incidColumn.Ordinal, newIncid));
                    if (_viewModelMain.HluTableAdapterManager.incid_ihs_formationTableAdapter.Update(_viewModelMain.HluDataset.incid_ihs_formation) == -1)
                            throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_ihs_formation.TableName));
                }

                if ((_viewModelMain.IncidIhsManagementRows != null) && (_viewModelMain.IncidIhsManagementRows.Length > 0))
                {
                    HluDataSet.incid_ihs_managementRow[] ihsManagementRows =
                        _viewModelMain.IncidIhsManagementRows.Where(r => r != null).ToArray();
                    foreach (HluDataSet.incid_ihs_managementRow row in ihsManagementRows)
                        _viewModelMain.HluDataset.incid_ihs_management.Addincid_ihs_managementRow(
                            CloneRow<HluDataSet.incid_ihs_managementRow, HluDataSet.incid_ihs_managementDataTable>(
                            row, _viewModelMain.HluDataset.incid_ihs_management.incidColumn.Ordinal, newIncid));
                    if (_viewModelMain.HluTableAdapterManager.incid_ihs_managementTableAdapter.Update(_viewModelMain.HluDataset.incid_ihs_management) == -1)
                        throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_ihs_management.TableName));
                }

                if ((_viewModelMain.IncidIhsComplexRows != null) && (_viewModelMain.IncidIhsComplexRows.Length > 0))
                {
                    HluDataSet.incid_ihs_complexRow[] ihsComplexRows =
                        _viewModelMain.IncidIhsComplexRows.Where(r => r != null).ToArray();
                    foreach (HluDataSet.incid_ihs_complexRow row in ihsComplexRows)
                        _viewModelMain.HluDataset.incid_ihs_complex.Addincid_ihs_complexRow(
                            CloneRow<HluDataSet.incid_ihs_complexRow, HluDataSet.incid_ihs_complexDataTable>(
                            row, _viewModelMain.HluDataset.incid_ihs_complex.incidColumn.Ordinal, newIncid));
                    if (_viewModelMain.HluTableAdapterManager.incid_ihs_complexTableAdapter.Update(_viewModelMain.HluDataset.incid_ihs_complex) == -1)
                        throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_ihs_complex.TableName));
                }

                if ((_viewModelMain.IncidBapRows != null) && (_viewModelMain.IncidBapRows.Length > 0))
                {
                    HluDataSet.incid_bapRow[] bapRows = _viewModelMain.IncidBapRows.Where(r => r != null).ToArray();
                    foreach (HluDataSet.incid_bapRow row in bapRows)
                    {
                        _viewModelMain.HluDataset.incid_bap.Addincid_bapRow(
                            CloneRow<HluDataSet.incid_bapRow, HluDataSet.incid_bapDataTable>(
                            row, _viewModelMain.HluDataset.incid_bap.incidColumn.Ordinal, newIncid));
                    }
                    if (_viewModelMain.HluTableAdapterManager.incid_bapTableAdapter.Update(_viewModelMain.HluDataset.incid_bap) == -1)
                        throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_bap.TableName));
                }

                if ((_viewModelMain.IncidSourcesRows != null) && (_viewModelMain.IncidSourcesRows.Length > 0))
                {
                    HluDataSet.incid_sourcesRow[] sourcesRows = _viewModelMain.IncidSourcesRows.Where(r => r != null).ToArray();
                    foreach (HluDataSet.incid_sourcesRow row in sourcesRows)
                        _viewModelMain.HluDataset.incid_sources.Addincid_sourcesRow(
                            CloneRow<HluDataSet.incid_sourcesRow, HluDataSet.incid_sourcesDataTable>(
                            row, _viewModelMain.HluDataset.incid_sources.incidColumn.Ordinal, newIncid));
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
    }
}
