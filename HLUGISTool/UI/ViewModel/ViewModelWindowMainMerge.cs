// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
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
using System.Data;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using HLU.Data.Model;
using HLU.UI.View;

namespace HLU.UI.ViewModel
{
    class ViewModelWindowMainMerge
    {
        private ViewModelWindowMain _viewModelMain;
        private WindowMergeFeatures _mergeFeaturesWindow;
        private ViewModelMergeFeatures<HluDataSet.incidDataTable, HluDataSet.incidRow> _mergeFeaturesViewModelLogical;
        private ViewModelMergeFeatures<HluDataSet.incid_mm_polygonsDataTable, HluDataSet.incid_mm_polygonsRow> 
            _mergeFeaturesViewModelPhysical;
        private int _mergeResultFeatureIndex;

        public ViewModelWindowMainMerge(ViewModelWindowMain viewModelMain)
        {
            _viewModelMain = viewModelMain;
        }

        /// <summary>
        /// There must be at least two selected features that either share the same toid but not the same incid,
        /// or they do not share the same incid.
        /// </summary>
        internal void LogicalMerge()
        {
            if ((_viewModelMain.GisSelection == null) || (_viewModelMain.GisSelection.Rows.Count == 0))
            {
                MessageBox.Show("Cannot logically merge: nothing is selected on the map.", "HLU: Logical Merge",
                    MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            else if (_viewModelMain.GisSelection.Rows.Count <= 1)
            {
                MessageBox.Show("Cannot merge: map selection must contain more than one feature for a merge.",
                    "HLU: Logical Merge", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            else if (_viewModelMain.IncidsSelectedMapCount > 1)
            {
                // selected features do not share same incid
                PerformLogicalMerge(false);
            }
            else if ((_viewModelMain.ToidsSelectedMapCount == 1) && (_viewModelMain.IncidsSelectedMapCount > 1))
            {
                // selected features share same toid but not incid
                PerformLogicalMerge(true);
            }
        }

        /// <summary>
        /// There must be at least two selected features that share the same incid and toid.
        /// </summary>
        internal void PhysicalMerge()
        {
            if ((_viewModelMain.GisSelection == null) || (_viewModelMain.GisSelection.Rows.Count == 0))
            {
                MessageBox.Show("Cannot physically merge: nothing is selected on the map.", "HLU: Physical Merge",
                    MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            else if (_viewModelMain.GisSelection.Rows.Count <= 1)
            {
                MessageBox.Show("Cannot physically merge: map selection must contain more than one feature for a merge.",
                    "HLU: Physical Merge", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            else if ((_viewModelMain.IncidsSelectedMapCount == 1) && (_viewModelMain.ToidsSelectedMapCount == 1))
            {
                // selected features share same incid and toid
                PerformPhysicalMerge();
            }
        }

        private void PerformLogicalMerge(bool physicallyMerge)
        {
            try
            {
                if (_viewModelMain.IncidsSelectedMapCount <= 0) return;

                _mergeFeaturesWindow = new WindowMergeFeatures();
                _mergeFeaturesWindow.Owner = App.Current.MainWindow;
                _mergeFeaturesWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                _mergeFeaturesWindow.MaxHeight = App.Current.MainWindow.ActualHeight;

                HluDataSet.incidDataTable selectTable = new HluDataSet.incidDataTable();
                _viewModelMain.HluTableAdapterManager.incidTableAdapter.Fill(selectTable,
                    ViewModelWindowMainHelpers.IncidSelectionToWhereClause(ViewModelWindowMain.IncidPageSize,
                    _viewModelMain.IncidTable.incidColumn.Ordinal, _viewModelMain.IncidTable, _viewModelMain.IncidsSelectedMap));

                HluDataSet.incid_mm_polygonsDataTable polygons = new HluDataSet.incid_mm_polygonsDataTable();
                _viewModelMain.GetIncidMMPolygonRows(ViewModelWindowMainHelpers.GisSelectionToWhereClause(
                    _viewModelMain.GisSelection.Select(), _viewModelMain.GisIDColumnOrdinals, 
                    ViewModelWindowMain.IncidPageSize, polygons), ref polygons);

                _mergeFeaturesViewModelLogical = new ViewModelMergeFeatures<HluDataSet.incidDataTable,
                    HluDataSet.incidRow>(selectTable, _viewModelMain.GisIDColumnOrdinals, 
                    _viewModelMain.IncidTable.incidColumn.Ordinal, polygons.Select(r => r).ToArray(), 
                    _viewModelMain.GISApplication);
                _mergeFeaturesViewModelLogical.DisplayName = "Select INCID To Keep";
                _mergeFeaturesViewModelLogical.RequestClose += new ViewModelMergeFeatures<HluDataSet.incidDataTable,
                        HluDataSet.incidRow>.RequestCloseEventHandler(_mergeFeaturesViewModelLogical_RequestClose);

                _mergeFeaturesWindow.DataContext = _mergeFeaturesViewModelLogical;
                _mergeResultFeatureIndex = -1;

                _mergeFeaturesWindow.ShowDialog();

                if (_mergeResultFeatureIndex != -1)
                {
                    _viewModelMain.ChangeCursor(Cursors.Wait, "Processing ...");

                    _viewModelMain.DataBase.BeginTransaction(true, IsolationLevel.ReadCommitted);

                    try
                    {
                        string keepIncid = selectTable[_mergeResultFeatureIndex].incid;

                        // assign selected incid to selected features except keepIncid
                        DataTable historyTable = _viewModelMain.GISApplication.MergeFeaturesLogically(
                            keepIncid, _viewModelMain.HistoryColumns);

                        if ((historyTable == null) || (historyTable.Rows.Count == 0))
                            throw new Exception("Failed to update GIS layer.");

                        // assign selected incid and new toid_fragment_id to selected features except keepIncid in DB shadow copy
                        string toidFragmentFormat = String.Format("D{0}",
                            _viewModelMain.HluDataset.incid_mm_polygons.toid_fragment_idColumn.MaxLength);

                        List<KeyValuePair<int, object>> updateFields = new List<KeyValuePair<int,object>>();
                        var keepPolygon = polygons.FirstOrDefault(r => r.incid == keepIncid);
                        if (keepPolygon != null)
                        {
                            updateFields = (from c in polygons.Columns.Cast<DataColumn>()
                                           where (c.Ordinal != _viewModelMain.HluDataset.incid_mm_polygons.incidColumn.Ordinal) &&
                                             (c.Ordinal != _viewModelMain.HluDataset.incid_mm_polygons.toidColumn.Ordinal) &&
                                             (c.Ordinal != _viewModelMain.HluDataset.incid_mm_polygons.toid_fragment_idColumn.Ordinal) &&
                                             (c.Ordinal != _viewModelMain.HluDataset.incid_mm_polygons.shape_lengthColumn.Ordinal) &&
                                             (c.Ordinal != _viewModelMain.HluDataset.incid_mm_polygons.shape_areaColumn.Ordinal)
                                           select new KeyValuePair<int, object>(c.Ordinal, keepPolygon[c.Ordinal])).ToList();
                        }

                        var updatePolygons = from r in polygons
                                             where r.incid != keepIncid
                                             orderby r.toid, r.toid_fragment_id
                                             select r;

                        // update shadow DB copy of GIS layer
                        foreach (HluDataSet.incid_mm_polygonsRow r in updatePolygons)
                        {
                            r.incid = keepIncid;
                            for (int i = 0; i < updateFields.Count; i++)
                                r[updateFields[i].Key] = updateFields[i].Value;
                        }

                        if (_viewModelMain.HluTableAdapterManager.incid_mm_polygonsTableAdapter.Update(polygons) == -1)
                            throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_mm_polygons.TableName));

                        // insert history rows (fixed value keepIncid)
                        Dictionary<int, string> fixedValues = new Dictionary<int, string>();
                        fixedValues.Add(_viewModelMain.HluDataset.history.incidColumn.Ordinal, keepIncid);
                        ViewModelWindowMainHistory vmHist = new ViewModelWindowMainHistory(_viewModelMain);
                        vmHist.HistoryWrite(fixedValues, historyTable, ViewModelWindowMain.Operations.LogicalMerge);

                        // count incid records no longer in use
                        List<string> deleteIncids = new List<string>();

                        IDataReader delReader = _viewModelMain.DataBase.ExecuteReader(String.Format(
                            "SELECT {0} FROM {1} WHERE {0} IN ({2}) GROUP BY {0} HAVING COUNT(*) = 0",
                            _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.IncidTable.incidColumn.ColumnName),
                            _viewModelMain.DataBase.QualifyTableName(_viewModelMain.IncidTable.TableName), 
                            selectTable.Aggregate(new StringBuilder(), (sb, r) => sb.Append("," + 
                                _viewModelMain.DataBase.QuoteValue(r.incid))).Remove(0, 1)),
                            _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text);

                        if (delReader == null) throw new Exception("Error reading incid database table.");

                        while (delReader.Read())
                            deleteIncids.Add(delReader.GetString(0));
                        delReader.Close();

                        // delete any incid records no longer in use
                        if (deleteIncids.Count > 0)
                        {
                            int numAffected = _viewModelMain.DataBase.ExecuteNonQuery(String.Format(
                                "DELETE FROM {0} WHERE {1} IN ({2})",
                                _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid.TableName),
                                _viewModelMain.DataBase.QuoteValue(_viewModelMain.HluDataset.incid.incidColumn.ColumnName),
                                String.Join(",", deleteIncids.Select(i => _viewModelMain.DataBase.QuoteValue(i)).ToArray())),
                                _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text);

                            if (numAffected > 0) _viewModelMain.IncidRowCount(true);
                        }

                        _viewModelMain.ViewModelUpdate.UpdateIncidModifiedColumns(keepIncid);

                        if (physicallyMerge && (MessageBox.Show("Perform physical merge as well?", "HLU: Physical Merge",
                            MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes))
                        {
                            // restore the selection
                            _viewModelMain.GisSelection.Clear();
                            foreach (HluDataSet.incid_mm_polygonsRow r in polygons)
                            {
                                DataRow newRow = _viewModelMain.GisSelection.NewRow();
                                for (int i = 0; i < _viewModelMain.GisIDColumnOrdinals.Length; i++)
                                    newRow[i] = r[_viewModelMain.GisIDColumnOrdinals[i]];
                                _viewModelMain.GisSelection.Rows.Add(newRow);
                            }

                            PerformPhysicalMerge();
                        }
                        else
                        {
                            _viewModelMain.DataBase.CommitTransaction();
                            _viewModelMain.HluDataset.AcceptChanges();

                            // Re-count the incid records in the database.
                            _viewModelMain.IncidRowCount(true);

                            // Reset the incid and map selections but don't move
                            // to the first incid in the database.
                            _viewModelMain.ClearFilter(false);

                            // Synch with the GIS selection.
                            //---------------------------------------------------------------------
                            // FIX: 027 Force refill of Incid table after split/merge
                            // Force the Incid table to be refilled because it has been
                            // updated directly in the database rather than via the
                            // local copy.
                            _viewModelMain.RefillIncidTable = true;
                            //---------------------------------------------------------------------
                            _viewModelMain.ReadMapSelection(true);
                        }
                    }
                    catch
                    {
                        _viewModelMain.DataBase.RollbackTransaction();
                        throw;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Merge operation failed. The error message returned was:\n\n" +
                    ex.Message, "HLU Merge Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally { _viewModelMain.ChangeCursor(Cursors.Arrow, null); }
        }

        private void _mergeFeaturesViewModelLogical_RequestClose(int selectedIndex)
        {
            _mergeFeaturesViewModelLogical.RequestClose -= _mergeFeaturesViewModelLogical_RequestClose;
            _mergeFeaturesWindow.Close();

            _mergeResultFeatureIndex = selectedIndex;
        }

        private void PerformPhysicalMerge()
        {
            try
            {
                HluDataSet.incid_mm_polygonsDataTable selectTable = new HluDataSet.incid_mm_polygonsDataTable();
                _viewModelMain.GetIncidMMPolygonRows(ViewModelWindowMainHelpers.GisSelectionToWhereClause(
                    _viewModelMain.GisSelection.Select(), _viewModelMain.GisIDColumnOrdinals, 
                    ViewModelWindowMain.IncidPageSize, selectTable), ref selectTable);

                if (selectTable.Count == 0)
                    return;
                else if (selectTable.Count != _viewModelMain.GisSelection.Rows.Count)
                    throw new Exception(String.Format("GIS Layer and database are out of sync:\n{0} map polygons, {1} rows in table {2}.", 
                        _viewModelMain.FragsSelectedMapCount, selectTable.Count, _viewModelMain.HluDataset.incid_mm_polygons.TableName));

                // lowest toid_fragment_id in selection assigned to result feature
                string newToidFragmentID = selectTable.Min(r => r.toid_fragment_id);

                if (selectTable.GroupBy(r => r.incid).Count() == 1)
                {
                    int minFragmID = Int32.Parse(newToidFragmentID);
                    _mergeResultFeatureIndex = selectTable.Select((r, index) => 
                        Int32.Parse(r.toid_fragment_id) == minFragmID ? index : -1).First(i => i != -1);
                }
                else
                {
                    _mergeFeaturesWindow = new WindowMergeFeatures();
                    _mergeFeaturesWindow.Owner = App.Current.MainWindow;
                    _mergeFeaturesWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;

                    _mergeFeaturesViewModelPhysical = new ViewModelMergeFeatures
                        <HluDataSet.incid_mm_polygonsDataTable, HluDataSet.incid_mm_polygonsRow>(selectTable,
                        _viewModelMain.GisIDColumnOrdinals, _viewModelMain.IncidTable.incidColumn.Ordinal,
                        null, _viewModelMain.GISApplication);
                    _mergeFeaturesViewModelPhysical.DisplayName = "Select Feature To Keep";
                    _mergeFeaturesViewModelPhysical.RequestClose += new ViewModelMergeFeatures
                        <HluDataSet.incid_mm_polygonsDataTable, HluDataSet.incid_mm_polygonsRow>
                        .RequestCloseEventHandler(_mergeFeaturesViewModelPhysical_RequestClose);

                    _mergeFeaturesWindow.DataContext = _mergeFeaturesViewModelPhysical;
                    _mergeResultFeatureIndex = -1;

                    _mergeFeaturesWindow.ShowDialog();
                }

                if (_mergeResultFeatureIndex != -1)
                {
                    _viewModelMain.ChangeCursor(Cursors.Wait, "Processing ...");
                    bool startTransaction = _viewModelMain.DataBase.Transaction != null;

                    if (startTransaction)
                        _viewModelMain.DataBase.BeginTransaction(true, IsolationLevel.ReadCommitted);

                    try
                    {
                        _viewModelMain.ViewModelUpdate.UpdateIncidModifiedColumns(_viewModelMain.IncidsSelectedMap.ElementAt(0));

                        List<List<SqlFilterCondition>> resultFeatureWhereClause =
                            ViewModelWindowMainHelpers.GisSelectionToWhereClause(
                            new HluDataSet.incid_mm_polygonsRow[] { selectTable[_mergeResultFeatureIndex] },
                                _viewModelMain.GisIDColumnOrdinals, ViewModelWindowMain.IncidPageSize, selectTable);

                        if (resultFeatureWhereClause.Count != 1)
                            throw new Exception("Error getting result feature from database.");

                        List<List<SqlFilterCondition>> mergeFeaturesWhereClause =
                            ViewModelWindowMainHelpers.GisSelectionToWhereClause(
                            selectTable.Where((r, index) => index != _mergeResultFeatureIndex).ToArray(),
                                _viewModelMain.GisIDColumnOrdinals, ViewModelWindowMain.IncidPageSize, selectTable);

                        // historyTable contains rows of features merged into result feature (i.e. no longer existing)
                        // and last row with data of result feature (remaining in GIS, lowest toid_fragment_id of merged features)
                        // this last row must be removed before writing history 
                        // but is needed to update geometry fields in incid_mm_polygons
                        DataTable historyTable = _viewModelMain.GISApplication.MergeFeatures(newToidFragmentID,
                            resultFeatureWhereClause[0].Select(c => c.Clone()).ToList(), _viewModelMain.HistoryColumns);

                        if (historyTable == null)
                            throw new Exception("GIS merge operation failed.");

                        DataTable resultTable = historyTable.Clone();
                        DataRow resultRow = historyTable.AsEnumerable().FirstOrDefault(r =>
                            r.Field<string>(_viewModelMain.HluDataset.history.toid_fragment_idColumn.ColumnName) == newToidFragmentID);
                        if (resultRow == null) throw new Exception(String.Format(
                            "Failed to obtain geometry data of result feature from {0}.", _viewModelMain.GISApplication.ApplicationType));
                        resultTable.LoadDataRow(resultRow.ItemArray, true);
                        resultRow.Delete();
                        historyTable.AcceptChanges();

                        // synchronize DB shadow copy of GIS layer
                        MergeSynchronizeIncidMMPolygons(selectTable, resultTable, newToidFragmentID,
                            resultFeatureWhereClause[0], mergeFeaturesWhereClause);

                        // insert history rows (fixed values incid, toid, newToidFragmentID)
                        Dictionary<int, string> fixedValues = new Dictionary<int, string>();
                        fixedValues.Add(_viewModelMain.HluDataset.history.incidColumn.Ordinal, selectTable[0].incid);
                        fixedValues.Add(_viewModelMain.HluDataset.history.toidColumn.Ordinal, selectTable[0].toid);
                        fixedValues.Add(_viewModelMain.HluDataset.history.toid_fragment_idColumn.Ordinal, newToidFragmentID);
                        ViewModelWindowMainHistory vmHist = new ViewModelWindowMainHistory(_viewModelMain);
                        vmHist.HistoryWrite(fixedValues, historyTable, ViewModelWindowMain.Operations.PhysicalMerge);

                        if (startTransaction)
                        {
                            _viewModelMain.DataBase.CommitTransaction();
                            _viewModelMain.HluDataset.AcceptChanges();
                        }

                        // Re-count the incid records in the database.
                        _viewModelMain.IncidRowCount(true);

                        // Reset the incid and map selections but don't move
                        // to the first incid in the database.
                        _viewModelMain.ClearFilter(false);

                        // Synch with the GIS selection.
                        //---------------------------------------------------------------------
                        // FIX: 027 Force refill of Incid table after split/merge
                        // Force the Incid table to be refilled because it has been
                        // updated directly in the database rather than via the
                        // local copy.
                        _viewModelMain.RefillIncidTable = true;
                        //---------------------------------------------------------------------
                        _viewModelMain.ReadMapSelection(true);
                    }
                    catch
                    {
                        if (startTransaction) _viewModelMain.DataBase.RollbackTransaction();
                        throw;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Merge operation failed. The error message returned was:\n\n" +
                    ex.Message, "HLU Merge Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally { _viewModelMain.ChangeCursor(Cursors.Arrow, null); }
        }

        private void _mergeFeaturesViewModelPhysical_RequestClose(int selectedIndex)
        {
            _mergeFeaturesViewModelPhysical.RequestClose -= _mergeFeaturesViewModelPhysical_RequestClose;
            _mergeFeaturesWindow.Close();

            _mergeResultFeatureIndex = selectedIndex;
        }

        private void MergeSynchronizeIncidMMPolygons(HluDataSet.incid_mm_polygonsDataTable selectTable,
            DataTable resultTable, string newToidFragmentID, List<SqlFilterCondition> resultFeatureWhereClause,
            List<List<SqlFilterCondition>> mergeFeaturesWhereClause)
        {
            // build an update statement for the result feature: lowest toid_fragment_id 
            // in the selection set and sum of shape_length/shape_area of merged features
            string updateWhereClause = _viewModelMain.DataBase.WhereClause(false, true, true, resultFeatureWhereClause);
            string updateStatement = null;
            switch (_viewModelMain.GisLayerType)
            {
                case ViewModelWindowMain.GeometryTypes.Point:
                    updateStatement = String.Format("UPDATE {0} SET {1} = {2} WHERE {3}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_mm_polygons.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(
                            _viewModelMain.HluDataset.incid_mm_polygons.toid_fragment_idColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue(newToidFragmentID), updateWhereClause);
                    break;
                case ViewModelWindowMain.GeometryTypes.Line:
                    double plineLength = resultTable.Rows[0].Field<double>(ViewModelWindowMain.HistoryGeometry1ColumnName);
                    updateStatement = String.Format("UPDATE {0} SET {1} = {2}, {3} = {4} WHERE {5}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_mm_polygons.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(
                            _viewModelMain.HluDataset.incid_mm_polygons.toid_fragment_idColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue(newToidFragmentID),
                        _viewModelMain.DataBase.QuoteIdentifier(
                            _viewModelMain.HluDataset.incid_mm_polygons.shape_lengthColumn.ColumnName),
                        plineLength, updateWhereClause);
                    break;
                case ViewModelWindowMain.GeometryTypes.Polygon:
                    double shapeLength = resultTable.Rows[0].Field<double>(ViewModelWindowMain.HistoryGeometry1ColumnName);
                    double shapeArea = resultTable.Rows[0].Field<double>(ViewModelWindowMain.HistoryGeometry2ColumnName);
                    updateStatement = String.Format("UPDATE {0} SET {1} = {2}, {3} = {4}, {5} = {6} WHERE {7}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_mm_polygons.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(
                            _viewModelMain.HluDataset.incid_mm_polygons.toid_fragment_idColumn.ColumnName),
                        _viewModelMain.DataBase.QuoteValue(newToidFragmentID),
                        _viewModelMain.DataBase.QuoteIdentifier(
                            _viewModelMain.HluDataset.incid_mm_polygons.shape_lengthColumn.ColumnName), shapeLength,
                        _viewModelMain.DataBase.QuoteIdentifier(
                            _viewModelMain.HluDataset.incid_mm_polygons.shape_areaColumn.ColumnName),
                        shapeArea, updateWhereClause);
                    break;
            }

            bool startTransaction = _viewModelMain.DataBase.Transaction == null;

            if (startTransaction) _viewModelMain.DataBase.BeginTransaction(true, IsolationLevel.ReadCommitted);
            try
            {
                // delete merged polygons from shadow table in DB
                List<List<SqlFilterCondition>> cleanList = _viewModelMain.DataBase.JoinWhereClauseLists(mergeFeaturesWhereClause);
                foreach (List<SqlFilterCondition> oneWhereClause in cleanList)
                {
                    if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format("DELETE FROM {0} WHERE {1}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_mm_polygons.TableName),
                        _viewModelMain.DataBase.WhereClause(false, true, true, oneWhereClause)),
                        _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                        throw new Exception(String.Format("Failed to delete from table {0}.", 
                            _viewModelMain.HluDataset.incid_mm_polygons.TableName));
                }

                // update the result feature
                if (_viewModelMain.DataBase.ExecuteNonQuery(updateStatement, 
                    _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                    throw new Exception(String.Format("Failed to update table {0}.", 
                        _viewModelMain.HluDataset.incid_mm_polygons.TableName));

                if (startTransaction) _viewModelMain.DataBase.CommitTransaction();
            }
            catch
            {
                if (startTransaction) _viewModelMain.DataBase.RollbackTransaction();
                throw;
            }
        }
    }
}
