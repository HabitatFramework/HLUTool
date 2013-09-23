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

namespace HLU.UI.ViewModel
{
    class ViewModelWindowMainBulkUpdate
    {
        private ViewModelWindowMain _viewModelMain;
        private HluDataSet.incid_ihs_matrixDataTable _ihsMatrixTable = new HluDataSet.incid_ihs_matrixDataTable();
        private HluDataSet.incid_ihs_formationDataTable _ihsFormationTable = new HluDataSet.incid_ihs_formationDataTable();
        private HluDataSet.incid_ihs_managementDataTable _ihsManagementTable = new HluDataSet.incid_ihs_managementDataTable();
        private HluDataSet.incid_ihs_complexDataTable _ihsComplexTable = new HluDataSet.incid_ihs_complexDataTable();

        public ViewModelWindowMainBulkUpdate(ViewModelWindowMain viewModelMain)
        {
            _viewModelMain = viewModelMain;
        }

        public void StartBulkUpdate()
        {
            _viewModelMain.BulkUpdateMode = true;
            _viewModelMain.TabItemHistoryEnabled = false;

            _viewModelMain.IhsCategory = null;
            _viewModelMain.IncidIhsHabitat = null;

            _viewModelMain.IncidCurrentRow = _viewModelMain.HluDataset.incid.NewincidRow();

            _viewModelMain.IncidIhsMatrixRows = new HluDataSet.incid_ihs_matrixRow[3]
                .Select(r => _viewModelMain.HluDataset.incid_ihs_matrix.Newincid_ihs_matrixRow()).ToArray();
            for (int i = 0; i < _viewModelMain.IncidIhsMatrixRows.Length; i++)
            {
                _viewModelMain.IncidIhsMatrixRows[i].matrix_id = i;
                _viewModelMain.IncidIhsMatrixRows[i].incid = _viewModelMain.RecIDs.CurrentIncid;
            }

            _viewModelMain.IncidIhsFormationRows = new HluDataSet.incid_ihs_formationRow[2]
                .Select(r => _viewModelMain.HluDataset.incid_ihs_formation.Newincid_ihs_formationRow()).ToArray();
            for (int i = 0; i < _viewModelMain.IncidIhsFormationRows.Length; i++)
            {
                _viewModelMain.IncidIhsFormationRows[i].formation_id = i;
                _viewModelMain.IncidIhsFormationRows[i].incid = _viewModelMain.RecIDs.CurrentIncid;
            }

            _viewModelMain.IncidIhsManagementRows = new HluDataSet.incid_ihs_managementRow[2]
                .Select(r => _viewModelMain.HluDataset.incid_ihs_management.Newincid_ihs_managementRow()).ToArray();
            for (int i = 0; i < _viewModelMain.IncidIhsManagementRows.Length; i++)
            {
                _viewModelMain.IncidIhsManagementRows[i].management_id = i;
                _viewModelMain.IncidIhsManagementRows[i].incid = _viewModelMain.RecIDs.CurrentIncid;
            }

            _viewModelMain.IncidIhsComplexRows = new HluDataSet.incid_ihs_complexRow[2]
                .Select(r => _viewModelMain.HluDataset.incid_ihs_complex.Newincid_ihs_complexRow()).ToArray();
            for (int i = 0; i < _viewModelMain.IncidIhsComplexRows.Length; i++)
            {
                _viewModelMain.IncidIhsComplexRows[i].complex_id = i;
                _viewModelMain.IncidIhsComplexRows[i].incid = _viewModelMain.RecIDs.CurrentIncid;
            }

            _viewModelMain.IncidBapRows = new HluDataSet.incid_bapRow[0]
                .Select(r => _viewModelMain.HluDataset.incid_bap.Newincid_bapRow()).ToArray();
            _viewModelMain.IncidBapRowsAuto = new ObservableCollection<BapEnvironment>();
            _viewModelMain.IncidBapRowsUser = new ObservableCollection<BapEnvironment>();

            _viewModelMain.IncidSourcesRows = new HluDataSet.incid_sourcesRow[3]
                .Select(r => _viewModelMain.HluDataset.incid_sources.Newincid_sourcesRow()).ToArray();
            for (int i = 0; i < _viewModelMain.IncidSourcesRows.Length; i++)
            {
                _viewModelMain.IncidSourcesRows[i].incid_source_id = i;
                _viewModelMain.IncidSourcesRows[i].source_id = Int32.MinValue;
                _viewModelMain.IncidSourcesRows[i].incid = _viewModelMain.RecIDs.CurrentIncid;
            }

            _viewModelMain.RefreshAll();
        }

        public void BulkUpdate()
        {
            _viewModelMain.ChangeCursor(Cursors.Wait, "Bulk updating...");

            _viewModelMain.DataBase.BeginTransaction(true, IsolationLevel.ReadCommitted);

            try
            {
                string incidWhereClause = String.Format(" WHERE {0} = {1}",
                    _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid.incidColumn.ColumnName), 
                    _viewModelMain.DataBase.QuoteValue("{0}"));

                // build select statement for incid row
                string selectCommandIncid = String.Format("SELECT {0} FROM {1}{2}",
                    _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.incid.ihs_habitatColumn.ColumnName),
                    _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid.TableName), incidWhereClause);

                // build update statement for incid table
                ViewModelWindowMainUpdate.IncidCurrentRowDerivedValuesUpdate(_viewModelMain);
                var incidUpdateVals = _viewModelMain.HluDataset.incid.Columns.Cast<DataColumn>()
                    .Where(c => c.Ordinal != _viewModelMain.HluDataset.incid.incidColumn.Ordinal &&
                        !_viewModelMain.IncidCurrentRow.IsNull(c.Ordinal));

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

                bool deleteExtraRows = Settings.Default.BulkUpdateBlankRowMeansDelete;

                SqlFilterCondition incidWhereCond = _viewModelMain.ChildRowFilter(_viewModelMain.HluDataset.incid, 
                    _viewModelMain.HluDataset.incid.incidColumn);
                List<SqlFilterCondition> incidWhereConds =
                    new List<SqlFilterCondition>(new SqlFilterCondition[] { incidWhereCond });
                int incidOrdinal = 
                    _viewModelMain.IncidSelection.Columns[_viewModelMain.HluDataset.incid.incidColumn.ColumnName].Ordinal;

                // prepare delete statements for IHS multiplex rows rendered obsolete by new IHS habitat
                List<string> ihsMultiplexDelStatements = new List<string>();
                if (incidUpdateVals.Contains(_viewModelMain.HluDataset.incid.ihs_habitatColumn))
                {
                    string newIncidIhsHabitatCode = _viewModelMain.IncidCurrentRow[_viewModelMain.HluDataset.incid.ihs_habitatColumn.Ordinal].ToString();

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

                if (Settings.Default.BulkUpdateUsesAdo)
                    BulkUpdateAdo(incidOrdinal, incidWhereCond, incidWhereConds, selectCommandIncid, 
                        updateCommandIncid, deleteExtraRows, ihsMultiplexDelStatements);
                else
                    BulkUpdateDb(incidOrdinal, incidWhereCond, incidWhereConds, selectCommandIncid, 
                        updateCommandIncid, deleteExtraRows, ihsMultiplexDelStatements);

                // update GIS, shadow copy in DB and history
                BulkUpdateGis(incidOrdinal, _viewModelMain.IncidSelection);

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

        private void BulkUpdateDb(int incidOrdinal, SqlFilterCondition incidWhereCond,
            List<SqlFilterCondition> incidWhereConds, string selectCommandIncid,
            string updateCommandIncid, bool deleteExtraRows, List<string> ihsMultiplexDelStatements)
        {
            string currIncid;

            foreach (DataRow r in _viewModelMain.IncidSelection.Rows)
            {
                incidWhereCond.Value = r[incidOrdinal];
                incidWhereConds[0] = incidWhereCond;

                currIncid = incidWhereCond.Value.ToString();

                // delete any orphaned IHS multiplex rows
                if (ihsMultiplexDelStatements != null)
                {
                    foreach (string s in ihsMultiplexDelStatements)
                        if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format(s, currIncid),
                            _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                            throw new Exception("Failed to delete IHS multiplex rows.");
                }

                // execute update statement for each row in selection
                if (!String.IsNullOrEmpty(updateCommandIncid))
                {
                    if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format(updateCommandIncid, currIncid),
                        _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                        throw new Exception("Failed to update incid table.");
                }

                object retValue = _viewModelMain.DataBase.ExecuteScalar(selectCommandIncid,
                    _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text);

                if (retValue == null) throw new Exception(
                    String.Format("No database row for incid '{0}'", currIncid));

                string ihsHabitat = retValue.ToString();

                object[] relValues = new object[] { currIncid };

                HluDataSet.incid_ihs_matrixRow[] incidIhsMatrixRows = _viewModelMain.IncidIhsMatrixRows;
                HluDataSet.incid_ihs_matrixDataTable ihsMatrixTable = 
                    (HluDataSet.incid_ihs_matrixDataTable)_viewModelMain.HluDataset.incid_ihs_matrix.Copy();
                BulkUpdateDbMultiplexSourceTable(deleteExtraRows, currIncid, relValues,
                    _viewModelMain.HluTableAdapterManager.incid_ihs_matrixTableAdapter, 
                    ihsMatrixTable, ref incidIhsMatrixRows);

                HluDataSet.incid_ihs_formationRow[] incidIhsFormationRows = _viewModelMain.IncidIhsFormationRows;
                HluDataSet.incid_ihs_formationDataTable ihsFormationTable = 
                    (HluDataSet.incid_ihs_formationDataTable)_viewModelMain.HluDataset.incid_ihs_formation.Copy();
                BulkUpdateDbMultiplexSourceTable(deleteExtraRows, currIncid, relValues,
                    _viewModelMain.HluTableAdapterManager.incid_ihs_formationTableAdapter, 
                    ihsFormationTable, ref incidIhsFormationRows);

                HluDataSet.incid_ihs_managementRow[] incidIhsManagementRows = _viewModelMain.IncidIhsManagementRows;
                HluDataSet.incid_ihs_managementDataTable ihsManagementTable = 
                    (HluDataSet.incid_ihs_managementDataTable)_viewModelMain.HluDataset.incid_ihs_management.Copy();
                BulkUpdateDbMultiplexSourceTable(deleteExtraRows, currIncid, relValues,
                    _viewModelMain.HluTableAdapterManager.incid_ihs_managementTableAdapter, 
                    ihsManagementTable, ref incidIhsManagementRows);

                HluDataSet.incid_ihs_complexRow[] incidIhsComplexRows = _viewModelMain.IncidIhsComplexRows;
                HluDataSet.incid_ihs_complexDataTable ihsComplexTable = 
                    (HluDataSet.incid_ihs_complexDataTable)_viewModelMain.HluDataset.incid_ihs_complex.Copy();
                BulkUpdateDbMultiplexSourceTable(deleteExtraRows, currIncid, relValues,
                    _viewModelMain.HluTableAdapterManager.incid_ihs_complexTableAdapter, 
                    ihsComplexTable, ref incidIhsComplexRows);

                HluDataSet.incid_bapDataTable bapTable = (HluDataSet.incid_bapDataTable)_viewModelMain.HluDataset.incid_bap.Copy();
                _viewModelMain.GetIncidChildRowsDb(relValues, _viewModelMain.HluTableAdapterManager.incid_bapTableAdapter, 
                    ref bapTable);
                BulkUpdateBap(currIncid, ihsHabitat, bapTable, incidIhsMatrixRows, incidIhsFormationRows,
                    incidIhsManagementRows, incidIhsComplexRows, deleteExtraRows);

                HluDataSet.incid_sourcesDataTable incidSourcesTable = 
                    (HluDataSet.incid_sourcesDataTable)_viewModelMain.HluDataset.incid_sources.Copy();
                HluDataSet.incid_sourcesRow[] incidSourcesRows = _viewModelMain.IncidSourcesRows;
                BulkUpdateDbMultiplexSourceTable(deleteExtraRows, currIncid, relValues,
                    _viewModelMain.HluTableAdapterManager.incid_sourcesTableAdapter, 
                    incidSourcesTable, ref incidSourcesRows);
            }
        }

        private void BulkUpdateAdo(int incidOrdinal, SqlFilterCondition incidWhereCond,
            List<SqlFilterCondition> incidWhereConds, string selectCommandIncid,
            string updateCommandIncid, bool deleteExtraRows, List<string> ihsMultiplexDelStatements)
        {
            foreach (DataRow r in _viewModelMain.IncidSelection.Rows)
            {
                string currIncid = r[incidOrdinal].ToString();

                incidWhereCond.Value = r[incidOrdinal];
                incidWhereConds[0] = incidWhereCond;

                // delete any orphaned IHS multiplex rows
                if (ihsMultiplexDelStatements != null)
                {
                    foreach (string s in ihsMultiplexDelStatements)
                        if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format(s, currIncid), 
                            _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                            throw new Exception("Failed to delete IHS multiplex rows.");
                }

                // execute update statement for each row in selection
                if (!String.IsNullOrEmpty(updateCommandIncid))
                {
                    if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format(updateCommandIncid, r[incidOrdinal]),
                        _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                        throw new Exception("Failed to update incid table.");
                }

                object retValue = _viewModelMain.DataBase.ExecuteScalar(String.Format(selectCommandIncid, currIncid),
                    _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text);

                if (retValue == null) throw new Exception(
                    String.Format("No database row for incid '{0}'", currIncid));
                string ihsHabitat = retValue.ToString();

                object[] relValues = new object[] { r[incidOrdinal] };

                HluDataSet.incid_ihs_matrixRow[] incidIhsMatrixRows = _viewModelMain.IncidIhsMatrixRows;
                HluDataSet.incid_ihs_matrixDataTable ihsMatrixTable = 
                    (HluDataSet.incid_ihs_matrixDataTable)_viewModelMain.HluDataset.incid_ihs_matrix.Copy();
                BulkUpdateAdoMultiplexSourceTable(deleteExtraRows, currIncid, relValues,
                    _viewModelMain.HluTableAdapterManager.incid_ihs_matrixTableAdapter, 
                    ihsMatrixTable, ref incidIhsMatrixRows);
                _viewModelMain.IncidIhsMatrixRows = incidIhsMatrixRows;

                HluDataSet.incid_ihs_formationRow[] incidIhsFormationRows = _viewModelMain.IncidIhsFormationRows;
                HluDataSet.incid_ihs_formationDataTable ihsFormationTable = 
                    (HluDataSet.incid_ihs_formationDataTable)_viewModelMain.HluDataset.incid_ihs_formation.Copy();
                BulkUpdateAdoMultiplexSourceTable(deleteExtraRows, currIncid, relValues,
                    _viewModelMain.HluTableAdapterManager.incid_ihs_formationTableAdapter, 
                    ihsFormationTable, ref incidIhsFormationRows);

                HluDataSet.incid_ihs_managementRow[] incidIhsManagementRows = _viewModelMain.IncidIhsManagementRows;
                HluDataSet.incid_ihs_managementDataTable ihsManagementTable = 
                    (HluDataSet.incid_ihs_managementDataTable)_viewModelMain.HluDataset.incid_ihs_management.Copy();
                BulkUpdateAdoMultiplexSourceTable(deleteExtraRows, currIncid, relValues,
                    _viewModelMain.HluTableAdapterManager.incid_ihs_managementTableAdapter, 
                    ihsManagementTable, ref incidIhsManagementRows);
                _viewModelMain.IncidIhsManagementRows = incidIhsManagementRows;

                HluDataSet.incid_ihs_complexRow[] incidIhsComplexRows = _viewModelMain.IncidIhsComplexRows;
                HluDataSet.incid_ihs_complexDataTable ihsComplexTable = 
                    (HluDataSet.incid_ihs_complexDataTable)_viewModelMain.HluDataset.incid_ihs_complex.Copy();
                BulkUpdateAdoMultiplexSourceTable(deleteExtraRows, currIncid, relValues,
                    _viewModelMain.HluTableAdapterManager.incid_ihs_complexTableAdapter, 
                    ihsComplexTable, ref incidIhsComplexRows);
                _viewModelMain.IncidIhsComplexRows = incidIhsComplexRows;

                HluDataSet.incid_bapDataTable bapTable = (HluDataSet.incid_bapDataTable)_viewModelMain.HluDataset.incid_bap.Copy();
                _viewModelMain.GetIncidChildRowsDb(relValues, 
                    _viewModelMain.HluTableAdapterManager.incid_bapTableAdapter, ref bapTable);
                BulkUpdateBap(currIncid, ihsHabitat, bapTable, incidIhsMatrixRows, incidIhsFormationRows,
                    incidIhsManagementRows, incidIhsComplexRows, deleteExtraRows);

                HluDataSet.incid_sourcesDataTable incidSourcesTable = 
                    (HluDataSet.incid_sourcesDataTable)_viewModelMain.HluDataset.incid_sources.Copy();
                HluDataSet.incid_sourcesRow[] incidSourcesRows = _viewModelMain.IncidSourcesRows;
                BulkUpdateAdoMultiplexSourceTable(deleteExtraRows, currIncid, relValues,
                    _viewModelMain.HluTableAdapterManager.incid_sourcesTableAdapter, 
                    incidSourcesTable, ref incidSourcesRows);
                _viewModelMain.IncidSourcesRows = incidSourcesRows;
            }
        }

        private void BulkUpdateAdoMultiplexSourceTable<T, R>(bool deleteExtraRows, string currIncid,
            object[] relValues, HluTableAdapter<T, R> adapter, T dbTable, ref R[] uiRows)
            where T : DataTable, new()
            where R : DataRow
        {
            if ((uiRows != null) && (uiRows.Count(r =>
                !r.IsNull(dbTable.PrimaryKey[0].Ordinal)) > 0))
            {
                T newRows = CloneUpdateRows<T, R>(uiRows, currIncid);
                _viewModelMain.GetIncidChildRowsDb(relValues, adapter, ref dbTable);
                BulkUpdateChildTable<T, R>(deleteExtraRows, newRows.AsEnumerable()
                    .Cast<R>().ToArray(), ref dbTable, adapter);
                uiRows = newRows.AsEnumerable().Cast<R>().ToArray();
            }
        }

        private void BulkUpdateDbMultiplexSourceTable<T, R>(bool deleteExtraRows, string currIncid,
            object[] relValues, HluTableAdapter<T, R> adapter, T dbTable, ref R[] uiRows)
            where T : DataTable, new()
            where R : DataRow
        {
            if ((uiRows != null) && (uiRows.Count(rm =>
                !rm.IsNull(dbTable.PrimaryKey[0].Ordinal)) > 0))
            {
                T newRows = CloneUpdateRows<T, R>(uiRows, currIncid);
                _viewModelMain.GetIncidChildRowsDb(relValues, adapter, ref dbTable);
                BulkUpdateChildTable<T, R>(deleteExtraRows, newRows.AsEnumerable()
                    .Cast<R>().ToArray(), ref dbTable);
                uiRows = newRows.AsEnumerable().Cast<R>().ToArray();
            }
        }

        private R[] FilterUpdateRows<T, R>(R[] rows)
            where T : DataTable
            where R : DataRow
        {
            if ((rows == null) || (rows.Length == 0)) return rows;

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

            List<R> newRows = new List<R>(rows.Length);

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
        /// Therefore, a bap_habitat must always be entered, even for existing records for which only other attributes
        /// are meant to be updated. 
        /// </summary>
        /// <param name="currIncid"></param>
        /// <param name="ihsHabitat"></param>
        /// <param name="incidBapTable"></param>
        /// <param name="ihsMatrixRows"></param>
        /// <param name="ihsFormationRows"></param>
        /// <param name="ihsManagementRows"></param>
        /// <param name="ihsComplexRows"></param>
        /// <param name="deleteExtraRows"></param>
        private void BulkUpdateBap(string currIncid, string ihsHabitat,
            HluDataSet.incid_bapDataTable incidBapTable,
            HluDataSet.incid_ihs_matrixRow[] ihsMatrixRows,
            HluDataSet.incid_ihs_formationRow[] ihsFormationRows,
            HluDataSet.incid_ihs_managementRow[] ihsManagementRows,
            HluDataSet.incid_ihs_complexRow[] ihsComplexRows,
            bool deleteExtraRows)
        {
            var mx = ihsMatrixRows.Where(r => r.RowState != DataRowState.Deleted).Select(r => r.matrix);
            string[] ihsMatrixVals = mx.Concat(new string[3 - mx.Count()]).ToArray();
            var fo = ihsFormationRows.Where(r => r.RowState != DataRowState.Deleted).Select(r => r.formation);
            string[] ihsFormationVals = fo.Concat(new string[2 - fo.Count()]).ToArray();
            var mg = ihsManagementRows.Where(r => r.RowState != DataRowState.Deleted).Select(r => r.management);
            string[] ihsManagementVals = mg.Concat(new string[2 - mg.Count()]).ToArray();
            var cx = ihsComplexRows.Where(r => r.RowState != DataRowState.Deleted).Select(r => r.complex);
            string[] ihsComplexVals = cx.Concat(new string[2 - cx.Count()]).ToArray();

            IEnumerable<string> potentialBap = _viewModelMain.PotentialBapEnvironments(ihsHabitat, 
                ihsMatrixVals[0], ihsMatrixVals[1], ihsMatrixVals[2], ihsFormationVals[0], ihsFormationVals[1],
                ihsManagementVals[0], ihsManagementVals[1], ihsComplexVals[0], ihsComplexVals[1]);

            int[] skipOrdinals = new int[3] { 
                _viewModelMain.HluDataset.incid_bap.bap_idColumn.Ordinal, 
                _viewModelMain.HluDataset.incid_bap.incidColumn.Ordinal, 
                _viewModelMain.HluDataset.incid_bap.bap_habitatColumn.Ordinal };

            List<HluDataSet.incid_bapRow> updateRows = new List<HluDataSet.incid_bapRow>();
            HluDataSet.incid_bapRow updateRow;

            // BAP environments from UI
            IEnumerable<BapEnvironment> beUI = from b in _viewModelMain.IncidBapRowsAuto.Concat(_viewModelMain.IncidBapRowsUser)
                                               group b by b.bap_habitat into habs
                                               select habs.First();

            foreach (BapEnvironment be in beUI)
            {
                // BAP environments from database
                IEnumerable<HluDataSet.incid_bapRow> dbRows =
                    incidBapTable.Where(r => r.bap_habitat == be.bap_habitat);

                bool isAdditional = !potentialBap.Contains(be.bap_habitat);
                if (isAdditional) be.MakeAdditional();

                switch (dbRows.Count())
                {
                    case 0: // insert newly added BAP environments
                        HluDataSet.incid_bapRow newRow = _viewModelMain.IncidBapTable.Newincid_bapRow();
                        newRow.ItemArray = be.ToItemArray(_viewModelMain.RecIDs.NextIncidBapId, currIncid);
                        if (be.IsValid(false, isAdditional, newRow)) // reset bulk update mode for full validation of a new row
                            _viewModelMain.HluTableAdapterManager.incid_bapTableAdapter.Insert(newRow);
                        break;
                    case 1: // update existing row
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

            incidBapTable.Where(r => !potentialBap.Contains(r.bap_habitat)).ToList().ForEach(delegate(HluDataSet.incid_bapRow r)
            {
                updateRows.Add(BapEnvironment.MakeAdditional(r));
            });

            if (updateRows.Count > 0)
            {
                if (_viewModelMain.HluTableAdapterManager.incid_bapTableAdapter.Update(updateRows.ToArray()) == -1)
                    throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_bap.TableName));
            }

            // delete non-potential BAP environments from DB that are not in _viewModelMain.IncidBapRowsUser
            if (deleteExtraRows)
            {
                var delRows = incidBapTable.Where(r => !potentialBap.Contains(r.bap_habitat) && 
                    _viewModelMain.IncidBapRowsUser.Count(be => be.bap_habitat == r.bap_habitat) == 0);
                foreach (HluDataSet.incid_bapRow r in delRows)
                    _viewModelMain.HluTableAdapterManager.incid_bapTableAdapter.Delete(r);
            }
        }

        private void BulkUpdateGis(int incidOrdinal, DataTable incidSelection)
        {
            // get update columns and values
            DataColumn[] updateColumns;
            object[] updateValues;
            BulkUpdateGisColumns(out updateColumns, out updateValues);

            if ((updateColumns == null) || (updateColumns.Length == 0)) return;

            // update DB shadow copy of GIS layer
            string incidMMPolygonsUpdateCmdTemplate;
            List<List<SqlFilterCondition>> incidWhereClause;
            DataTable historyTable = null;

            int ixIhsSummary = System.Array.IndexOf(updateColumns, _viewModelMain.HluDataset.incid_mm_polygons.ihs_summaryColumn);

            if (ixIhsSummary == -1)
            {
                incidMMPolygonsUpdateCmdTemplate = String.Format("UPDATE {0} SET {1} WHERE {2}",
                    _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_mm_polygons.TableName),
                    updateColumns.Select((c, index) => new string[] { _viewModelMain.DataBase.QuoteIdentifier(c.ColumnName),
                        _viewModelMain.DataBase.QuoteValue(updateValues[index]) }).Aggregate(new StringBuilder(), (sb, a) =>
                            sb.Append(String.Format(", {0} = {1}", a[0], a[1]))).Remove(0, 2), "{0}");

                incidWhereClause = ViewModelWindowMainHelpers.IncidSelectionToWhereClause(ViewModelWindowMain.IncidPageSize, 
                    _viewModelMain.HluDataset.incid_mm_polygons.incidColumn.Ordinal, _viewModelMain.HluDataset.incid_mm_polygons, 
                    incidSelection.AsEnumerable().Select(r => r.Field<string>(incidOrdinal)));

                foreach (List<SqlFilterCondition> w in incidWhereClause)
                {
                    if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format(incidMMPolygonsUpdateCmdTemplate,
                        _viewModelMain.DataBase.WhereClause(false, true, true, w)),
                        _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                        throw new Exception(String.Format("Failed to update table {0}.",
                            _viewModelMain.HluDataset.incid_mm_polygons.TableName));
                }

                // update GIS layer
                ScratchDb.WriteSelectionScratchTable(_viewModelMain.GisIDColumns, _viewModelMain.IncidSelection);
                historyTable = _viewModelMain.GISApplication.UpdateFeatures(updateColumns, updateValues,
                    _viewModelMain.HistoryColumns, ScratchDb.ScratchMdbPath, ScratchDb.ScratchSelectionTable);
            }
            else
            {
                incidMMPolygonsUpdateCmdTemplate = String.Format("UPDATE {0} SET {1} WHERE {2}",
                    _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_mm_polygons.TableName),
                    updateColumns.Select((c, index) => new string[] { _viewModelMain.DataBase.QuoteIdentifier(c.ColumnName), 
                        c.ColumnName == _viewModelMain.HluDataset.incid_mm_polygons.ihs_summaryColumn.ColumnName ?
                        "{0}" : _viewModelMain.DataBase.QuoteValue(updateValues[index]) }).Aggregate(new StringBuilder(), 
                        (sb, a) => sb.Append(String.Format(", {0} = {1}", a[0], a[1]))).Remove(0, 2), "{1}");

                incidWhereClause = ViewModelWindowMainHelpers.IncidSelectionToWhereClause(1,
                    _viewModelMain.HluDataset.incid_mm_polygons.incidColumn.Ordinal, _viewModelMain.HluDataset.incid_mm_polygons,
                    incidSelection.AsEnumerable().Select(r => r.Field<string>(incidOrdinal)));

                foreach (List<SqlFilterCondition> w in incidWhereClause)
                {
                    updateValues[ixIhsSummary] = BuildIhsSummary(w[0].Value.ToString());

                    if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format(incidMMPolygonsUpdateCmdTemplate,
                        _viewModelMain.DataBase.QuoteValue(updateValues[ixIhsSummary]), 
                        _viewModelMain.DataBase.WhereClause(false, true, true, w)),
                        _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                        throw new Exception("Failed to update GIS layer shadow copy");

                    // update GIS layer row by row; no need for joined scratch table
                    DataTable historyTmp = _viewModelMain.GISApplication.UpdateFeatures(updateColumns, 
                        updateValues, _viewModelMain.HistoryColumns, w);

                    if (historyTmp == null) 
                        throw new Exception(String.Format("Failed to update GIS layer for incid '{0}'", w[0].Value));

                    // append history rows to history table
                    if (_viewModelMain.BulkUpdateCreateHistory)
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

            // write history for affected incids
            if (_viewModelMain.BulkUpdateCreateHistory && (historyTable != null))
            {
                ViewModelWindowMainHistory vmHist = new ViewModelWindowMainHistory(_viewModelMain);
                vmHist.HistoryWrite(null, historyTable, ViewModelWindowMain.Operations.BulkUpdate);
            }
        }

        private string BuildIhsSummary(string incid)
        {
            try
            {
                var incidRow = _viewModelMain.IncidTable.Where(r => r.incid == incid);
                string ihsHabitat = incidRow.Count() == 1 ? incidRow.ElementAt(0).ihs_habitat : String.Empty;

                _viewModelMain.GetIncidChildRowsDb(new object[] { incid },
                    _viewModelMain.HluTableAdapterManager.incid_ihs_matrixTableAdapter, ref _ihsMatrixTable);

                _viewModelMain.GetIncidChildRowsDb(new object[] { incid },
                    _viewModelMain.HluTableAdapterManager.incid_ihs_formationTableAdapter, ref _ihsFormationTable);

                _viewModelMain.GetIncidChildRowsDb(new object[] { incid },
                    _viewModelMain.HluTableAdapterManager.incid_ihs_managementTableAdapter, ref _ihsManagementTable);

                _viewModelMain.GetIncidChildRowsDb(new object[] { incid },
                    _viewModelMain.HluTableAdapterManager.incid_ihs_complexTableAdapter, ref _ihsComplexTable);

                object ihsSummaryValue = ViewModelWindowMainHelpers.IhsSummary(new string[] { ihsHabitat }
                    .Concat(_viewModelMain.IncidIhsMatrixRows.Select((ihs, index) =>  
                        !ihs.IsNull(_viewModelMain.HluDataset.incid_ihs_matrix.matrixColumn.Ordinal) ? 
                        ihs.matrix : _ihsMatrixTable.Rows.Count >= index && !_ihsMatrixTable.Rows[index]
                        .IsNull(_viewModelMain.HluDataset.incid_ihs_matrix.matrixColumn.Ordinal) ?
                        _ihsMatrixTable.Rows[index][_viewModelMain.HluDataset.incid_ihs_matrix.matrixColumn.Ordinal].ToString() :
                        String.Empty))
                    .Concat(_viewModelMain.IncidIhsFormationRows.Select((ihs, index) =>
                        !ihs.IsNull(_viewModelMain.HluDataset.incid_ihs_formation.formationColumn.Ordinal) ?
                        ihs.formation : _ihsFormationTable.Rows.Count >= index && !_ihsFormationTable.Rows[index]
                        .IsNull(_viewModelMain.HluDataset.incid_ihs_formation.formationColumn.Ordinal) ?
                        _ihsFormationTable.Rows[index][_viewModelMain.HluDataset.incid_ihs_formation.formationColumn.Ordinal].ToString() :
                        String.Empty))
                    .Concat(_viewModelMain.IncidIhsManagementRows.Select((ihs, index) =>
                        !ihs.IsNull(_viewModelMain.HluDataset.incid_ihs_management.managementColumn.Ordinal) ?
                        ihs.management : _ihsManagementTable.Rows.Count >= index && !_ihsManagementTable.Rows[index]
                        .IsNull(_viewModelMain.HluDataset.incid_ihs_management.managementColumn.Ordinal) ?
                        _ihsManagementTable.Rows[index][_viewModelMain.HluDataset.incid_ihs_management.managementColumn.Ordinal].ToString() :
                        String.Empty))
                    .Concat(_viewModelMain.IncidIhsComplexRows.Select((ihs, index) =>
                        !ihs.IsNull(_viewModelMain.HluDataset.incid_ihs_complex.complexColumn.Ordinal) ?
                        ihs.complex : _ihsComplexTable.Rows.Count >= index && !_ihsComplexTable.Rows[index]
                        .IsNull(_viewModelMain.HluDataset.incid_ihs_complex.complexColumn.Ordinal) ?
                        _ihsComplexTable.Rows[index][_viewModelMain.HluDataset.incid_ihs_complex.complexColumn.Ordinal].ToString() :
                        String.Empty)).ToArray());
                
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

        private void BulkUpdateGisColumns(out DataColumn[] updateColumns, out object[] updateValues)
        {
            List<DataColumn> updateColumnList = new List<DataColumn>();
            List<object> updateValueList = new List<object>();
            if ((_viewModelMain.HistoryColumns.Contains(_viewModelMain.HluDataset.incid_mm_polygons.ihs_categoryColumn)) &&
                (!_viewModelMain.IncidCurrentRow.Isihs_habitatNull()))
            {
                updateColumnList.Add(_viewModelMain.HluDataset.incid_mm_polygons.ihs_categoryColumn);
                updateValueList.Add(_viewModelMain.HluDataset.lut_ihs_habitat
                    .Single(h => h.code == _viewModelMain.IncidCurrentRow.ihs_habitat).category);
            }
            if (_viewModelMain.HistoryColumns.Contains(_viewModelMain.HluDataset.incid_mm_polygons.ihs_summaryColumn))
            {
                updateColumnList.Add(_viewModelMain.HluDataset.incid_mm_polygons.ihs_summaryColumn);
                updateValueList.Add(null);
            }

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

        private void BulkUpdateChildTable<T, R>(bool deleteExtraRows, R[] newRows, ref T dbRows)
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
            else if (deleteExtraRows && (numRowsDb > numRowsNew))
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

        private void BulkUpdateChildTable<T, R>(bool deleteExtraRows, R[] newRows, ref T dbRows,
            HluTableAdapter<T, R> adapter)
            where T : DataTable, new()
            where R : DataRow
        {
            if (dbRows == null) throw new ArgumentException("dbRows");
            if (newRows == null) throw new ArgumentException("newRows");
            if (adapter == null) throw new ArgumentException("adapter");

            if ((dbRows.PrimaryKey.Length != 1) || (dbRows.PrimaryKey[0].DataType != typeof(Int32)))
                throw new ArgumentException("Table must have a single column primary key of type Int32", "dbRows");

            int pkOrdinal = dbRows.PrimaryKey[0].Ordinal;

            // remove any new rows with no or duplicate data
            R[] dbRowsEnum = dbRows.AsEnumerable().Select(r => (R)r).ToArray();
            DataColumn[] cols = dbRows.Columns.Cast<DataColumn>().Where(c => c.Ordinal != pkOrdinal).ToArray();
            R[] newRowsNoDups = (from nr in newRows
                                 where cols.Count(col => !nr.IsNull(col.Ordinal)) > 0 &&
                                     dbRowsEnum.Count(dr => cols.Count(c => !dr[c.Ordinal].Equals(nr[c.Ordinal])) == 0) == 0
                                 select nr).OrderBy(r => r[pkOrdinal]).ToArray();

            // number of non-blank controls with nn-duplicate data corresponding to child table dbRows
            int numRowsNew = newRowsNoDups.Count();

            if (!deleteExtraRows && (numRowsNew == 0)) return;

            // number of child rows in database for current incid
            int numRowsDb = dbRows.Rows.Count;

            // update existing rows matching controls to rows in order of PK values (same as UI display order)
            for (int i = 0; i < numRowsNew; i++)
            {
                R newRow = newRows[i];
                int ixDbRow = (int)newRow[pkOrdinal];
                if (ixDbRow < numRowsDb)
                {
                    R dbRow = dbRowsEnum[i];
                    for (int j = 0; j < dbRows.Columns.Count; j++)
                    {
                        if ((j != pkOrdinal) && !newRow.IsNull(j))
                            dbRow[j] = newRow[j];
                    }
                    adapter.Update(dbRow);
                }
            }

            // insert any new rows
            string recordIdPropertyName = dbRows.TableName.Split('_')
                .Aggregate(new StringBuilder(), (sb, s) => sb.Append(char.ToUpper(s[0])).Append(s.Substring(1)))
                .Insert(0, "Next").Append("Id").ToString();

            PropertyInfo recordIDPropInfo = typeof(RecordIds).GetProperty(recordIdPropertyName);

            foreach (R newRow in newRowsNoDups.Where(nr => (int)nr[pkOrdinal] >= numRowsDb))
            {
                int pkValue = (int)newRow[pkOrdinal];
                if (!dbRows.Columns[pkOrdinal].AutoIncrement)
                    newRow[pkOrdinal] = (int)recordIDPropInfo.GetValue(_viewModelMain.RecIDs, null);
                adapter.Insert(newRow);
                newRow[pkOrdinal] = pkValue;
            }

            // delete extra rows
            if (deleteExtraRows)
                System.Array.ForEach(dbRowsEnum.Where((dr, index) =>
                    newRows.Count(nr => nr[pkOrdinal].Equals(index)) == 0).ToArray(),
                    new Action<R>(r => adapter.Delete(r)));

            if (numRowsNew > 0)
            {
                newRows[0].Table.AcceptChanges();
                dbRows = (T)newRows[0].Table;
            }
        }

        public void CancelBulkUpdate()
        {
            BulkUpdateResetControls();
        }

        private void BulkUpdateResetControls()
        {
            _viewModelMain.BulkUpdateMode = null;
            _viewModelMain.IncidCurrentRowIndex = 1;
            _viewModelMain.BulkUpdateMode = false;
            _viewModelMain.TabItemHistoryEnabled = true;
            _viewModelMain.RefreshAll();
            _viewModelMain.ChangeCursor(Cursors.Arrow, String.Empty);
        }
    }
}
