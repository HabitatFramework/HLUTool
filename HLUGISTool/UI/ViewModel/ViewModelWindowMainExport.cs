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
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using HLU.Data.Connection;
using HLU.Data.Model;
using HLU.GISApplication;
using HLU.Properties;
using HLU.UI.View;

namespace HLU.UI.ViewModel
{
    class ViewModelWindowMainExport
    {
        ViewModelWindowMain _viewModelMain;
        private WindowExport _windowExport;
        private ViewModelExport _viewModelExport;

        public ViewModelWindowMainExport(ViewModelWindowMain viewModelMain)
        {
            _viewModelMain = viewModelMain;
        }

        public void InitiateExport()
        {
            _windowExport = new WindowExport();
            _windowExport.Owner = App.Current.MainWindow;
            _windowExport.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            _viewModelMain.HluTableAdapterManager.exportsTableAdapter.ClearBeforeFill = true;
            _viewModelMain.HluTableAdapterManager.exportsTableAdapter.Fill(_viewModelMain.HluDataset.exports,
                String.Format("EXISTS (SELECT {0}.{1} FROM {0} WHERE {0}.{1} = {2}.{3})",
                _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.exports_fields.TableName),
                _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.exports_fields.export_idColumn.ColumnName),
                _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.exports.TableName),
                _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.exports.export_idColumn.ColumnName)));

            if (_viewModelMain.HluDataset.exports.Count == 0)
            {
                MessageBox.Show("Cannot export: there are no export formats defined.",
                    "HLU: Export", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            _viewModelExport = new ViewModelExport(_viewModelMain.GisSelection == null ? 0 :
                _viewModelMain.GisSelection.Rows.Count, _viewModelMain.GISApplication.HluLayerName,
                _viewModelMain.GISApplication.ApplicationType, _viewModelMain.HluDataset.exports);
            _viewModelExport.DisplayName = "Export";
            _viewModelExport.RequestClose += new ViewModelExport.RequestCloseEventHandler(_viewModelExport_RequestClose);

            _windowExport.DataContext = _viewModelExport;

            _windowExport.ShowDialog();
        }

        private void _viewModelExport_RequestClose(int exportID, bool exportDescriptions, bool selectedOnly)
        {
            _viewModelExport.RequestClose -= _viewModelExport_RequestClose;
            _windowExport.Close();

            if (exportID != -1)
            {
                DispatcherHelper.DoEvents();
                Export(exportID, exportDescriptions, selectedOnly);
            }
        }

        private void Export(int userExportId, bool exportDescriptions, bool selectedOnly)
        {
            string tempPath = null;

            try
            {
                string tableAlias = GetTableAlias();
                if (tableAlias == null)
                    throw new Exception("Failed to find a table alias that does not match a table name in the HLU dataset.");

                // refresh export fields from DB
                _viewModelMain.HluTableAdapterManager.exportsFieldsTableAdapter.ClearBeforeFill = true;
                _viewModelMain.HluTableAdapterManager.exportsFieldsTableAdapter.Fill(
                    _viewModelMain.HluDataset.exports_fields, String.Format("{0} = {1} ORDER BY {2}, {3}",
                    _viewModelMain.DataBase.QuoteIdentifier(
                    _viewModelMain.HluDataset.exports_fields.export_idColumn.ColumnName), userExportId, 
                    _viewModelMain.DataBase.QuoteIdentifier(
                    _viewModelMain.HluDataset.exports_fields.table_nameColumn.ColumnName),
                    _viewModelMain.DataBase.QuoteIdentifier(
                    _viewModelMain.HluDataset.exports_fields.field_ordinalColumn.ColumnName)));

                if (_viewModelMain.HluDataset.exports_fields.Count == 0)
                    throw new Exception(String.Format("No export fields are defined for format '{0}'.",
                        _viewModelMain.HluDataset.exports.FindByexport_id(userExportId).export_name));

                DataTable exportTable;
                List<int> fieldCountList;
                StringBuilder targetList;
                StringBuilder fromClause;
                int incidOrdinal;
                ExportJoins(exportDescriptions, tableAlias, out exportTable,
                    out fieldCountList, out targetList, out fromClause, out incidOrdinal);

                string whereClause = String.Empty;

                List<List<SqlFilterCondition>> exportFilter = null;
                if (!_viewModelExport.SelectedOnly)
                {
                    SqlFilterCondition cond = new SqlFilterCondition("AND",
                        _viewModelMain.IncidTable, _viewModelMain.IncidTable.incidColumn, null);
                    cond.Operator = "IS NOT NULL";
                    exportFilter = new List<List<SqlFilterCondition>>(new List<SqlFilterCondition>[] { 
                        new List<SqlFilterCondition>(new SqlFilterCondition[] { cond }) });
                }
                else
                {
                    if ((_viewModelMain.IncidSelectionWhereClause == null) &&
                        (_viewModelMain.GisSelection != null) && (_viewModelMain.GisSelection.Rows.Count > 0))
                        _viewModelMain.IncidSelectionWhereClause = ViewModelWindowMainHelpers.GisSelectionToWhereClause(
                            _viewModelMain.GisSelection.Select(), _viewModelMain.GisIDColumnOrdinals,
                            ViewModelWindowMain.IncidPageSize, _viewModelMain.HluDataset.incid);
                    exportFilter = _viewModelMain.IncidSelectionWhereClause;
                }

                string sql = ScratchDb.UnionQuery(targetList.ToString(), fromClause.ToString(),
                    incidOrdinal, exportFilter, _viewModelMain.DataBase);

                _viewModelMain.ChangeCursor(Cursors.Wait, "Exporting ...");

                int exportRowCount;
                tempPath = ExportMdb(sql, exportTable, incidOrdinal, fieldCountList, out exportRowCount);

                if (!String.IsNullOrEmpty(tempPath) && (exportRowCount > 0))
                    _viewModelMain.GISApplication.Export(tempPath, exportTable.TableName, exportRowCount);
                else
                    throw new Exception("Export query did not retrieve any rows");

                _viewModelMain.IncidSelection = null;
                _viewModelMain.GisSelection = null;
                _viewModelMain.OnPropertyChanged("IsFiltered");
                _viewModelMain.OnPropertyChanged("StatusIncid");
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Export failed. The error message was:\n\n{0}.",
                    ex.Message), "HLU: Export", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (!String.IsNullOrEmpty(tempPath))
                {
                    string[] tempFiles = Directory.GetFiles(Path.GetDirectoryName(tempPath),
                        Path.GetFileNameWithoutExtension(tempPath) + ".*");
                    foreach (string fName in tempFiles)
                    {
                        try { File.Delete(fName); }
                        catch { _viewModelMain.ExportMdbs.Add(fName); }
                    }
                }
                _viewModelMain.ChangeCursor(Cursors.Arrow, null);
            }
        }

        private void ExportJoins(bool exportDescriptions, string tableAlias, out DataTable exportTable,
            out List<int> fieldCountList, out StringBuilder targetList, out StringBuilder fromClause,
            out int incidOrdinal)
        {
            exportTable = new DataTable("HluExport");
            List<int> lutDescrColOrdinals = new List<int>();
            fieldCountList = new List<int>();
            targetList = new StringBuilder();
            List<string> fromList = new List<string>();
            List<string> leftJoined = new List<string>();
            fromClause = new StringBuilder();

            int tableAliasNum = 1;
            bool firstJoin = true;
            incidOrdinal = -1;

            foreach (HluDataSet.exports_fieldsRow r in 
                _viewModelMain.HluDataset.exports_fields.OrderBy(r => r.field_ordinal))
            {
                bool multipleFields = false;
                if (!r.IsNull(_viewModelMain.HluDataset.exports_fields.fields_countColumn))
                {
                    fieldCountList.Add(r.fields_count);
                    multipleFields = true;
                }
                else
                {
                    if ((r.table_name == _viewModelMain.HluDataset.incid.TableName) &&
                        (r.column_name == _viewModelMain.HluDataset.incid.incidColumn.ColumnName))
                    {
                        incidOrdinal = fieldCountList.Count;
                    }
                    fieldCountList.Add(0);
                }

                string currTable = _viewModelMain.DataBase.QualifyTableName(r.table_name);
                if (!fromList.Contains(currTable))
                {
                    fromList.Add(currTable);

                    var incidRelation = _viewModelMain.HluDataset.incid.ChildRelations.Cast<DataRelation>()
                        .Where(dr => dr.ChildTable.TableName == r.table_name);

                    if (incidRelation.Count() == 0)
                    {
                        fromClause.Append(currTable);
                    }
                    else
                    {
                        DataRelation incidRel = incidRelation.ElementAt(0);
                        if (firstJoin)
                            firstJoin = false;
                        else
                            fromClause.Insert(0, "(").Append(")");
                        fromClause.Append(RelationJoinClause("LEFT", currTable, true,
                            _viewModelMain.DataBase.QuoteIdentifier(
                            incidRel.ParentTable.TableName), incidRel, fromList));
                        leftJoined.Add(currTable);
                    }
                }

                var relations = exportDescriptions ? _viewModelMain.HluDataRelations.Where(rel =>
                    rel.ChildTable.TableName == r.table_name && rel.ChildColumns
                    .Count(ch => ch.ColumnName == r.column_name) == 1) : new DataRelation[0];

                switch (relations.Count())
                {
                    case 0:
                        targetList.Append(String.Format(",{0}.{1} AS {2}", currTable,
                            _viewModelMain.DataBase.QuoteIdentifier(r.column_name), r.field_name));
                        AddExportColumn(multipleFields ? r.fields_count : 0, r.field_name,
                            _viewModelMain.HluDataset.Tables[r.table_name].Columns[r.column_name], ref exportTable,
                            ref lutDescrColOrdinals);
                        break;
                    case 1:
                        DataRelation lutRelation = relations.ElementAt(0);
                        string parentTable = _viewModelMain.DataBase.QualifyTableName(lutRelation.ParentTable.TableName);

                        string parentTableAlias = tableAlias + tableAliasNum++;
                        fromList.Add(parentTable);

                        if (lutRelation.ParentTable.Columns.Contains(ViewModelWindowMain.LutDescriptionFieldName))
                        {
                            targetList.Append(String.Format(",{0}.{1} AS {2}", parentTableAlias,
                                _viewModelMain.DataBase.QuoteIdentifier(
                                ViewModelWindowMain.LutDescriptionFieldName), r.field_name));
                            AddExportColumn(multipleFields ? r.fields_count : 0, r.field_name,
                                lutRelation.ParentTable.Columns[ViewModelWindowMain.LutDescriptionFieldName],
                                ref exportTable, ref lutDescrColOrdinals);
                        }
                        else if (lutRelation.ParentTable.Columns.Count >= ViewModelWindowMain.LutDescriptionFieldOrdinal)
                        {
                            targetList.Append(String.Format(",{0}.{1} AS {2}", 
                                parentTableAlias, _viewModelMain.DataBase.QuoteIdentifier(
                                lutRelation.ParentTable.Columns[ViewModelWindowMain.LutDescriptionFieldOrdinal - 1].ColumnName),
                                r.field_name));
                            AddExportColumn(multipleFields ? r.fields_count : 0, r.field_name,
                                lutRelation.ParentTable.Columns[ViewModelWindowMain.LutDescriptionFieldOrdinal - 1],
                                ref exportTable, ref lutDescrColOrdinals);
                        }
                        else
                        {
                            continue;
                        }

                        string joinType;
                        if (leftJoined.Contains(currTable))
                        {
                            joinType = "LEFT";
                            leftJoined.Add(parentTableAlias);
                        }
                        else
                        {
                            joinType = "INNER";
                        }

                        if (firstJoin)
                            firstJoin = false;
                        else
                            fromClause.Insert(0, "(").Append(")");

                        fromClause.Append(RelationJoinClause(joinType, currTable,
                            false, parentTableAlias, lutRelation, fromList));

                        break;
                }
            }

            exportTable.PrimaryKey = new DataColumn[] { exportTable.Columns[incidOrdinal] };
            if (targetList.Length > 1) targetList.Remove(0, 1);
        }

        private string ExportMdb(string sql, DataTable exportTable, int incidOrdinal,
            List<int> fieldCountList, out int exportRowCount)
        {
            exportRowCount = -1;
            DbOleDb dbOut = null;
            string tempPath = String.Empty;
            try { tempPath = Path.GetTempPath(); }
            catch { tempPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData); }
            tempPath = Path.Combine(tempPath, Path.GetFileNameWithoutExtension(Path.GetRandomFileName()) + ".mdb");

            try
            {
                if (File.Exists(tempPath)) File.Delete(tempPath);
                OdbcCP32 odbc = new OdbcCP32();
                odbc.CreateDatabase(tempPath);
                string connString = String.Format(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};", tempPath);
                string defaultSchema = "";
                bool promptPwd = false;
                dbOut = new DbOleDb(ref connString, ref defaultSchema, ref promptPwd,
                    Properties.Resources.PasswordMaskString, Settings.Default.UseAutomaticCommandBuilders,
                    true, Settings.Default.DbIsUnicode, Settings.Default.DbUseTimeZone, 255,
                    Settings.Default.DbBinaryLength, Settings.Default.DbTimePrecision,
                    Settings.Default.DbNumericPrecision, Settings.Default.DbNumericScale);
                dbOut.CreateTable(exportTable);
                DataSet datasetOut = new DataSet("Export");

                IDbDataAdapter adapterOut = dbOut.CreateAdapter(exportTable);
                adapterOut.Fill(datasetOut);
                int[] pkOrdinals = exportTable.PrimaryKey.Select(c => c.Ordinal).ToArray();
                exportTable.PrimaryKey = pkOrdinals.Select(o => exportTable.Columns[o]).ToArray();
                adapterOut.TableMappings.Clear();
                adapterOut.TableMappings.Add(exportTable.TableName, datasetOut.Tables[0].TableName);
                exportTable = datasetOut.Tables[0];

                DataRow exportRow = null;
                bool rowAdded = false;

                using (IDataReader reader = _viewModelMain.DataBase.ExecuteReader(sql, 
                    _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text))
                {
                    int runTotal = 0;
                    int[][] fieldMapTemplate = fieldCountList.Select((i, index) =>
                        new int[] { index, index + runTotal, i, (runTotal += i - (i > 0 ? 1 : 0)) })
                        .Select(e => new int[] { e[0], e[1], e[1] + e[2] }).ToArray();

                    int[] dupsAllowed = (from e in exportTable.Columns.Cast<DataColumn>()
                                         let q = from c in _viewModelMain.HluDataset.incid_sources.Columns.Cast<DataColumn>()
                                                 where !Regex.IsMatch(c.ColumnName,
                                                    @"(\Aincid\z|_(importance|id)\z)", RegexOptions.IgnoreCase)
                                                 select c.ColumnName
                                         where q.Count(n => Regex.IsMatch(e.ColumnName,
                                             n + @"(_[0-9]+)*\z", RegexOptions.IgnoreCase)) == 1
                                         select e.Ordinal).ToArray();

                    int[][] fieldMap = new int[fieldMapTemplate.Length][];
                    string currIncid = String.Empty;
                    string prevIncid = String.Empty;

                    while (reader.Read())
                    {
                        currIncid = reader.GetString(incidOrdinal);
                        if (currIncid != prevIncid)
                        {
                            prevIncid = currIncid;
                            fieldMap = fieldMapTemplate.Select(a => new int[] { a[0], a[1], a[2] }).ToArray();
                            if (exportRow != null)
                            {
                                exportTable.Rows.Add(exportRow);
                                rowAdded = true;
                            }
                            exportRow = exportTable.NewRow();
                            rowAdded = false;
                            for (int i = 0; i < fieldMap.GetLength(0); i++)
                            {
                                object item = reader.GetValue(fieldMap[i][0]);
                                if (item != DBNull.Value)
                                    exportRow[fieldMap[i][1]] = reader.GetValue(fieldMap[i][0]);
                                fieldMap[i][1]++;
                            }
                        }
                        else
                        {
                            for (int i = 0; i < fieldMap.GetLength(0); i++)
                            {
                                if (fieldMap[i][1] < fieldMap[i][2])
                                {
                                    object item = reader.GetValue(fieldMap[i][0]);
                                    if ((item != DBNull.Value) && (!item.Equals(exportRow[fieldMap[i][1] - 1]) ||
                                        (Array.IndexOf(dupsAllowed, fieldMap[i][1]) != -1)))
                                        exportRow[fieldMap[i][1]++] = item;
                                }
                            }
                        }
                    }
                }

                if (!rowAdded && (exportRow != null)) exportTable.Rows.Add(exportRow);

                exportRowCount = adapterOut.Update(datasetOut);

                return exportRowCount != -1 ? tempPath : null;
            }
            catch
            {
                if (File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); }
                    catch { _viewModelMain.ExportMdbs.Add(tempPath); }
                }
                return null;
            }
            finally
            {
                if ((dbOut != null) && (dbOut.Connection.State != ConnectionState.Closed))
                {
                    try { dbOut.Connection.Close(); }
                    catch { }
                }
            }
        }

        private void AddExportColumn(int numFields, string columnName, DataColumn templateColumn,
            ref DataTable exportTable, ref List<int> lutDescrColOrdinals)
        {
            if (numFields > 0)
            {
                for (int i = 1; i <= numFields; i++)
                {
                    DataColumn c = new DataColumn(String.Format("{0}_{1}", columnName, i), 
                        templateColumn.DataType);
                    c.MaxLength = templateColumn.MaxLength;
                    exportTable.Columns.Add(c);
                }
            }
            else
            {
                DataColumn c = new DataColumn(columnName, templateColumn.DataType);
                c.MaxLength = templateColumn.MaxLength;
                exportTable.Columns.Add(c);
            }
        }

        private string RelationJoinClause(string joinType, string currTable, bool parentLeft,
            string parentTableAlias, DataRelation rel, List<string> fromList)
        {
            StringBuilder joinClausePart = new StringBuilder();

            for (int i = 0; i < rel.ParentColumns.Length; i++)
            {
                joinClausePart.Append(String.Format(" AND {0}.{2} = {1}.{3}", parentTableAlias, 
                    currTable, _viewModelMain.DataBase.QuoteIdentifier(rel.ParentColumns[i].ColumnName),
                    _viewModelMain.DataBase.QuoteIdentifier(rel.ChildColumns[i].ColumnName)));
            }

            if (parentTableAlias == _viewModelMain.DataBase.QuoteIdentifier(rel.ParentTable.TableName))
                parentTableAlias = String.Empty;
            else
                parentTableAlias = " " + parentTableAlias;

            string leftTable = String.Empty;
            string rightTable = string.Empty;
            if (parentLeft)
            {
                leftTable = _viewModelMain.DataBase.QuoteIdentifier(rel.ParentTable.TableName) + parentTableAlias;
                rightTable = currTable;
            }
            else
            {
                leftTable = currTable;
                rightTable = _viewModelMain.DataBase.QuoteIdentifier(rel.ParentTable.TableName) + parentTableAlias;
            }

            if (!fromList.Contains(currTable))
                return joinClausePart.Remove(0, 5).Insert(0, String.Format(" {0} {1} JOIN {2} ON ",
                    leftTable, joinType, rightTable)).ToString();
            else
                return joinClausePart.Remove(0, 5).Insert(0, String.Format(" {0} JOIN {1} ON ",
                    joinType, rightTable)).ToString();
        }

        private string GetTableAlias()
        {
            for (int i = 1; i < 5; i++)
            {
                for (int j = 122; j > 96; j--)
                {
                    char[] testCharArray = new char[i];
                    for (int k = 0; k < i; k++)
                        testCharArray[k] = (char)j;
                    string testString = new string(testCharArray);
                    if (_viewModelMain.HluDataset.Tables.Cast<DataTable>().Count(t => Regex.IsMatch(t.TableName,
                        testString + "[0-9]+", RegexOptions.IgnoreCase)) == 0)
                    {
                        return testString;
                    }
                }
            }
            return null;
        }
    }
}
