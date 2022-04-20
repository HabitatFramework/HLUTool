// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// Copyright © 2013 Thames Valley Environmental Records Centre
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
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using HLU.Data;
using HLU.Data.Connection;
using HLU.Data.Model;
using HLU.Properties;

namespace HLU.GISApplication
{
    static class ScratchDb
    {
        private static string _scratchMdbPath;
        private static string _scratchSelTable = "HluSelection";
        private static DbBase _scratchDb;
        private static HluGISLayer.incid_mm_polygonsDataTable _hluLayerStructure = 
            new HluGISLayer.incid_mm_polygonsDataTable();
        private static HluDataSet.incidDataTable _incidTable;
        private static HluDataSet.incid_mm_polygonsDataTable _incidMMTable;

        public static string ScratchMdbPath
        {
            get { return _scratchMdbPath; }
        }

        public static string ScratchSelectionTable
        {
            get { return _scratchSelTable; }
        }

        public static bool CreateScratchMdb(HluDataSet.incidDataTable incidTable, 
            HluDataSet.incid_mm_polygonsDataTable incidMMTable)
        {
            try
            {
                _incidTable = incidTable;
                _incidMMTable = incidMMTable;

                _scratchMdbPath = String.Empty;
                try { _scratchMdbPath = Path.GetTempPath(); }
                catch
                {
                    _scratchMdbPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    _scratchMdbPath += Path.DirectorySeparatorChar.ToString();
                }

                _scratchMdbPath += Path.GetFileNameWithoutExtension(Path.GetRandomFileName()) + ".mdb";

                OdbcCP32 odbc = new OdbcCP32();
                odbc.CreateDatabase(_scratchMdbPath);
                string connString = String.Format(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};", _scratchMdbPath);
                string defaultSchema = "";
                bool promptPwd = false;

                _scratchDb = new DbOleDb(ref connString, ref defaultSchema, ref promptPwd,
                    Properties.Resources.PasswordMaskString, Settings.Default.UseAutomaticCommandBuilders,
                    true, Settings.Default.DbIsUnicode, Settings.Default.DbUseTimeZone, 255,
                    Settings.Default.DbBinaryLength, Settings.Default.DbTimePrecision,
                    Settings.Default.DbNumericPrecision, Settings.Default.DbNumericScale);

                return true;
            }
            catch
            {
                if (File.Exists(_scratchMdbPath))
                {
                    try
                    {
                        if ((_scratchDb != null) && (_scratchDb.Connection.State != ConnectionState.Closed))
                            _scratchDb.Connection.Close();
                        File.Delete(_scratchMdbPath);
                    }
                    catch { }
                }
                return false;
            }
        }

        public static void WriteSelectionScratchTable(DataColumn[] targetColumns, DataTable idList)
        {
            try
            {
                int incidOrdinal = -1;

                if (idList.Columns.Contains(_incidTable.incidColumn.ColumnName))
                {
                    if (idList.Columns[_incidTable.incidColumn.ColumnName].DataType == _hluLayerStructure.incidColumn.DataType)
                        incidOrdinal = idList.Columns[_incidTable.incidColumn.ColumnName].Ordinal;
                    else
                        return;
                }
                else
                {
                    var q = idList.Columns.Cast<DataColumn>().Where(c => c.ColumnName.EndsWith(
                        Resources.ColumnTableNameSeparator + _incidTable.incidColumn.ColumnName) &&
                        c.DataType == _incidTable.incidColumn.DataType);
                    if (q.Count() == 1)
                        incidOrdinal = q.ElementAt(0).Ordinal;
                    else
                        return;
                }

                // incid column always has the same name as in the GIS layer structure
                if (idList.Columns[incidOrdinal].ColumnName != _hluLayerStructure.incidColumn.ColumnName)
                    idList.Columns[incidOrdinal].ColumnName = _hluLayerStructure.incidColumn.ColumnName;

                try
                {
                    _scratchDb.ExecuteNonQuery(String.Format("DROP TABLE {0}", _scratchSelTable), 
                        _scratchDb.Connection.ConnectionTimeout, CommandType.Text);
                }
                catch { }

                if (String.IsNullOrEmpty(idList.TableName)) idList.TableName = _scratchSelTable;

                if ((idList.PrimaryKey == null) || (idList.PrimaryKey.Length == 0))
                    idList.PrimaryKey = new DataColumn[] { idList.Columns[incidOrdinal] };

                if (!_scratchDb.CreateTable(idList)) return;

                DataTable scratchTable = idList.Clone();

                DataSet datasetOut = new DataSet(_scratchSelTable);
                IDbDataAdapter adapterOut = _scratchDb.CreateAdapter(scratchTable);
                adapterOut.Fill(datasetOut);
                adapterOut.TableMappings.Clear();
                adapterOut.TableMappings.Add(scratchTable.TableName, datasetOut.Tables[0].TableName);
                scratchTable = datasetOut.Tables[0];

                foreach (DataRow r in idList.Rows)
                    scratchTable.LoadDataRow(r.ItemArray, false);

                adapterOut.Update(datasetOut);
            }
            catch { }
            finally
            {
                if ((_scratchDb != null) && (_scratchDb.Connection.State != ConnectionState.Closed))
                {
                    try { _scratchDb.Connection.Close(); }
                    catch { }
                }
            }
        }

        /// <summary>
        /// Create a list of sql filter conditions using a table of selected Incids.
        /// </summary>
        /// <param name="incidSelection">A DataTable containing the selected Incids.</param>
        /// <param name="gisApp">Which GIS system is being used.</param>
        /// <returns>
        /// A list of sql filter conditions.
        /// </returns>
        public static List<SqlFilterCondition> GisWhereClause(DataTable incidSelection, GISApp gisApp, bool useIncidTable)
        {
            List<SqlFilterCondition> whereClause = new List<SqlFilterCondition>();
            SqlFilterCondition cond = new SqlFilterCondition();

            //StringBuilder incidList = new StringBuilder();

            // Split the table of selected Incids into chunks of continuous Incids so
            // that each chunk contains a continuous series of one or more Incids.
            var query = incidSelection.AsEnumerable().Select((r, index) => new
            {
                RowIndex = RecordIds.IncidNumber(r.Field<string>(0)) - index,
                Incid = r.Field<string>(0)
            }).ChunkBy(r => r.RowIndex);

            // Create a temporary list for storing some of the Incids.
            List<string> inList = new List<string>();

            // Determine which table to base the conditions on.
            DataTable condTable;
            if (useIncidTable)
                condTable = _incidTable;
            else
                condTable = _incidMMTable;

            // Loop through each chunk/series of Incids. If there are at least three
            // then it is worth processing them within ">=" and "<=" operators (as
            // this keeps the filter conditions short). If there are only one or two
            // Incids in the chunk then add them to the temporary 'inList' for
            // processing later.
            foreach (var item in query)
            {
                if (item.Count() < 3)
                {
                    if (gisApp != null)
                        inList.AddRange(item.Select(t => gisApp.QuoteValue(t.Incid)));
                    else
                        inList.AddRange(item.Select(t => t.Incid));
                }
                else
                {
                    cond = new SqlFilterCondition();
                    cond.BooleanOperator = "OR";
                    cond.OpenParentheses = "(";
                    cond.Column = _incidMMTable.incidColumn;
                    cond.Table = condTable;
                    cond.ColumnSystemType = _incidMMTable.incidColumn.DataType;
                    cond.Operator = ">=";
                    cond.Value = item.First().Incid;
                    cond.CloseParentheses = String.Empty;
                    whereClause.Add(cond);

                    cond = new SqlFilterCondition();
                    cond.BooleanOperator = "AND";
                    cond.OpenParentheses = String.Empty;
                    cond.Column = _incidMMTable.incidColumn;
                    cond.Table = condTable;
                    cond.ColumnSystemType = _incidMMTable.incidColumn.DataType;
                    cond.Operator = "<=";
                    cond.Value = item.Last().Incid;
                    cond.CloseParentheses = ")";
                    whereClause.Add(cond);
                }
            }

            // Any Incids that are not part of a continuous series of at least
            // three Incid must be queried using an "=" or "IN" sql operator.
            // So loop through all these Incids and lump them together into
            // strings of 254 Incids at a time so that each string can be used
            // in an "IN" statement.
            int i = 0;
            while (i < inList.Count)
            {
                int numElems = i < inList.Count - 254 ? 254 : inList.Count - i;
                string[] oneList = new string[numElems];
                inList.CopyTo(i, oneList, 0, numElems);

                cond = new SqlFilterCondition();
                cond.BooleanOperator = "OR";
                cond.OpenParentheses = "(";
                cond.Column = _incidMMTable.incidColumn;
                cond.Table = condTable;
                cond.ColumnSystemType = _incidMMTable.incidColumn.DataType;
                //---------------------------------------------------------------------
                // FIXOLD: 001 Improve speed of 'Select current Incid on Map'
                // Use " INCID =" in SQL statement instrad of "INCID IN ()"
                // if there is only on item in the list (as it is much quicker)
                if (inList.Count == 1)
                    cond.Operator = "=";
                else
                    cond.Operator = "IN ()";
                //---------------------------------------------------------------------
                cond.Value = String.Join(",", oneList);
                cond.CloseParentheses = ")";
                whereClause.Add(cond);

                i += numElems;
            }

            return whereClause;
        }

        /// <summary>
        /// Builds a UNION query from a list of SqlFilterCondition.
        /// All constituent SELECT statement use the same target list and FROM clause but different WHERE clauses.
        /// </summary>
        /// <param name="targetList">Target list for union query. Same for each select query.</param>
        /// <param name="fromClause">From clause for union query. Same for each select query.</param>
        /// <param name="orderByOrdinal">Ordinal of column by which to order otput.</param>
        /// <param name="IncidSelectionWhereClause">List of where clauses from which to build UNION query.
        /// Input is assumed to be 0 based.</param>
        /// <param name="db">Database against which UNION query will be run.</param>
        /// <returns></returns>
        public static string UnionQuery(string targetList, string fromClause, int[] sortOrdinals,
            List<SqlFilterCondition> IncidSelectionWhereClause, DbBase db)
        {
            //---------------------------------------------------------------------
            // CHANGED: CR43 (Sort multiple fields in exports)
            //
            // Add order by from list of sort ordinals.
            StringBuilder sql = new StringBuilder();
            sql.Append(String.Format("SELECT {0} FROM {1}{2}", targetList, fromClause, db.WhereClause(true, true, true, IncidSelectionWhereClause)))
                    .Append(sortOrdinals != null ? String.Format(" ORDER BY {0}", string.Join(", ", sortOrdinals.Select(x => x.ToString()).ToArray())) : String.Empty);
            //---------------------------------------------------------------------
            
            return sql.ToString();

            //return db.WhereClause(true, true, true, IncidSelectionWhereClause)
            //    .Aggregate(new StringBuilder(), (sb, s) => sb.Append(String.Format(
            //        "\nUNION\nSELECT {0} FROM {1}{2}", targetList, fromClause, s))).Remove(0, 7)
            //        .Append(orderByOrdinal > 0 ? String.Format(" ORDER BY {0}", orderByOrdinal + 1) : String.Empty)
            //        .ToString();
        }
        
        public static void CleanUp()
        {
            if (File.Exists(_scratchMdbPath))
            {
                try
                {
                    if ((_scratchDb != null) && (_scratchDb.Connection.State != ConnectionState.Closed))
                        _scratchDb.Connection.Close();
                    File.Delete(_scratchMdbPath);
                }
                catch { }
            }
        }
    }
}
