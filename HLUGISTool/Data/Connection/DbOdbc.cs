// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
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
using System.Data.Common;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Windows;

namespace HLU.Data.Connection
{
    class DbOdbc : DbBase
    {
        #region Private Members

        private string _errorMessage;
        private OdbcConnectionStringBuilder _connStrBuilder;
        private OdbcConnection _connection;
        private OdbcCommand _command;
        private OdbcDataAdapter _adapter;
        private OdbcCommandBuilder _commandBuilder;
        private OdbcTransaction _transaction;
        private Dictionary<Type, OdbcDataAdapter> _adaptersDic = new Dictionary<Type, OdbcDataAdapter>();

        private HLU.UI.View.Connection.ViewConnectOdbc _connWindow;
        private HLU.UI.ViewModel.ViewModelConnectOdbc _connViewModel;

        private Backends _backend;
        private string _quotePrefix;
        private string _quoteSuffix;
        private string _stringLiteralDelimiter;
        private string _dateLiteralPrefix;
        private string _dateLiteralSuffix;
        private string _wildcardSingleMatch;
        private string _wildcardManyMatch;
        #endregion

        #region Constructor

        public DbOdbc(ref string connString, ref string defaultSchema, ref bool promptPwd,
            string pwdMask, bool useCommandBuilder, bool useColumnNames, bool isUnicode, 
            bool useTimeZone, uint textLength, uint binaryLength, uint timePrecision, 
            uint numericPrecision, uint numericScale)
            : base(ref connString, ref defaultSchema, ref promptPwd, pwdMask, 
            useCommandBuilder, useColumnNames, isUnicode, useTimeZone, textLength, 
            binaryLength, timePrecision, numericPrecision, numericScale)
        {
            if (String.IsNullOrEmpty(ConnectionString)) throw (new Exception("No connection string"));

            try
            {
                Login(_backend == Backends.Oracle ? "User ID" : "User name", ConnectionString, 
                    ref promptPwd, ref _connStrBuilder, ref _connection);

                PopulateTypeMaps(IsUnicode, UseTimeZone, TextLength, BinaryLength,
                    TimePrecision, NumericPrecision, NumericScale);
                SetDefaults();

                _command = _connection.CreateCommand();
                _adapter = new OdbcDataAdapter(_command);
                _commandBuilder = new OdbcCommandBuilder(_adapter);

                connString = MaskPassword(_connStrBuilder, pwdMask);
                defaultSchema = DefaultSchema;
            }
            catch { throw; }
        }
        
        #endregion

        #region DbBase Members

        #region Public Static

        public static Backends GetBackend(OdbcConnection cn)
        {
            ConnectionState previousConnectionState = cn.State;
            if (String.IsNullOrEmpty(cn.Driver) && 
                (previousConnectionState != ConnectionState.Open)) cn.Open();
            
            string driver = cn.Driver.ToLower();

            if ((cn.State == ConnectionState.Open) && 
                (previousConnectionState != ConnectionState.Open)) cn.Close();

            if (driver.StartsWith("odbcjt32"))
                return Backends.Access;
            else if ((driver.StartsWith("sqlncli")) || (driver.StartsWith("sqlsrv")))
                return Backends.SqlServer;
            else if ((driver.StartsWith("sqora")) || (driver.StartsWith("msorcl")))
                return Backends.Oracle;
            else if (driver.StartsWith("psql"))
                return Backends.PostgreSql;
            else if (driver.StartsWith("db2"))
                return Backends.DB2;
            else
                return Backends.UndeterminedOdbc;
        }

        public static Backends GetBackend(OdbcConnectionStringBuilder connStrBuilder)
        {
            if ((connStrBuilder == null) || String.IsNullOrEmpty(connStrBuilder.ConnectionString))
                return Backends.UndeterminedOdbc;

            try
            {
                OdbcConnection cn = new OdbcConnection(connStrBuilder.ConnectionString);
                return GetBackend(cn);
            }
            catch { return Backends.UndeterminedOdbc; }
        }

        public static Backends GetBackend(string connString)
        {
            if (String.IsNullOrEmpty(connString)) return Backends.UndeterminedOdbc;

            try
            {
                OdbcConnection cn = new OdbcConnection(connString);
                return GetBackend(cn);
            }
            catch { return Backends.UndeterminedOdbc; }
        }

        #endregion

        #region Public Members

        public override Backends Backend { get { return _backend; } }

        public override bool ContainsDataSet(DataSet ds, out string errorMessage)
        {
            errorMessage = null;

            try
            {
                DataTable schemaTable = GetSchema("Columns", 
                    _restrictionNameSchema, DefaultSchema, _connection, _transaction);
                var dbSchema = schemaTable.AsEnumerable();

                StringBuilder messageText = new StringBuilder();

                foreach (DataTable t in ds.Tables)
                {
                    dbSchema.Select(r => r.Field<string>("TABLE_NAME") == t.TableName);
                    var dbSchemaCols = from r in dbSchema
                                       let tableName = r.Field<string>("TABLE_NAME")
                                       where tableName == t.TableName
                                       select new
                                       {
                                           TableName = tableName,
                                           ColumnName = r.Field<string>("COLUMN_NAME"),
                                           DataType = r.Field<string>("TYPE_NAME") //<int>("DATA_TYPE")
                                       };

                    if (dbSchemaCols.Count() == 0)
                    {
                        messageText.Append(String.Format("\n\nMissing table: {0}", QuoteIdentifier(t.TableName)));
                    }
                    else
                    {
                        string[] checkColumns = (from dsCol in t.Columns.Cast<DataColumn>()
                                                 let dbCols = from dbCol in dbSchemaCols
                                                              where dbCol.ColumnName == dsCol.ColumnName &&
                                                              DbToSystemType(SQLCodeToSQLType(dbCol.DataType)) == dsCol.DataType
                                                              select dbCol
                                                 where dbCols.Count() == 0
                                                 select QuoteIdentifier(dsCol.ColumnName) + " (" +
                                                 ((OdbcType)SystemToDbType(dsCol.DataType) + ")").ToString()).ToArray();
                        if (checkColumns.Length > 0) messageText.Append(String.Format("\n\nTable: {0}\nColumns: {1}",
                            QuoteIdentifier(t.TableName), String.Join(", ", checkColumns)));
                    }
                }

                if (messageText.Length == 0)
                {
                    return true;
                }
                else
                {
                    errorMessage = String.Format("Connection does not point to a valid HLU database." +
                            "\nBad schema objects: {0}", messageText);
                    return false;
                }
            }
            catch { }

            return false;
        }
        public override IDbConnection Connection { get { return _connection; } }

        public override DbConnectionStringBuilder ConnectionStringBuilder { get { return _connStrBuilder; } }

        public override IDbTransaction Transaction
        {
            get { return _transaction; }
        }

        public override IDbCommand CreateCommand()
        {
            if (_connection != null)
                return _connection.CreateCommand();
            else
                return new OdbcCommand();
        }

        public override IDbDataAdapter CreateAdapter()
        {
            return new OdbcDataAdapter();
        }

        public override IDbDataAdapter CreateAdapter<T>(T table)
        {
            if (table == null) table = new T();

            OdbcDataAdapter adapter;

            if (!_adaptersDic.TryGetValue(typeof(T), out adapter))
            {
                adapter = new OdbcDataAdapter();

                DataColumn[] pk = table.PrimaryKey;
                if ((pk == null) || (pk.Length == 0)) return null;

                DataTableMapping tableMapping = new DataTableMapping();
                tableMapping.SourceTable = table.TableName;
                tableMapping.DataSetTable = table.TableName;

                List<OdbcParameter> deleteParams = new List<OdbcParameter>();
                List<OdbcParameter> insertParams = new List<OdbcParameter>();
                List<OdbcParameter> updateParams = new List<OdbcParameter>();
                List<OdbcParameter> updateParamsOrig = new List<OdbcParameter>();

                StringBuilder sbTargetList = new StringBuilder();
                StringBuilder sbInsValues = new StringBuilder();
                StringBuilder sbUpdSetList = new StringBuilder();
                StringBuilder sbWhereDel = new StringBuilder();
                StringBuilder sbWhereUpd = new StringBuilder();
                StringBuilder sbWherePkUpd = new StringBuilder();
                StringBuilder sbWherePkIns = new StringBuilder();

                string tableName = QualifyTableName(table.TableName);

                int isNullTypeInt;
                _typeMapSystemToSQL.TryGetValue(typeof(int), out isNullTypeInt);
                OdbcType isNullType = (OdbcType)isNullTypeInt;

                DataColumn c;
                string delOrigParamName;
                string insColParamName;
                string updColParamName;
                string updOrigParamName;
                string delIsNullParamName;
                string updIsNullParamName;
                int columnCount = table.Columns.Count;
                int nullParamCount = 0;
                string delAddString;
                string updAddString;

                for (int i = 0; i < table.Columns.Count; i++)
                {
                    c = table.Columns[i];
                    tableMapping.ColumnMappings.Add(c.ColumnName, c.ColumnName);
                    string colName = QuoteIdentifier(c.ColumnName);

                    int colType;
                    if (!_typeMapSystemToSQL.TryGetValue(c.DataType, out colType)) continue;

                    if (c.AllowDBNull)
                    {
                        delIsNullParamName = ParameterName(_parameterPrefixNull, c.ColumnName,
                            deleteParams.Count + _startParamNo);
                        deleteParams.Add(CreateParameter(delIsNullParamName, isNullType,
                            ParameterDirection.Input, c.ColumnName, DataRowVersion.Original, true));

                        updIsNullParamName = ParameterName(_parameterPrefixNull, c.ColumnName,
                            i + columnCount + nullParamCount + _startParamNo);
                        updateParamsOrig.Add(CreateParameter(updIsNullParamName, isNullType,
                            ParameterDirection.Input, c.ColumnName, DataRowVersion.Original, true));

                        delAddString = String.Format(" AND ((({0} = 1) AND ({1} IS NULL)) OR ({1} = ",
                            ParameterMarker(delIsNullParamName), colName) + "{0}))";
                        updAddString = String.Format(" AND ((({0} = 1) AND ({1} IS NULL)) OR ({1} = ",
                            ParameterMarker(updIsNullParamName), colName) + "{0}))";

                        nullParamCount++;
                    }
                    else
                    {
                        delAddString = String.Format(" AND ({0} = ", colName) + "{0})";
                        updAddString = String.Format(" AND ({0} = ", colName) + "{0})";
                    }

                    delOrigParamName = ParameterName(_parameterPrefixOrig, c.ColumnName, deleteParams.Count + _startParamNo);
                    deleteParams.Add(CreateParameter(delOrigParamName,
                        (OdbcType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Original, false));

                    insColParamName = ParameterName(_parameterPrefixCurr, c.ColumnName, i + _startParamNo);
                    insertParams.Add(CreateParameter(insColParamName,
                        (OdbcType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Current, false));

                    updColParamName = ParameterName(_parameterPrefixCurr, c.ColumnName, i + _startParamNo);
                    updateParams.Add(CreateParameter(updColParamName,
                        (OdbcType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Current, false));

                    updOrigParamName = ParameterName(_parameterPrefixOrig, c.ColumnName, i + _startParamNo + columnCount + nullParamCount);
                    updateParamsOrig.Add(CreateParameter(updOrigParamName,
                        (OdbcType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Original, false));

                    sbTargetList.Append(", ").Append(colName);
                    sbUpdSetList.Append(String.Format(", {0} = {1}", colName, ParameterMarker(updColParamName)));
                    sbInsValues.Append(", ").Append(ParameterMarker(insColParamName));

                    sbWhereDel.Append(String.Format(delAddString, ParameterMarker(delOrigParamName)));
                    sbWhereUpd.Append(String.Format(updAddString, ParameterMarker(updOrigParamName)));

                    if (Array.IndexOf(pk, c) != -1)
                    {
                        sbWherePkUpd.Append(String.Format(" AND ({0} = {1})", colName, ParameterMarker(updColParamName)));
                        sbWherePkIns.Append(String.Format(" AND ({0} = {1})", colName, ParameterMarker(insColParamName)));
                    }
                }
                updateParams.AddRange(updateParamsOrig);
                sbTargetList.Remove(0, 2);
                sbInsValues.Remove(0, 2);
                sbUpdSetList.Remove(0, 2);
                sbWhereDel.Remove(0, 5);
                sbWhereUpd.Remove(0, 5);
                sbWherePkUpd.Remove(0, 5);
                sbWherePkIns.Remove(0, 5);

                adapter.TableMappings.Add(tableMapping);

                adapter.SelectCommand = new OdbcCommand();
                adapter.SelectCommand.CommandType = CommandType.Text;
                adapter.SelectCommand.Connection = _connection;
                adapter.SelectCommand.CommandText = String.Format("SELECT {0} FROM {1}", sbTargetList, tableName);

                if (!_useCommandBuilder)
                {
                    adapter.DeleteCommand = new OdbcCommand();
                    adapter.DeleteCommand.CommandType = CommandType.Text;
                    adapter.DeleteCommand.Connection = _connection;
                    adapter.DeleteCommand.CommandText = String.Format("DELETE FROM {0} WHERE {1}", tableName, sbWhereDel);
                    adapter.DeleteCommand.Parameters.AddRange(deleteParams.ToArray());

                    adapter.UpdateCommand = new OdbcCommand();
                    adapter.UpdateCommand.Connection = _connection;
                    adapter.UpdateCommand.CommandType = CommandType.Text;
                    adapter.UpdateCommand.CommandText =
                        String.Format("UPDATE {0} SET {1} WHERE {2}", tableName, sbUpdSetList, sbWhereUpd);
                    adapter.UpdateCommand.Parameters.AddRange(updateParams.ToArray());

                    adapter.InsertCommand = new OdbcCommand();
                    adapter.InsertCommand.CommandType = CommandType.Text;
                    adapter.InsertCommand.Connection = _connection;
                    adapter.InsertCommand.CommandText = String.Format("INSERT INTO {0} ({1}) VALUES ({2})",
                        tableName, sbTargetList, sbInsValues);
                    adapter.InsertCommand.Parameters.AddRange(insertParams.ToArray());
                }
                else
                {
                    OdbcCommandBuilder cmdBuilder = new OdbcCommandBuilder(adapter);
                    adapter.DeleteCommand = cmdBuilder.GetDeleteCommand(_useColumnNames);
                    adapter.UpdateCommand = cmdBuilder.GetUpdateCommand(_useColumnNames);
                    adapter.InsertCommand = cmdBuilder.GetInsertCommand(_useColumnNames);
                }

                if (_backend != Backends.Access)
                {
                    adapter.UpdateCommand.CommandText += ";\r\n" +
                        String.Format("SELECT {0} FROM {1} WHERE {2}", sbTargetList, tableName, sbWherePkUpd);
                    adapter.InsertCommand.CommandText += ";\r\n" +
                        String.Format("SELECT {0} FROM {1} WHERE {2}", sbTargetList, tableName, sbWherePkIns);
                }

                if (typeof(T) != typeof(DataTable))
                    _adaptersDic.Add(typeof(T), adapter);
            }

            return adapter;
        }

        private OdbcParameter CreateParameter(string name, OdbcType type, ParameterDirection direction,
            string srcColumn, DataRowVersion srcVersion, bool nullMapping)
        {
            OdbcParameter param = new OdbcParameter(name, type);
            param.Direction = direction;
            param.SourceColumn = srcColumn;
            param.SourceVersion = srcVersion;
            param.SourceColumnNullMapping = nullMapping;
            return param;
        }

        private OdbcParameter CreateParameter(string name, object value, ParameterDirection direction,
            string srcColumn, DataRowVersion srcVersion, bool nullMapping)
        {
            OdbcParameter param = new OdbcParameter(name, value);
            param.Direction = direction;
            param.SourceColumn = srcColumn;
            param.SourceVersion = srcVersion;
            param.SourceColumnNullMapping = nullMapping;
            return param;
        }

        protected override string ParameterName(string prefix, string columnName, int paramNo)
        {
            if (_useColumnNames)
                return ParameterPrefix + prefix + columnName;
            else
                return String.Format("{0}p{1}", ParameterPrefix, paramNo);
        }

        protected override string ParameterMarker(string parameterName)
        {
            return "?";
        }

        public override bool FillSchema<T>(SchemaType schemaType, string sql, ref T table)
        {
            if (String.IsNullOrEmpty(sql)) return false;
            ConnectionState previousConnectionState = _connection.State;
            try
            {
                _errorMessage = String.Empty;
                if (table == null) table = new T();
                if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();
                OdbcDataAdapter adapter = UpdateAdapter(table);
                if (adapter != null)
                {
                    if (_transaction != null)
                        adapter.SelectCommand.Transaction = _transaction;
                    adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;
                    adapter.FillSchema(table, schemaType);
                }
                else
                {
                    _command.CommandText = sql;
                    _command.CommandType = CommandType.Text;
                    if (_transaction != null) _command.Transaction = _transaction;
                    _adapter = new OdbcDataAdapter(_command);
                    _adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;
                    _adapter.FillSchema(table, schemaType);
                }
                return true;
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return false;
            }
            finally { if (previousConnectionState == ConnectionState.Closed) _connection.Close(); }
        }

        public override int FillTable<T>(string sql, ref T table)
        {
            if (String.IsNullOrEmpty(sql)) return 0;
            ConnectionState previousConnectionState = _connection.State;
            try
            {
                _errorMessage = String.Empty;
                if (table == null) table = new T();
                _command.CommandText = sql;
                _command.CommandType = CommandType.Text;
                if (_transaction != null) _command.Transaction = _transaction;
                if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();
                OdbcDataAdapter adapter = UpdateAdapter(table);
                if (adapter != null)
                {
                    return adapter.Fill(table);
                }
                else
                {
                    _adapter = new OdbcDataAdapter(_command);
                    return _adapter.Fill(table);
                }
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return -1;
            }
            finally { if (previousConnectionState == ConnectionState.Closed) _connection.Close(); }
        }

        public override bool BeginTransaction(bool commitPrevious, System.Data.IsolationLevel isolationLevel)
        {
            try
            {
                _errorMessage = String.Empty;
                if (_transaction != null)
                {
                    if (commitPrevious)
                        _transaction.Commit();
                    else
                        _transaction.Rollback();
                }

                if (_connection.State != ConnectionState.Open) _connection.Open();

                _transaction = _connection.BeginTransaction(isolationLevel);
                _commandBuilder.RefreshSchema();

                return true;
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return false;
            }
        }

        public override bool CommitTransaction()
        {
            try
            {
                _errorMessage = String.Empty;
                if (_transaction != null)
                {
                    _transaction.Commit();
                    _commandBuilder.RefreshSchema();
                }
                return false;
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return false;
            }
        }

        public override bool RollbackTransaction()
        {
            try
            {
                _errorMessage = String.Empty;
                if (_transaction != null)
                {
                    _transaction.Rollback();
                    _commandBuilder.RefreshSchema();
                }
                return true;
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return false;
            }
        }

        public override IDataReader ExecuteReader(string sql, int commandTimeout, CommandType commandType)
        {
            _errorMessage = String.Empty;
            if (String.IsNullOrEmpty(sql)) return null;
            ConnectionState previousConnectionState = _connection.State;
            try
            {
                _command.CommandType = commandType;
                _command.CommandTimeout = commandTimeout;
                _command.CommandText = sql;

                if (_transaction != null) _command.Transaction = _transaction;
                _commandBuilder.RefreshSchema();
                if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();
                return _command.ExecuteReader() as IDataReader;
            }
            catch (Exception ex)
            {
                if (previousConnectionState == ConnectionState.Closed) _connection.Close();
                _errorMessage = ex.Message;
                return null;
            }
        }

        public override int ExecuteNonQuery(string sql, int commandTimeout, CommandType commandType)
        {
            _errorMessage = String.Empty;
            if (String.IsNullOrEmpty(sql)) return -1;
            ConnectionState previousConnectionState = _connection.State;
            try
            {
                _command.CommandType = commandType;
                _command.CommandTimeout = commandTimeout;
                _command.CommandText = sql;

                if (_transaction != null) _command.Transaction = _transaction;
                _commandBuilder.RefreshSchema();
                if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();
                return _command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return -1;
            }
            finally { if (previousConnectionState == ConnectionState.Closed) _connection.Close(); }
        }

        public override object ExecuteScalar(string sql, int commandTimeout, CommandType commandType)
        {
            _errorMessage = String.Empty;
            if (String.IsNullOrEmpty(sql)) return null;
            ConnectionState previousConnectionState = _connection.State;
            try
            {
                _command.CommandType = commandType;
                _command.CommandTimeout = commandTimeout;
                _command.CommandText = sql;

                if (_transaction != null) _command.Transaction = _transaction;
                _commandBuilder.RefreshSchema();
                if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();
                return _command.ExecuteScalar();
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return null;
            }
            finally { if (previousConnectionState == ConnectionState.Closed) _connection.Close(); }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR5 (Select by attributes interface)
        // Execute a SQL query to check if it is valid. Errors
        // will be thrown back to the calling method.
        //
        /// <summary>
        /// Validates the SQL query by executing it and throwing
        /// any execeptions raised back to the calling method.
        /// </summary>
        /// <param name="sql">The SQL query to validate (which should be a non-update query).</param>
        /// <param name="commandTimeout">The command timeout.</param>
        /// <param name="commandType">Type of the sql command.</param>
        /// <returns></returns>
        /// <exception cref="System.Exception">Sql is null or empty</exception>
        public override bool ValidateQuery(string sql, int commandTimeout, CommandType commandType)
        {
            _errorMessage = String.Empty;
            if (String.IsNullOrEmpty(sql)) throw (new Exception("Sql is null or empty"));
            ConnectionState previousConnectionState = _connection.State;
            try
            {
                _command.CommandType = commandType;
                _command.CommandTimeout = commandTimeout;
                _command.CommandText = sql;

                if (_transaction != null) _command.Transaction = _transaction;
                _commandBuilder.RefreshSchema();
                if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();
                _command.ExecuteNonQuery();
                return true;
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                throw ex;
            }
            finally { if (previousConnectionState == ConnectionState.Closed) _connection.Close(); }
        }
        //---------------------------------------------------------------------

        public override int Update<T>(T table, string insertCommand, string updateCommand, string deleteCommand)
        {
            ConnectionState previousConnectionState = _connection.State;
            if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();

            if ((table == null) || (table.Rows.Count == 0)) return 0;
            if (_adapter == null) return -1;
            
            try
            {
                if (!String.IsNullOrEmpty(insertCommand))
                    _adapter.InsertCommand = new OdbcCommand(insertCommand);
                if (!String.IsNullOrEmpty(updateCommand))
                    _adapter.UpdateCommand = new OdbcCommand(updateCommand);
                if (!String.IsNullOrEmpty(deleteCommand))
                    _adapter.DeleteCommand = new OdbcCommand(deleteCommand);
                
                return _adapter.Update(table); ;
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return -1;
            }
            finally { if (previousConnectionState == ConnectionState.Closed) _connection.Close(); }
        }

        public override int Update<T>(T table)
        {
            ConnectionState previousConnectionState = _connection.State;
            if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();

            if ((table == null) || (table.Rows.Count == 0)) return 0;

            try
            {
                OdbcDataAdapter adapter = UpdateAdapter(table);

                if (adapter != null)
                    return adapter.Update(table);
                else
                    return -1;
            }
            catch { return -1; }
            finally { if (previousConnectionState == ConnectionState.Closed) _connection.Close(); }
        }

        public override int Update<T>(T dataSet, string sourceTable)
        {
            ConnectionState previousConnectionState = _connection.State;
            if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();

            if ((dataSet == null) || String.IsNullOrEmpty(sourceTable) || 
                !dataSet.Tables.Contains(sourceTable)) return 0;

            try
            {
                DataTable table = dataSet.Tables[sourceTable];

                OdbcDataAdapter adapter = UpdateAdapter(table);

                if (adapter != null)
                    return adapter.Update(table);
                else
                    return -1;
            }
            catch { return -1; }
            finally { if (previousConnectionState == ConnectionState.Closed) _connection.Close(); }
        }

        public override int Update<T, R>(R[] rows)
        {
            ConnectionState previousConnectionState = _connection.State;
            if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();

            if ((rows == null) || (rows.Length == 0)) return 0;

            try
            {
                T table = (T)rows[0].Table;

                OdbcDataAdapter adapter = UpdateAdapter(table);

                if (adapter != null)
                    return adapter.Update(rows);
                else
                    return -1;
            }
            catch { return -1; }
            finally { if (previousConnectionState == ConnectionState.Closed) _connection.Close(); }
        }

        private OdbcDataAdapter UpdateAdapter<T>(T table) where T : DataTable, new()
        {
            if (table == null) return null;

            OdbcDataAdapter adapter;
            if (!_adaptersDic.TryGetValue(typeof(T), out adapter))
            {
                CreateAdapter<T>(table);
                if (!_adaptersDic.TryGetValue(typeof(T), out adapter)) return null;
            }

            if (_transaction != null)
            {
                if ((adapter.InsertCommand != null) &&
                    ((adapter.InsertCommand.Transaction == null) || !adapter.InsertCommand.Transaction.Equals(_transaction)))
                    adapter.InsertCommand.Transaction = _transaction;
                if ((adapter.UpdateCommand != null) &&
                    ((adapter.UpdateCommand.Transaction == null) || !adapter.UpdateCommand.Transaction.Equals(_transaction)))
                    adapter.UpdateCommand.Transaction = _transaction;
                if ((adapter.DeleteCommand != null) &&
                    ((adapter.DeleteCommand.Transaction == null) || !adapter.DeleteCommand.Transaction.Equals(_transaction)))
                    adapter.DeleteCommand.Transaction = _transaction;
            }

            return adapter;
        }
        
        #endregion

        #region Protected Members

        protected override string ParameterPrefix
        {
            get { return "@"; }
        }

        #region Browse Connection

        protected override void BrowseConnection()
        {
            try
            {
                _connWindow = new HLU.UI.View.Connection.ViewConnectOdbc();
                if ((_connWindow.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));
                _connWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;

                // create ViewModel to which main window binds
                _connViewModel = new HLU.UI.ViewModel.ViewModelConnectOdbc();
                _connViewModel.DisplayName = "ODBC Connection";

                // when ViewModel asks to be closed, close window
                _connViewModel.RequestClose +=
                    new HLU.UI.ViewModel.ViewModelConnectOdbc.RequestCloseEventHandler(_connViewModel_RequestClose);

                // allow all controls in window to bind to ViewModel by setting DataContext
                _connWindow.DataContext = _connViewModel;

                _connWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                _connWindow.Topmost = true;

                // show window
                _connWindow.ShowDialog();

                if (!String.IsNullOrEmpty(_errorMessage)) throw (new Exception(_errorMessage));
            }
            catch (Exception ex)
            {
                MessageBox.Show("ODBC Server responded with an error:\n\n" + ex.Message, "ODBC Server Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }

        void _connViewModel_RequestClose(string connString, string defaultSchema, string errorMsg)
        {
            _connViewModel.RequestClose -= _connViewModel_RequestClose;
            _connWindow.Close();

            if (!String.IsNullOrEmpty(errorMsg))
            {
                _errorMessage = errorMsg;
            }
            else if (!String.IsNullOrEmpty(connString))
            {
                ConnectionString = connString;
                DefaultSchema = defaultSchema;
            }
        }

        #endregion

        #endregion

        #endregion

        #region SQLBuilder Members

        #region Public Members

        public override string QuotePrefix { get { return _quotePrefix; } }

        public override string QuoteSuffix { get { return _quoteSuffix; } }

        public override string StringLiteralDelimiter { get { return _stringLiteralDelimiter; } }

        public override string DateLiteralPrefix { get { return _dateLiteralPrefix; } }

        public override string DateLiteralSuffix { get { return _dateLiteralSuffix; } }

        public override string WildcardSingleMatch { get { return _wildcardSingleMatch; } }

        public override string WildcardManyMatch { get { return _wildcardManyMatch; } }

        /// <summary>
        /// Does not escape string delimiter or other special characters.
        /// Does check if value is already quoted.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public override string QuoteValue(object value)
        {
            if (value == null) return "NULL";
            Type valueType = value.GetType();
            int colType;
            if (_typeMapSystemToSQL.TryGetValue(valueType, out colType))
            {
                //---------------------------------------------------------------------
                // FIX: 030 Include time in DateTime fields with ODBC connections
                // Ensure that updates to databases using ODBC connection type
                // include the time when updating DateTime fields.
                string s = valueType == typeof(DateTime) ? ((DateTime)value).ToString("s").Replace("T", " ") : value.ToString();
                //---------------------------------------------------------------------
                switch ((OdbcType)colType)
                {
                    case OdbcType.Char:
                    case OdbcType.NChar:
                    case OdbcType.NText:
                    case OdbcType.NVarChar:
                    case OdbcType.Text:
                    case OdbcType.VarChar:
                        if (s.Length == 0) return StringLiteralDelimiter + StringLiteralDelimiter;
                        if (!s.StartsWith(StringLiteralDelimiter)) s = StringLiteralDelimiter + s;
                        if (!s.EndsWith(StringLiteralDelimiter)) s += StringLiteralDelimiter;
                        return s;
                    case OdbcType.Date:
                    case OdbcType.DateTime:
                    case OdbcType.SmallDateTime:
                    case OdbcType.Time:
                    case OdbcType.Timestamp:
                        if (s.Length == 0) return DateLiteralPrefix + DateLiteralSuffix;
                        if (!s.StartsWith(DateLiteralPrefix)) s = DateLiteralPrefix + s;
                        if (!s.EndsWith(DateLiteralSuffix)) s += DateLiteralSuffix;
                        return s;
                    default:
                        return s;
                }
            }
            else
            {
                return value.ToString();
            }
        }
        
        #endregion
        
        #endregion

        #region Private Methods
        
        private void PopulateTypeMaps(bool isUnicode, bool useTimeZone, uint textLength, uint binaryLength,
            uint timePrecision, uint numericPrecision, uint numericScale)
        {
            string sTimeZone = useTimeZone ? " WITH TIME ZONE" : "";

            GetMetaData(typeof(OdbcType), _connection, _transaction);

            Dictionary<Type, int> typeMapSystemToSQLAdd = new Dictionary<Type, int>();
            typeMapSystemToSQLAdd.Add(typeof(Object), (int)OdbcType.VarBinary);
            typeMapSystemToSQLAdd.Add(typeof(Boolean), (int)OdbcType.Bit);
            typeMapSystemToSQLAdd.Add(typeof(SByte), (int)OdbcType.Int);
            typeMapSystemToSQLAdd.Add(typeof(Byte), (int)OdbcType.TinyInt);
            typeMapSystemToSQLAdd.Add(typeof(Int16), (int)OdbcType.SmallInt);
            typeMapSystemToSQLAdd.Add(typeof(UInt16), (int)OdbcType.SmallInt);
            typeMapSystemToSQLAdd.Add(typeof(Int32), (int)OdbcType.Int);
            typeMapSystemToSQLAdd.Add(typeof(UInt32), (int)OdbcType.Int);
            typeMapSystemToSQLAdd.Add(typeof(Int64), (int)OdbcType.BigInt);
            typeMapSystemToSQLAdd.Add(typeof(UInt64), (int)OdbcType.BigInt);
            typeMapSystemToSQLAdd.Add(typeof(Single), (int)OdbcType.Real);
            typeMapSystemToSQLAdd.Add(typeof(Double), (int)OdbcType.Double);
            typeMapSystemToSQLAdd.Add(typeof(Decimal), (int)OdbcType.Decimal);
            typeMapSystemToSQLAdd.Add(typeof(DateTime), (int)OdbcType.DateTime);
            typeMapSystemToSQLAdd.Add(typeof(TimeSpan), (int)OdbcType.DateTime);
            typeMapSystemToSQLAdd.Add(typeof(Byte[]), (int)OdbcType.VarBinary);
            typeMapSystemToSQLAdd.Add(typeof(Guid), (int)OdbcType.UniqueIdentifier);
            if (isUnicode)
            {
                typeMapSystemToSQLAdd.Add(typeof(Char), (int)OdbcType.NChar);
                typeMapSystemToSQLAdd.Add(typeof(String), (int)OdbcType.NText);
                typeMapSystemToSQLAdd.Add(typeof(Char[]), (int)OdbcType.NText);
            }
            else
            {
                typeMapSystemToSQLAdd.Add(typeof(Char), (int)OdbcType.Char);
                typeMapSystemToSQLAdd.Add(typeof(String), (int)OdbcType.Text);
                typeMapSystemToSQLAdd.Add(typeof(Char[]), (int)OdbcType.Text);
            }

            Dictionary<int, Type> typeMapSQLToSystemAdd = new Dictionary<int, Type>();
            typeMapSQLToSystemAdd.Add((int)OdbcType.BigInt, typeof(Int64));
            typeMapSQLToSystemAdd.Add((int)OdbcType.Binary, typeof(Byte[]));
            typeMapSQLToSystemAdd.Add((int)OdbcType.Bit, typeof(Boolean));
            typeMapSQLToSystemAdd.Add((int)OdbcType.Char, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OdbcType.Date, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)OdbcType.Decimal, typeof(System.Decimal));
            typeMapSQLToSystemAdd.Add((int)OdbcType.Double, typeof(Double));
            typeMapSQLToSystemAdd.Add((int)OdbcType.Image, typeof(Byte[]));
            typeMapSQLToSystemAdd.Add((int)OdbcType.Int, typeof(Int32));
            typeMapSQLToSystemAdd.Add((int)OdbcType.NChar, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OdbcType.NText, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OdbcType.Numeric, typeof(Decimal));
            typeMapSQLToSystemAdd.Add((int)OdbcType.NVarChar, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OdbcType.Real, typeof(Single));
            typeMapSQLToSystemAdd.Add((int)OdbcType.SmallDateTime, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)OdbcType.SmallInt, typeof(Int16));
            typeMapSQLToSystemAdd.Add((int)OdbcType.Text, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OdbcType.Time, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)OdbcType.Timestamp, typeof(Byte[]));
            typeMapSQLToSystemAdd.Add((int)OdbcType.TinyInt, typeof(Byte));
            typeMapSQLToSystemAdd.Add((int)OdbcType.UniqueIdentifier, typeof(Guid));
            typeMapSQLToSystemAdd.Add((int)OdbcType.VarBinary, typeof(Byte[]));
            typeMapSQLToSystemAdd.Add((int)OdbcType.VarChar, typeof(String));

            Dictionary<string, int> sqlSynonymsAdd = new Dictionary<string, int>();
            sqlSynonymsAdd.Add("bit", (int)OdbcType.Bit);
            sqlSynonymsAdd.Add("character", (int)OdbcType.Char);
            sqlSynonymsAdd.Add("char", (int)OdbcType.Char);
            sqlSynonymsAdd.Add("date", (int)OdbcType.Date);
            sqlSynonymsAdd.Add("decimal", (int)OdbcType.Decimal);
            sqlSynonymsAdd.Add("dec", (int)OdbcType.Decimal);
            sqlSynonymsAdd.Add("float", (int)OdbcType.Double);
            sqlSynonymsAdd.Add("double precision", (int)OdbcType.Double);
            sqlSynonymsAdd.Add("integer", (int)OdbcType.Int);
            sqlSynonymsAdd.Add("int", (int)OdbcType.Int);
            sqlSynonymsAdd.Add("national character", (int)OdbcType.NChar);
            sqlSynonymsAdd.Add("national char", (int)OdbcType.NChar);
            sqlSynonymsAdd.Add("nchar", (int)OdbcType.NChar);
            sqlSynonymsAdd.Add("numeric", (int)OdbcType.Numeric);
            sqlSynonymsAdd.Add("national character varying", (int)OdbcType.NVarChar);
            sqlSynonymsAdd.Add("national char varying", (int)OdbcType.NVarChar);
            sqlSynonymsAdd.Add("nchar varying", (int)OdbcType.NVarChar);
            sqlSynonymsAdd.Add("real", (int)OdbcType.Real);
            sqlSynonymsAdd.Add("smallint", (int)OdbcType.SmallInt);
            sqlSynonymsAdd.Add("time", (int)OdbcType.Time);
            sqlSynonymsAdd.Add("time with time zone", (int)OdbcType.Time);
            sqlSynonymsAdd.Add("timestamp", (int)OdbcType.Timestamp);
            sqlSynonymsAdd.Add("timestamp with time zone", (int)OdbcType.Timestamp);
            sqlSynonymsAdd.Add("bit varying", (int)OdbcType.VarBinary);
            sqlSynonymsAdd.Add("character varying", (int)OdbcType.VarChar);
            sqlSynonymsAdd.Add("char varying", (int)OdbcType.VarChar);
            sqlSynonymsAdd.Add("varchar", (int)OdbcType.VarChar);

            foreach (KeyValuePair<Type, int> kv in typeMapSystemToSQLAdd)
            {
                if (!_typeMapSystemToSQL.ContainsKey(kv.Key))
                    _typeMapSystemToSQL.Add(kv.Key, kv.Value);
            }

            if (isUnicode)
            {
                ReplaceType(typeof(Char), (int)OdbcType.NChar, _typeMapSystemToSQL);
                ReplaceType(typeof(String), (int)OdbcType.NText, _typeMapSystemToSQL);
                ReplaceType(typeof(Char[]), (int)OdbcType.NText, _typeMapSystemToSQL);
            }
            else
            {
                ReplaceType(typeof(Char), (int)OdbcType.Char, _typeMapSystemToSQL);
                ReplaceType(typeof(String), (int)OdbcType.Text, _typeMapSystemToSQL);
                ReplaceType(typeof(Char[]), (int)OdbcType.Text, _typeMapSystemToSQL);
            }

            foreach (KeyValuePair<int, Type> kv in typeMapSQLToSystemAdd)
            {
                if (!_typeMapSQLToSystem.ContainsKey(kv.Key))
                    _typeMapSQLToSystem.Add(kv.Key, kv.Value);
            }

            foreach (KeyValuePair<string, int> kv in sqlSynonymsAdd)
            {
                if (!_sqlSynonyms.ContainsKey(kv.Key))
                    _sqlSynonyms.Add(kv.Key, kv.Value);
            }

            _typeMapSQLToSQLCode = new Dictionary<int, string>();
            _typeMapSQLCodeToSQL = new Dictionary<string, int>();

            _typeMapSQLToSQLCode.Add((int)OdbcType.BigInt, "INTEGER");
            _typeMapSQLToSQLCode.Add((int)OdbcType.Bit, "BIT");
            _typeMapSQLToSQLCode.Add((int)OdbcType.Char, "CHARACTER");
            _typeMapSQLToSQLCode.Add((int)OdbcType.Date, "DATE");
            _typeMapSQLToSQLCode.Add((int)OdbcType.Double, "DOUBLE PRECISION");
            _typeMapSQLToSQLCode.Add((int)OdbcType.Int, "INTEGER");
            _typeMapSQLToSQLCode.Add((int)OdbcType.NChar, "NATIONAL CHARACTER");
            _typeMapSQLToSQLCode.Add((int)OdbcType.Real, "REAL");
            _typeMapSQLToSQLCode.Add((int)OdbcType.SmallDateTime, "DATE");
            _typeMapSQLToSQLCode.Add((int)OdbcType.SmallInt, "SMALLINT");
            _typeMapSQLToSQLCode.Add((int)OdbcType.Timestamp, String.Format("TIMESTAMP ({0})", sTimeZone));
            _typeMapSQLToSQLCode.Add((int)OdbcType.TinyInt, "SMALLINT");
            _typeMapSQLToSQLCode.Add((int)OdbcType.UniqueIdentifier, "CHARACTER (36)");

            //_typeMapSQLCodeToSQL.Add("INTEGER", (int)OdbcType.BigInt);
            _typeMapSQLCodeToSQL.Add("INTEGER", (int)OdbcType.Int);
            _typeMapSQLCodeToSQL.Add("BIT", (int)OdbcType.Bit);
            _typeMapSQLCodeToSQL.Add("CHARACTER", (int)OdbcType.Char);
            //_typeMapSQLCodeToSQL.Add("CHARACTER (36)", (int)OdbcType.UniqueIdentifier);
            _typeMapSQLCodeToSQL.Add("DATE", (int)OdbcType.Date);
            //_typeMapSQLCodeToSQL.Add("DATE", (int)OdbcType.SmallDateTime);
            _typeMapSQLCodeToSQL.Add("REAL", (int)OdbcType.Real);
            _typeMapSQLCodeToSQL.Add("SMALLINT", (int)OdbcType.SmallInt);
            //_typeMapSQLCodeToSQL.Add("SMALLINT", (int)OdbcType.TinyInt);
            
            if (binaryLength > 0)
            {
                _typeMapSQLToSQLCode.Add((int)OdbcType.Binary, String.Format("BIT VARYING ({0})", binaryLength));
                _typeMapSQLToSQLCode.Add((int)OdbcType.Image, String.Format("BIT VARYING ({0})", binaryLength));
                _typeMapSQLToSQLCode.Add((int)OdbcType.VarBinary, String.Format("BIT VARYING ({0})", binaryLength));

                //_typeMapSQLCodeToSQL.Add(String.Format("BIT VARYING ({0})", binaryLength), (int)OdbcType.Binary);
                //_typeMapSQLCodeToSQL.Add(String.Format("BIT VARYING ({0})", binaryLength), (int)OdbcType.Image);
                _typeMapSQLCodeToSQL.Add(String.Format("BIT VARYING ({0})", binaryLength), (int)OdbcType.VarBinary);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)OdbcType.Binary, "BIT VARYING");
                _typeMapSQLToSQLCode.Add((int)OdbcType.Image, "BIT VARYING");
                _typeMapSQLToSQLCode.Add((int)OdbcType.VarBinary, "BIT VARYING");

                //_typeMapSQLCodeToSQL.Add("BIT VARYING", (int)OdbcType.Binary);
                //_typeMapSQLCodeToSQL.Add("BIT VARYING", (int)OdbcType.Image);
                _typeMapSQLCodeToSQL.Add("BIT VARYING", (int)OdbcType.VarBinary);
            }
            if (textLength > 0)
            {
                _typeMapSQLToSQLCode.Add((int)OdbcType.NText, String.Format("NATIONAL CHARACTER VARYING ({0})", textLength));
                _typeMapSQLToSQLCode.Add((int)OdbcType.NVarChar, String.Format("NATIONAL CHARACTER VARYING ({0})", textLength));
                _typeMapSQLToSQLCode.Add((int)OdbcType.Text, String.Format("CHARACTER VARYING ({0})", textLength));
                _typeMapSQLToSQLCode.Add((int)OdbcType.VarChar, String.Format("CHARACTER VARYING ({0})", textLength));

                //_typeMapSQLCodeToSQL.Add(String.Format("NATIONAL CHARACTER VARYING ({0})", textLength), (int)OdbcType.NText);
                _typeMapSQLCodeToSQL.Add(String.Format("NATIONAL CHARACTER VARYING ({0})", textLength), (int)OdbcType.NVarChar);
                //_typeMapSQLCodeToSQL.Add(String.Format("CHARACTER VARYING ({0})", textLength), (int)OdbcType.Text);
                _typeMapSQLCodeToSQL.Add(String.Format("CHARACTER VARYING ({0})", textLength), (int)OdbcType.VarChar);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)OdbcType.NText, "NATIONAL CHARACTER VARYING");
                _typeMapSQLToSQLCode.Add((int)OdbcType.NVarChar, "NATIONAL CHARACTER VARYING");
                _typeMapSQLToSQLCode.Add((int)OdbcType.Text, "CHARACTER VARYING");
                _typeMapSQLToSQLCode.Add((int)OdbcType.VarChar, "CHARACTER VARYING");

                //_typeMapSQLCodeToSQL.Add("NATIONAL CHARACTER VARYING", (int)OdbcType.NText);
                _typeMapSQLCodeToSQL.Add("NATIONAL CHARACTER VARYING", (int)OdbcType.NVarChar);
                //_typeMapSQLCodeToSQL.Add("CHARACTER VARYING", (int)OdbcType.Text);
                _typeMapSQLCodeToSQL.Add("CHARACTER VARYING", (int)OdbcType.VarChar);
            }
            if ((numericPrecision > 0) && (numericScale > 0))
            {
                _typeMapSQLToSQLCode.Add((int)OdbcType.Decimal, String.Format("DECIMAL ({0},{1})", numericPrecision, numericScale));
                _typeMapSQLToSQLCode.Add((int)OdbcType.Numeric, String.Format("NUMERIC ({0},{1})", numericPrecision, numericScale));

                _typeMapSQLCodeToSQL.Add(String.Format("DECIMAL ({0},{1})", numericPrecision, numericScale), (int)OdbcType.Decimal);
                _typeMapSQLCodeToSQL.Add(String.Format("NUMERIC ({0},{1})", numericPrecision, numericScale), (int)OdbcType.Numeric);
            }
            else if (numericPrecision > 0)
            {
                _typeMapSQLToSQLCode.Add((int)OdbcType.Decimal, String.Format("DECIMAL ({0})", numericPrecision));
                _typeMapSQLToSQLCode.Add((int)OdbcType.Numeric, String.Format("NUMERIC ({0})", numericPrecision));

                _typeMapSQLCodeToSQL.Add(String.Format("DECIMAL ({0})", numericPrecision), (int)OdbcType.Decimal);
                _typeMapSQLCodeToSQL.Add(String.Format("NUMERIC ({0})", numericPrecision), (int)OdbcType.Numeric);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)OdbcType.Decimal, "DECIMAL");
                _typeMapSQLToSQLCode.Add((int)OdbcType.Numeric, "NUMERIC");

                _typeMapSQLCodeToSQL.Add("DECIMAL", (int)OdbcType.Decimal);
                _typeMapSQLCodeToSQL.Add("NUMERIC", (int)OdbcType.Numeric);
            }
            if (timePrecision > 0)
            {
                _typeMapSQLToSQLCode.Add((int)OdbcType.Time, String.Format("TIME ({0}){1}", timePrecision, sTimeZone));
                _typeMapSQLCodeToSQL.Add(String.Format("TIME ({0}){1}", timePrecision, sTimeZone), (int)OdbcType.Time);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)OdbcType.Time, String.Format("TIME ({0})", sTimeZone));
                _typeMapSQLCodeToSQL.Add(String.Format("TIME ({0})", sTimeZone), (int)OdbcType.Time);
            }
        }

        private void SetDefaults()
        {
            _backend = GetBackend(_connection);

            switch (_backend)
            {
                case Backends.Access:
                    _quotePrefix = "[";
                    _quoteSuffix = "]";
                    _stringLiteralDelimiter = "'";
                    _dateLiteralPrefix = "#";
                    _dateLiteralSuffix = "#";
                    _wildcardSingleMatch = "?";
                    _wildcardManyMatch = "*";
                    break;
                case Backends.SqlServer:
                    _quotePrefix = "[";
                    _quoteSuffix = "]";
                    _stringLiteralDelimiter = "'";
                    _dateLiteralPrefix = "'";
                    _dateLiteralSuffix = "'";
                    _wildcardSingleMatch = "_";
                    _wildcardManyMatch = "%";
                    break;
                default:
                    _quotePrefix = "\"";
                    _quoteSuffix = "\"";
                    _stringLiteralDelimiter = "'";
                    _dateLiteralPrefix = "'";
                    _dateLiteralSuffix = "'";
                    _wildcardSingleMatch = "_";
                    _wildcardManyMatch = "%";
                    break;
            }
        }
        
        #endregion
    }
}
