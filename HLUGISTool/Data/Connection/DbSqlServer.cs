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
using System.Data.SqlClient;
using System.Linq;
using System.Windows;
using System.Text;

namespace HLU.Data.Connection
{
    class DbSqlServer : DbBase
    {
        #region Private Members
        
        private string _errorMessage;
        private SqlConnectionStringBuilder _connStrBuilder;
        private SqlConnection _connection;
        private SqlCommand _command;
        private SqlDataAdapter _adapter;
        private SqlCommandBuilder _commandBuilder;
        private SqlTransaction _transaction;
        private Dictionary<Type, SqlDataAdapter> _adaptersDic = new Dictionary<Type, SqlDataAdapter>();

        private HLU.UI.View.Connection.ViewConnectSqlServer _connWindow;
        private HLU.UI.ViewModel.ViewModelConnectSqlServer _connViewModel;
       
        #endregion

        #region Constructor

        public DbSqlServer(ref string connString, ref string defaultSchema, ref bool promptPwd, string pwdMask,
            bool useCommandBuilder, bool useColumnNames, bool isUnicode, bool useTimeZone, uint textLength, 
            uint binaryLength, uint timePrecision, uint numericPrecision, uint numericScale)
            : base(ref connString, ref defaultSchema, ref promptPwd, pwdMask, useCommandBuilder, useColumnNames, 
            isUnicode, useTimeZone, textLength, binaryLength,timePrecision, numericPrecision, numericScale)
        {
            if (String.IsNullOrEmpty(ConnectionString)) throw (new Exception("No connection string"));

            try
            {
                Login("User name", ConnectionString, ref promptPwd, ref _connStrBuilder, ref _connection);

                PopulateTypeMaps(IsUnicode, TextLength, BinaryLength,
                    TimePrecision, NumericPrecision, NumericScale);

                _command = _connection.CreateCommand();
                _adapter = new SqlDataAdapter(_command);
                _commandBuilder = new SqlCommandBuilder(_adapter);

                connString = MaskPassword(_connStrBuilder, pwdMask);
                defaultSchema = DefaultSchema;
            }
            catch { throw; }
        }
        
        #endregion

        #region DbBase Members
        
        #region Public Members

        public override Backends Backend { get { return Backends.SqlServer; } }

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
                                           DataType = r.Field<string>("DATA_TYPE")
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
                                                 ((SqlDbType)SystemToDbType(dsCol.DataType) + ")").ToString()).ToArray();
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
                return new SqlCommand();
        }

        public override IDbDataAdapter CreateAdapter()
        {
            return new SqlDataAdapter();
        }

        public override IDbDataAdapter CreateAdapter<T>(T table)
        {
            if (table == null) table = new T();

            SqlDataAdapter adapter;

            if (!_adaptersDic.TryGetValue(typeof(T), out adapter))
            {
                adapter = new SqlDataAdapter();

                DataColumn[] pk = table.PrimaryKey;
                if ((pk == null) || (pk.Length == 0)) return null;

                DataTableMapping tableMapping = new DataTableMapping();
                tableMapping.SourceTable = table.TableName; // "Table";
                tableMapping.DataSetTable = table.TableName; // "Exports";

                List<SqlParameter> deleteParams = new List<SqlParameter>();
                List<SqlParameter> insertParams = new List<SqlParameter>();
                List<SqlParameter> updateParams = new List<SqlParameter>();
                List<SqlParameter> updateParamsOrig = new List<SqlParameter>();

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
                SqlDbType isNullType = (SqlDbType)isNullTypeInt;

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
                        (SqlDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Original, false));

                    insColParamName = ParameterName(_parameterPrefixCurr, c.ColumnName, i + _startParamNo);
                    insertParams.Add(CreateParameter(insColParamName,
                        (SqlDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Current, false));

                    updColParamName = ParameterName(_parameterPrefixCurr, c.ColumnName, i + _startParamNo);
                    updateParams.Add(CreateParameter(updColParamName,
                        (SqlDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Current, false));

                    updOrigParamName = ParameterName(_parameterPrefixOrig, c.ColumnName, 
                        i + _startParamNo + columnCount + nullParamCount);
                    updateParamsOrig.Add(CreateParameter(updOrigParamName,
                        (SqlDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Original, false));

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

                adapter.SelectCommand = new SqlCommand();
                adapter.SelectCommand.CommandType = CommandType.Text;
                adapter.SelectCommand.Connection = _connection;
                adapter.SelectCommand.CommandText = String.Format("SELECT {0} FROM {1}", sbTargetList, tableName);

                if (!_useCommandBuilder)
                {
                    adapter.DeleteCommand = new SqlCommand();
                    adapter.DeleteCommand.CommandType = CommandType.Text;
                    adapter.DeleteCommand.Connection = _connection;
                    adapter.DeleteCommand.CommandText = String.Format("DELETE FROM {0} WHERE {1}", tableName, sbWhereDel);
                    adapter.DeleteCommand.Parameters.AddRange(deleteParams.ToArray());

                    adapter.UpdateCommand = new SqlCommand();
                    adapter.UpdateCommand.Connection = _connection;
                    adapter.UpdateCommand.CommandType = CommandType.Text;
                    adapter.UpdateCommand.CommandText =
                        String.Format("UPDATE {0} SET {1} WHERE {2}", tableName, sbUpdSetList, sbWhereUpd);
                    adapter.UpdateCommand.Parameters.AddRange(updateParams.ToArray());

                    adapter.InsertCommand = new SqlCommand();
                    adapter.InsertCommand.CommandType = CommandType.Text;
                    adapter.InsertCommand.Connection = _connection;
                    adapter.InsertCommand.CommandText = String.Format("INSERT INTO {0} ({1}) VALUES ({2})",
                        tableName, sbTargetList, sbInsValues);
                    adapter.InsertCommand.Parameters.AddRange(insertParams.ToArray());
                }
                else
                {
                    SqlCommandBuilder cmdBuilder = new SqlCommandBuilder(adapter);
                    adapter.DeleteCommand = cmdBuilder.GetDeleteCommand(_useColumnNames);
                    adapter.UpdateCommand = cmdBuilder.GetUpdateCommand(_useColumnNames);
                    adapter.InsertCommand = cmdBuilder.GetInsertCommand(_useColumnNames);
                }

                adapter.UpdateCommand.CommandText += ";\r\n" +
                    String.Format("SELECT {0} FROM {1} WHERE {2}", sbTargetList, tableName, sbWherePkUpd);
                adapter.InsertCommand.CommandText += ";\r\n" +
                    String.Format("SELECT {0} FROM {1} WHERE {2}", sbTargetList, tableName, sbWherePkIns);

                if (typeof(T) != typeof(DataTable))
                    _adaptersDic.Add(typeof(T), adapter);
            }

            return adapter;
        }

        private SqlParameter CreateParameter(string name, SqlDbType type, ParameterDirection direction,
            string srcColumn, DataRowVersion srcVersion, bool nullMapping)
        {
            SqlParameter param = new SqlParameter(name, type);
            param.Direction = direction;
            param.SourceColumn = srcColumn;
            param.SourceVersion = srcVersion;
            param.SourceColumnNullMapping = nullMapping;
            return param;
        }

        private SqlParameter CreateParameter(string name, object value, ParameterDirection direction,
            string srcColumn, DataRowVersion srcVersion, bool nullMapping)
        {
            SqlParameter param = new SqlParameter(name, value);
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
            return parameterName;
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
                SqlDataAdapter adapter = UpdateAdapter(table);
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
                    _adapter = new SqlDataAdapter(_command);
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
                if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();
                SqlDataAdapter adapter = UpdateAdapter(table);
                if (adapter != null)
                {
                    if (_transaction != null)
                        adapter.SelectCommand.Transaction = _transaction;
                    return adapter.Fill(table);
                }
                else
                {
                    _command.CommandText = sql;
                    _command.CommandType = CommandType.Text;
                    if (_transaction != null) _command.Transaction = _transaction;
                    _adapter = new SqlDataAdapter(_command);
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

        public override bool BeginTransaction(bool commitPrevious, IsolationLevel isolationLevel)
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

        public override int Update<T>(T table, string insertCommand, string updateCommand, string deleteCommand)
        {
            ConnectionState previousConnectionState = _connection.State;
            if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();

            if ((table == null) || (table.Rows.Count == 0)) return 0;
            if (_adapter == null) return -1;

            try
            {
                if (!String.IsNullOrEmpty(insertCommand))
                    _adapter.InsertCommand = new SqlCommand(insertCommand);
                if (!String.IsNullOrEmpty(updateCommand))
                    _adapter.UpdateCommand = new SqlCommand(updateCommand);
                if (!String.IsNullOrEmpty(deleteCommand))
                    _adapter.DeleteCommand = new SqlCommand(deleteCommand);

                return _adapter.Update(table);;
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
                SqlDataAdapter adapter = UpdateAdapter(table);

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

                SqlDataAdapter adapter = UpdateAdapter(table);

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

                SqlDataAdapter adapter = UpdateAdapter(table);

                if (adapter != null)
                    return adapter.Update(rows);
                else
                    return -1;
            }
            catch { return -1; }
            finally { if (previousConnectionState == ConnectionState.Closed) _connection.Close(); }
        }

        private SqlDataAdapter UpdateAdapter<T>(T table) where T : DataTable, new()
        {
            if (table == null) return null;

            SqlDataAdapter adapter;
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
                _connWindow = new HLU.UI.View.Connection.ViewConnectSqlServer();
                if ((_connWindow.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));
                _connWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;

                // create ViewModel to which main window binds
                _connViewModel = new HLU.UI.ViewModel.ViewModelConnectSqlServer();
                _connViewModel.DisplayName = "SQL Server Connection";

                // when ViewModel asks to be closed, close window
                _connViewModel.RequestClose +=
                    new HLU.UI.ViewModel.ViewModelConnectSqlServer.RequestCloseEventHandler(_connViewModel_RequestClose);

                // allow all controls in window to bind to ViewModel by setting DataContext
                _connWindow.DataContext = _connViewModel;

                _connWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                _connWindow.Topmost = true;

                // show window
                _connWindow.ShowDialog();

                // throw error if connection failed
                if (!String.IsNullOrEmpty(_errorMessage)) throw (new Exception(_errorMessage));
            }
            catch (Exception ex)
            {
                MessageBox.Show("SQL Server responded with an error:\n\n" + ex.Message,
                    "SQL Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }

        protected void _connViewModel_RequestClose(string connString, string defaultSchema, string errorMsg)
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
        
        public override string QuotePrefix { get { return "["; } }

        public override string QuoteSuffix { get { return "]"; } }

        public override string StringLiteralDelimiter { get { return "'"; } }

        public override string DateLiteralPrefix { get { return "'"; } }

        public override string DateLiteralSuffix { get { return "'"; } }

        public override string WildcardSingleMatch { get { return "_"; } }

        public override string WildcardManyMatch { get { return "%"; } }

        public override string ConcatenateOperator { get { return "+"; } }

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
                string s = valueType == typeof(DateTime) ? ((DateTime)value).ToString("s") : value.ToString();
                switch ((SqlDbType)colType)
                {
                    case SqlDbType.Char:
                    case SqlDbType.Text:
                    case SqlDbType.VarChar:
                    case SqlDbType.Xml:
                        if (s.Length == 0) return StringLiteralDelimiter + StringLiteralDelimiter;
                        if (!s.StartsWith(StringLiteralDelimiter)) s = StringLiteralDelimiter + s;
                        if (!s.EndsWith(StringLiteralDelimiter)) s += StringLiteralDelimiter;
                        return s;
                    case SqlDbType.NChar:
                    case SqlDbType.NText:
                    case SqlDbType.NVarChar:
                        if (s.Length == 0) return "N" + StringLiteralDelimiter + StringLiteralDelimiter;
                        if (!s.StartsWith(StringLiteralDelimiter)) s = "N" + StringLiteralDelimiter + s;
                        if (!s.EndsWith(StringLiteralDelimiter)) s += StringLiteralDelimiter;
                        return s;
                    case SqlDbType.Date:
                    case SqlDbType.DateTime:
                    case SqlDbType.DateTime2:
                    case SqlDbType.DateTimeOffset:
                    case SqlDbType.Time:
                    case SqlDbType.Timestamp:
                        if (s.Length == 0) return DateLiteralPrefix + DateLiteralSuffix ;
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

        private void PopulateTypeMaps(bool isUnicode, uint textLength, uint binaryLength,
            uint timePrecision, uint numericPrecision, uint numericScale)
        {
            GetMetaData(typeof(SqlDbType), _connection, _transaction);

            Dictionary<Type, int> typeMapSystemToSQLAdd = new Dictionary<Type, int>();
            typeMapSystemToSQLAdd.Add(typeof(Object), (int)SqlDbType.Variant);
            typeMapSystemToSQLAdd.Add(typeof(Boolean), (int)SqlDbType.Bit);
            typeMapSystemToSQLAdd.Add(typeof(SByte), (int)SqlDbType.Int);
            typeMapSystemToSQLAdd.Add(typeof(Byte), (int)SqlDbType.TinyInt);
            typeMapSystemToSQLAdd.Add(typeof(Int16), (int)SqlDbType.SmallInt);
            typeMapSystemToSQLAdd.Add(typeof(UInt16), (int)SqlDbType.SmallInt);
            typeMapSystemToSQLAdd.Add(typeof(Int32), (int)SqlDbType.Int);
            typeMapSystemToSQLAdd.Add(typeof(UInt32), (int)SqlDbType.Int);
            typeMapSystemToSQLAdd.Add(typeof(Int64), (int)SqlDbType.BigInt);
            typeMapSystemToSQLAdd.Add(typeof(UInt64), (int)SqlDbType.BigInt);
            typeMapSystemToSQLAdd.Add(typeof(Single), (int)SqlDbType.Real);
            typeMapSystemToSQLAdd.Add(typeof(Double), (int)SqlDbType.Float);
            typeMapSystemToSQLAdd.Add(typeof(Decimal), (int)SqlDbType.Decimal);
            typeMapSystemToSQLAdd.Add(typeof(DateTime), (int)SqlDbType.DateTime2);
            typeMapSystemToSQLAdd.Add(typeof(TimeSpan), (int)SqlDbType.DateTimeOffset);
            typeMapSystemToSQLAdd.Add(typeof(Byte[]), (int)SqlDbType.Image);
            typeMapSystemToSQLAdd.Add(typeof(Guid), (int)SqlDbType.UniqueIdentifier);
            if (isUnicode)
            {
                typeMapSystemToSQLAdd.Add(typeof(Char), (int)SqlDbType.NChar);
                typeMapSystemToSQLAdd.Add(typeof(String), (int)SqlDbType.NVarChar);
                typeMapSystemToSQLAdd.Add(typeof(Char[]), (int)SqlDbType.NVarChar);
            }
            else
            {
                typeMapSystemToSQLAdd.Add(typeof(Char), (int)SqlDbType.Char);
                typeMapSystemToSQLAdd.Add(typeof(String), (int)SqlDbType.VarChar);
                typeMapSystemToSQLAdd.Add(typeof(Char[]), (int)SqlDbType.VarChar);
            }

            Dictionary<int, Type> typeMapSQLToSystemAdd = new Dictionary<int, Type>();
            typeMapSQLToSystemAdd.Add((int)SqlDbType.BigInt, typeof(Int64));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.Binary, typeof(Byte[]));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.Bit, typeof(Boolean));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.Char, typeof(String));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.Date, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.DateTime, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.DateTime2, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.DateTimeOffset, typeof(TimeSpan));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.Decimal, typeof(Decimal));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.Float, typeof(Double));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.Image, typeof(Byte[]));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.Int, typeof(Int32));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.Money, typeof(Decimal));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.NChar, typeof(String));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.NText, typeof(String));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.NVarChar, typeof(String));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.Real, typeof(Single));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.SmallDateTime, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.SmallInt, typeof(Int16));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.SmallMoney, typeof(Decimal));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.Text, typeof(String));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.Time, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.Timestamp, typeof(Byte[]));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.TinyInt, typeof(Byte));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.UniqueIdentifier, typeof(Guid));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.VarBinary, typeof(Byte[]));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.VarChar, typeof(String));
            typeMapSQLToSystemAdd.Add((int)SqlDbType.Variant, typeof(Object));

            Dictionary<string, int> sqlSynonymsAdd = new Dictionary<string, int>();
            sqlSynonymsAdd.Add("bigint", (int)SqlDbType.BigInt);
            sqlSynonymsAdd.Add("binary", (int)SqlDbType.Binary);
            sqlSynonymsAdd.Add("bit", (int)SqlDbType.Bit);
            sqlSynonymsAdd.Add("char", (int)SqlDbType.Char);
            sqlSynonymsAdd.Add("character", (int)SqlDbType.Char);
            sqlSynonymsAdd.Add("date", (int)SqlDbType.Date);
            sqlSynonymsAdd.Add("datetime", (int)SqlDbType.DateTime);
            sqlSynonymsAdd.Add("datetime2", (int)SqlDbType.DateTime2);
            sqlSynonymsAdd.Add("datetimeoffset", (int)SqlDbType.DateTimeOffset);
            sqlSynonymsAdd.Add("decimal", (int)SqlDbType.Decimal);
            sqlSynonymsAdd.Add("dec", (int)SqlDbType.Decimal);
            sqlSynonymsAdd.Add("numeric", (int)SqlDbType.Decimal);
            sqlSynonymsAdd.Add("float", (int)SqlDbType.Float);
            sqlSynonymsAdd.Add("double precision", (int)SqlDbType.Float);
            sqlSynonymsAdd.Add("image", (int)SqlDbType.Image);
            sqlSynonymsAdd.Add("int", (int)SqlDbType.Int);
            sqlSynonymsAdd.Add("integer", (int)SqlDbType.Int);
            sqlSynonymsAdd.Add("money", (int)SqlDbType.Money);
            sqlSynonymsAdd.Add("nchar", (int)SqlDbType.NChar);
            sqlSynonymsAdd.Add("national character", (int)SqlDbType.NChar);
            sqlSynonymsAdd.Add("national char", (int)SqlDbType.NChar);
            sqlSynonymsAdd.Add("ntext", (int)SqlDbType.NText);
            sqlSynonymsAdd.Add("national text", (int)SqlDbType.NText);
            sqlSynonymsAdd.Add("nvarchar", (int)SqlDbType.NVarChar);
            sqlSynonymsAdd.Add("national character varying", (int)SqlDbType.NVarChar);
            sqlSynonymsAdd.Add("national char varying", (int)SqlDbType.NVarChar);
            sqlSynonymsAdd.Add("real", (int)SqlDbType.Real);
            sqlSynonymsAdd.Add("smaldatetime", (int)SqlDbType.SmallDateTime);
            sqlSynonymsAdd.Add("smallint", (int)SqlDbType.SmallInt);
            sqlSynonymsAdd.Add("smallmoney", (int)SqlDbType.SmallMoney);
            sqlSynonymsAdd.Add("text", (int)SqlDbType.Text);
            sqlSynonymsAdd.Add("time", (int)SqlDbType.Time);
            sqlSynonymsAdd.Add("timestamp", (int)SqlDbType.Timestamp);
            sqlSynonymsAdd.Add("rowversion", (int)SqlDbType.Timestamp);
            sqlSynonymsAdd.Add("tinyint", (int)SqlDbType.TinyInt);
            sqlSynonymsAdd.Add("uniqueidentifier", (int)SqlDbType.UniqueIdentifier);
            sqlSynonymsAdd.Add("varbinary", (int)SqlDbType.VarBinary);
            sqlSynonymsAdd.Add("binary varying", (int)SqlDbType.VarBinary);
            sqlSynonymsAdd.Add("varchar", (int)SqlDbType.VarChar);
            sqlSynonymsAdd.Add("char varying", (int)SqlDbType.VarChar);
            sqlSynonymsAdd.Add("character varying", (int)SqlDbType.VarChar);
            sqlSynonymsAdd.Add("sql_variant", (int)SqlDbType.Variant);

            foreach (KeyValuePair<Type, int> kv in typeMapSystemToSQLAdd)
            {
                if (!_typeMapSystemToSQL.ContainsKey(kv.Key))
                    _typeMapSystemToSQL.Add(kv.Key, kv.Value);
            }

            if (isUnicode)
            {
                ReplaceType(typeof(Char), (int)SqlDbType.NChar, _typeMapSystemToSQL);
                ReplaceType(typeof(String), (int)SqlDbType.NVarChar, _typeMapSystemToSQL);
                ReplaceType(typeof(Char[]), (int)SqlDbType.NVarChar, _typeMapSystemToSQL);
            }
            else
            {
                ReplaceType(typeof(Char), (int)SqlDbType.Char, _typeMapSystemToSQL);
                ReplaceType(typeof(String), (int)SqlDbType.VarChar, _typeMapSystemToSQL);
                ReplaceType(typeof(Char[]), (int)SqlDbType.VarChar, _typeMapSystemToSQL);
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

            _typeMapSQLToSQLCode.Add((int)SqlDbType.BigInt, "BIGINT");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.Bit, "BIT");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.Char, "CHAR");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.Date, "DATETIME");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.DateTime, "DATETIME");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.Float, "FLOAT");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.Image, "IMAGE");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.Int, "INT");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.Money, "MONEY");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.NChar, "NCHAR");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.NText, "NTEXT");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.SmallDateTime, "SMALLDATETIME");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.Real, "REAL");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.SmallInt, "SMALLINT");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.SmallMoney, "SMALLMONEY");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.Text, "TEXT");

            _typeMapSQLCodeToSQL.Add("BIGINT", (int)SqlDbType.BigInt);
            _typeMapSQLCodeToSQL.Add("BIT", (int)SqlDbType.Bit);
            _typeMapSQLCodeToSQL.Add("CHAR", (int)SqlDbType.Char);
            //_typeMapSQLCodeToSQL.Add("DATETIME", (int)SqlDbType.Date);
            _typeMapSQLCodeToSQL.Add("DATETIME", (int)SqlDbType.DateTime);
            _typeMapSQLCodeToSQL.Add("FLOAT", (int)SqlDbType.Float);
            _typeMapSQLCodeToSQL.Add("IMAGE", (int)SqlDbType.Image);
            _typeMapSQLCodeToSQL.Add("INT", (int)SqlDbType.Int);
            _typeMapSQLCodeToSQL.Add("MONEY", (int)SqlDbType.Money);
            _typeMapSQLCodeToSQL.Add("NCHAR", (int)SqlDbType.NChar);
            _typeMapSQLCodeToSQL.Add("NTEXT", (int)SqlDbType.NText);
            _typeMapSQLCodeToSQL.Add("REAL", (int)SqlDbType.Real);
            _typeMapSQLCodeToSQL.Add("SMALLDATETIME", (int)SqlDbType.SmallDateTime);
            _typeMapSQLCodeToSQL.Add("SMALLINT", (int)SqlDbType.SmallInt);
            _typeMapSQLCodeToSQL.Add("SMALLMONEY", (int)SqlDbType.SmallMoney);
            _typeMapSQLCodeToSQL.Add("TEXT", (int)SqlDbType.Text);
            if (timePrecision > 0)
            {
                _typeMapSQLToSQLCode.Add((int)SqlDbType.Time, String.Format("TIME ({0})", timePrecision));
                _typeMapSQLToSQLCode.Add((int)SqlDbType.DateTime2, String.Format("DATETIME2 ({0})", timePrecision));
                _typeMapSQLToSQLCode.Add((int)SqlDbType.DateTimeOffset, String.Format("DATETIMEOFFSET ({0})", timePrecision));

                _typeMapSQLCodeToSQL.Add(String.Format("TIME ({0})", timePrecision), (int)SqlDbType.Time);
                _typeMapSQLCodeToSQL.Add(String.Format("DATETIME2 ({0})", timePrecision), (int)SqlDbType.DateTime2);
                _typeMapSQLCodeToSQL.Add(String.Format("DATETIMEOFFSET ({0})", timePrecision), (int)SqlDbType.DateTimeOffset);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)SqlDbType.Time, "TIME");
                _typeMapSQLToSQLCode.Add((int)SqlDbType.DateTime2, "DATETIME2");
                _typeMapSQLToSQLCode.Add((int)SqlDbType.DateTimeOffset, "DATETIMEOFFSET");

                _typeMapSQLCodeToSQL.Add("TIME", (int)SqlDbType.Time);
                _typeMapSQLCodeToSQL.Add("DATETIME2", (int)SqlDbType.DateTime2);
                _typeMapSQLCodeToSQL.Add("DATETIMEOFFSET", (int)SqlDbType.DateTimeOffset);
            }
            _typeMapSQLToSQLCode.Add((int)SqlDbType.Timestamp, "TIMESTAMP");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.TinyInt, "TINYINT");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.UniqueIdentifier, "UNIQUEIDENTIFIER");
            _typeMapSQLToSQLCode.Add((int)SqlDbType.Variant, "SQL_VARIANT");

            _typeMapSQLCodeToSQL.Add("TIMESTAMP", (int)SqlDbType.Timestamp);
            _typeMapSQLCodeToSQL.Add("TINYINT", (int)SqlDbType.TinyInt);
            _typeMapSQLCodeToSQL.Add("UNIQUEIDENTIFIER", (int)SqlDbType.UniqueIdentifier);
            _typeMapSQLCodeToSQL.Add("SQL_VARIANT", (int)SqlDbType.Variant);

            if (binaryLength > 0)
            {
                _typeMapSQLToSQLCode.Add((int)SqlDbType.Binary, String.Format("BINARY ({0})", binaryLength));
                _typeMapSQLToSQLCode.Add((int)SqlDbType.VarBinary, String.Format("VARBINARY ({0})", binaryLength));

                _typeMapSQLCodeToSQL.Add(String.Format("BINARY ({0})", binaryLength), (int)SqlDbType.Binary);
                _typeMapSQLCodeToSQL.Add(String.Format("VARBINARY ({0})", binaryLength), (int)SqlDbType.VarBinary);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)SqlDbType.Binary, "BINARY");
                _typeMapSQLToSQLCode.Add((int)SqlDbType.VarBinary, "VARBINARY");

                _typeMapSQLCodeToSQL.Add(String.Format("BINARY", binaryLength), (int)SqlDbType.Binary);
                _typeMapSQLCodeToSQL.Add(String.Format("VARBINARY", binaryLength), (int)SqlDbType.VarBinary);
            }
            if ((numericPrecision > 0) && (numericScale > 0))
            {
                _typeMapSQLToSQLCode.Add((int)SqlDbType.Decimal, String.Format("DECIMAL ({0},{1})", numericPrecision, numericScale));
                _typeMapSQLCodeToSQL.Add(String.Format("DECIMAL ({0},{1})", numericPrecision, numericScale), (int)SqlDbType.Decimal);
            }
            else if (numericPrecision > 0)
            {
                _typeMapSQLToSQLCode.Add((int)SqlDbType.Decimal, String.Format("DECIMAL ({0})", numericPrecision));
                _typeMapSQLCodeToSQL.Add(String.Format("DECIMAL ({0})", numericPrecision), (int)SqlDbType.Decimal);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)SqlDbType.Decimal, "DECIMAL");
                _typeMapSQLCodeToSQL.Add("DECIMAL", (int)SqlDbType.Decimal);
            }
            if (textLength > 0)
            {
                _typeMapSQLToSQLCode.Add((int)SqlDbType.NVarChar, String.Format("NVARCHAR ({0})", textLength));
                _typeMapSQLToSQLCode.Add((int)SqlDbType.VarChar, String.Format("VARCHAR ({0})", textLength));

                _typeMapSQLCodeToSQL.Add(String.Format("NVARCHAR ({0})", textLength), (int)SqlDbType.NVarChar);
                _typeMapSQLCodeToSQL.Add(String.Format("VARCHAR ({0})", textLength), (int)SqlDbType.VarChar);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)SqlDbType.NVarChar, "NVARCHAR");
                _typeMapSQLToSQLCode.Add((int)SqlDbType.VarChar, "VARCHAR");

                _typeMapSQLCodeToSQL.Add("NVARCHAR", (int)SqlDbType.NVarChar);
                _typeMapSQLCodeToSQL.Add("VARCHAR", (int)SqlDbType.VarChar);
            }
        }
        
        #endregion
    }
}
