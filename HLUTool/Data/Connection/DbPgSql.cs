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
using System.Linq;
using System.Windows;
using Npgsql;
using NpgsqlTypes;
using System.Text;

namespace HLU.Data.Connection
{
    class DbPgSql : DbBase
    {
        #region Private Members

        private string _errorMessage;
        private NpgsqlConnectionStringBuilder _connStrBuilder;
        private NpgsqlConnection _connection;
        private NpgsqlCommand _command;
        private NpgsqlDataAdapter _adapter;
        private NpgsqlCommandBuilder _commandBuilder;
        private NpgsqlTransaction _transaction;
        private string _encoding;
        private Dictionary<Type, NpgsqlDataAdapter> _adaptersDic = new Dictionary<Type, NpgsqlDataAdapter>();

        private HLU.UI.View.Connection.ViewConnectPgSql _connWindow;
        private HLU.UI.ViewModel.ViewModelConnectPgSql _connViewModel;
       
        #endregion

        #region Constructor

        public DbPgSql(ref string connString, ref string defaultSchema, ref bool promptPwd, string pwdMask,
            bool useCommandBuilder, bool useColumnNames, bool isUnicode, bool useTimeZone, uint textLength, 
            uint binaryLength, uint timePrecision, uint numericPrecision, uint numericScale)
            : base(ref connString, ref defaultSchema, ref promptPwd, pwdMask, useCommandBuilder, useColumnNames, 
            isUnicode, useTimeZone, textLength, binaryLength, timePrecision, numericPrecision, numericScale)
        {
            if (String.IsNullOrEmpty(ConnectionString)) throw (new Exception("No connection string"));

            try
            {
                Login("User Name", ConnectionString, ref promptPwd, ref _connStrBuilder, ref _connection);

                PopulateTypeMaps(IsUnicode, UseTimeZone, TextLength, BinaryLength,
                    TimePrecision, NumericPrecision, NumericScale);
                SetPgClientEncoding();

                _command = _connection.CreateCommand();
                _adapter = new NpgsqlDataAdapter(_command);
                _commandBuilder = new NpgsqlCommandBuilder(_adapter);

                connString = MaskPassword(_connStrBuilder, pwdMask);
                defaultSchema = DefaultSchema;
            }
            catch { throw; }
        }
        #endregion

        #region DbBase Members

        #region Public Members

        public override Backends Backend { get { return Backends.PostgreSql; } }

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
                return new NpgsqlCommand();
        }

        public override IDbDataAdapter CreateAdapter()
        {
            return new NpgsqlDataAdapter();
        }

        public override IDbDataAdapter CreateAdapter<T>(T table)
        {
            if (table == null) table = new T();

            NpgsqlDataAdapter adapter;

            if (!_adaptersDic.TryGetValue(typeof(T), out adapter))
            {
                adapter = new NpgsqlDataAdapter();

                DataColumn[] pk = table.PrimaryKey;
                if ((pk == null) || (pk.Length == 0)) return null;

                DataTableMapping tableMapping = new DataTableMapping();
                tableMapping.SourceTable = table.TableName; // "Table";
                tableMapping.DataSetTable = table.TableName; // "Exports";

                List<NpgsqlParameter> deleteParams = new List<NpgsqlParameter>();
                List<NpgsqlParameter> insertParams = new List<NpgsqlParameter>();
                List<NpgsqlParameter> updateParams = new List<NpgsqlParameter>();
                List<NpgsqlParameter> updateParamsOrig = new List<NpgsqlParameter>();

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
                NpgsqlDbType isNullType = (NpgsqlDbType)isNullTypeInt;

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
                        (NpgsqlDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Original, false));

                    insColParamName = ParameterName(_parameterPrefixCurr, c.ColumnName, i + _startParamNo);
                    insertParams.Add(CreateParameter(insColParamName,
                        (NpgsqlDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Current, false));

                    updColParamName = ParameterName(_parameterPrefixCurr, c.ColumnName, i + _startParamNo);
                    updateParams.Add(CreateParameter(updColParamName,
                        (NpgsqlDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Current, false));

                    updOrigParamName = ParameterName(_parameterPrefixOrig, c.ColumnName, i + _startParamNo + columnCount + nullParamCount);
                    updateParamsOrig.Add(CreateParameter(updOrigParamName, 
                        (NpgsqlDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Original, false));

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

                adapter.SelectCommand = new NpgsqlCommand();
                adapter.SelectCommand.CommandType = CommandType.Text;
                adapter.SelectCommand.Connection = _connection;
                adapter.SelectCommand.CommandText = String.Format("SELECT {0} FROM {1}", sbTargetList, tableName);

                if (!_useCommandBuilder)
                {
                    adapter.DeleteCommand = new NpgsqlCommand();
                    adapter.DeleteCommand.CommandType = CommandType.Text;
                    adapter.DeleteCommand.Connection = _connection;
                    adapter.DeleteCommand.CommandText = String.Format("DELETE FROM {0} WHERE {1}", tableName, sbWhereDel);
                    adapter.DeleteCommand.Parameters.AddRange(deleteParams.ToArray());

                    adapter.UpdateCommand = new NpgsqlCommand();
                    adapter.UpdateCommand.Connection = _connection;
                    adapter.UpdateCommand.CommandType = CommandType.Text;
                    adapter.UpdateCommand.CommandText =
                        String.Format("UPDATE {0} SET {1} WHERE {2}", tableName, sbUpdSetList, sbWhereUpd);
                    adapter.UpdateCommand.Parameters.AddRange(updateParams.ToArray());

                    adapter.InsertCommand = new NpgsqlCommand();
                    adapter.InsertCommand.CommandType = CommandType.Text;
                    adapter.InsertCommand.Connection = _connection;
                    adapter.InsertCommand.CommandText = String.Format("INSERT INTO {0} ({1}) VALUES ({2})",
                        tableName, sbTargetList, sbInsValues);
                    adapter.InsertCommand.Parameters.AddRange(insertParams.ToArray());
                }
                else
                {
                    NpgsqlCommandBuilder cmdBuilder = new NpgsqlCommandBuilder(adapter);
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

        private NpgsqlParameter CreateParameter(string name, NpgsqlDbType type, ParameterDirection direction,
            string srcColumn, DataRowVersion srcVersion, bool nullMapping)
        {
            NpgsqlParameter param = new NpgsqlParameter(name, type);
            param.Direction = direction;
            param.SourceColumn = srcColumn;
            param.SourceVersion = srcVersion;
            param.SourceColumnNullMapping = nullMapping;
            return param;
        }

        private NpgsqlParameter CreateParameter(string name, object value, ParameterDirection direction,
            string srcColumn, DataRowVersion srcVersion, bool nullMapping)
        {
            NpgsqlParameter param = new NpgsqlParameter(name, value);
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
                NpgsqlDataAdapter adapter = UpdateAdapter(table);
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
                    _adapter = new NpgsqlDataAdapter(_command);
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
                NpgsqlDataAdapter adapter = UpdateAdapter(table);
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
                    _adapter = new NpgsqlDataAdapter(_command);
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

                if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();

                if (_transaction != null)
                {
                    _command.Transaction = _transaction;
                    _commandBuilder.RefreshSchema();
                }
                
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

                if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();
                
                if (_transaction != null) _command.Transaction = _transaction;
                _commandBuilder.RefreshSchema();
                
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

                if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();
                
                if (_transaction != null) _command.Transaction = _transaction;
                _commandBuilder.RefreshSchema();
                
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

                if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();

                if (_transaction != null) _command.Transaction = _transaction;
                _commandBuilder.RefreshSchema();

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
                    _adapter.InsertCommand = new NpgsqlCommand(insertCommand);
                if (!String.IsNullOrEmpty(updateCommand))
                    _adapter.UpdateCommand = new NpgsqlCommand(updateCommand);
                if (!String.IsNullOrEmpty(deleteCommand))
                    _adapter.DeleteCommand = new NpgsqlCommand(deleteCommand);

                return _adapter.Update(table);
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
            if ((_connection.State & ConnectionState.Open) != ConnectionState.Open)
                _connection.Open();

            if ((table == null) || (table.Rows.Count == 0)) return 0;

            try
            {
                NpgsqlDataAdapter adapter = UpdateAdapter(table);

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
            if ((_connection.State & ConnectionState.Open) != ConnectionState.Open)
                _connection.Open();

            if ((dataSet == null) || String.IsNullOrEmpty(sourceTable) ||
                !dataSet.Tables.Contains(sourceTable)) return 0;

            try
            {
                DataTable table = dataSet.Tables[sourceTable];

                NpgsqlDataAdapter adapter = UpdateAdapter(table);

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
            if ((_connection.State & ConnectionState.Open) != ConnectionState.Open)
                _connection.Open();

            if ((rows == null) || (rows.Length == 0)) return 0;

            try
            {
                T table = (T)rows[0].Table;

                NpgsqlDataAdapter adapter = UpdateAdapter(table);

                if (adapter != null)
                    return adapter.Update(rows);
                else
                    return -1;
            }
            catch { return -1; }
            finally { if (previousConnectionState == ConnectionState.Closed) _connection.Close(); }
        }

        private NpgsqlDataAdapter UpdateAdapter<T>(T table) where T : DataTable, new()
        {
            if (table == null) return null;

            NpgsqlDataAdapter adapter;
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
                _connWindow = new HLU.UI.View.Connection.ViewConnectPgSql();
                if ((_connWindow.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));
                _connWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;

                // create ViewModel to which main window binds
                _connViewModel = new HLU.UI.ViewModel.ViewModelConnectPgSql();
                _connViewModel.DisplayName = "PostgreSQL Connection";

                // when ViewModel asks to be closed, close window
                _connViewModel.RequestClose +=
                    new HLU.UI.ViewModel.ViewModelConnectPgSql.RequestCloseEventHandler(_connViewModel_RequestClose);

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
                MessageBox.Show("PostgreSQL Server responded with an error:\n\n" + ex.Message,
                     "PostgreSQL Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }

        protected void _connViewModel_RequestClose(string connString, string encoding, 
            string defaultSchema, string errorMsg)
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
                if (!String.IsNullOrEmpty(encoding)) _encoding = encoding;
            }
        }

        #endregion

        #endregion
        
        #endregion

        #region SQLBuilder Members
        
        #region Public Members
        
        public override string QuotePrefix { get { return "\""; } }

        public override string QuoteSuffix { get { return "\""; } }

        public override string StringLiteralDelimiter { get { return "'"; } }

        public override string DateLiteralPrefix { get { return "'"; } }

        public override string DateLiteralSuffix { get { return "'"; } }

        public override string WildcardSingleMatch { get { return "_"; } }

        public override string WildcardManyMatch { get { return "%"; } }

        public override string ConcatenateOperator { get { return "||"; } }

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
                switch ((NpgsqlDbType)colType)
                {
                    case NpgsqlDbType.Char:
                    case NpgsqlDbType.Name:
                    case NpgsqlDbType.Text:
                    case NpgsqlDbType.Varchar:
                    case NpgsqlDbType.Xml:
                        if (s.Length == 0) return StringLiteralDelimiter + StringLiteralDelimiter;
                        if (!s.StartsWith(StringLiteralDelimiter)) s = StringLiteralDelimiter + s;
                        if (!s.EndsWith(StringLiteralDelimiter)) s += StringLiteralDelimiter;
                        return s;
                    case NpgsqlDbType.Date:
                    case NpgsqlDbType.Time:
                    case NpgsqlDbType.Timestamp:
                    case NpgsqlDbType.TimestampTZ:
                    case NpgsqlDbType.TimeTZ:
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

        private void PopulateTypeMaps(bool isUnicode, bool useTimeZone, uint textLength, 
            uint binaryLength, uint timePrecision, uint numericPrecision, uint numericScale)
        {
            string timeZoneSuffix = useTimeZone ? "tz" : "";

            GetMetaData(typeof(NpgsqlDbType), _connection, _transaction);

            Dictionary<Type, int> typeMapSystemToSQLAdd = new Dictionary<Type, int>();
            typeMapSystemToSQLAdd.Add(typeof(Boolean), (int)NpgsqlDbType.Boolean);
            typeMapSystemToSQLAdd.Add(typeof(Byte), (int) NpgsqlDbType.Smallint);
            typeMapSystemToSQLAdd.Add(typeof(Char), (int) NpgsqlDbType.Char);
            typeMapSystemToSQLAdd.Add(typeof(DateTime), (int)(useTimeZone ? NpgsqlDbType.TimestampTZ : NpgsqlDbType.Timestamp));
            typeMapSystemToSQLAdd.Add(typeof(TimeSpan), (int)NpgsqlDbType.Interval);
            typeMapSystemToSQLAdd.Add(typeof(Decimal), (int) NpgsqlDbType.Numeric);
            typeMapSystemToSQLAdd.Add(typeof(Double), (int) NpgsqlDbType.Numeric);
            typeMapSystemToSQLAdd.Add(typeof(Int16), (int) NpgsqlDbType.Smallint);
            typeMapSystemToSQLAdd.Add(typeof(Int32), (int) NpgsqlDbType.Integer);
            typeMapSystemToSQLAdd.Add(typeof(Int64), (int) NpgsqlDbType.Bigint);
            typeMapSystemToSQLAdd.Add(typeof(Object), (int) NpgsqlDbType.Bytea);
            typeMapSystemToSQLAdd.Add(typeof(SByte), (int) NpgsqlDbType.Smallint);
            typeMapSystemToSQLAdd.Add(typeof(Single), (int) NpgsqlDbType.Real);
            typeMapSystemToSQLAdd.Add(typeof(String), (int) NpgsqlDbType.Varchar);
            typeMapSystemToSQLAdd.Add(typeof(UInt16), (int) NpgsqlDbType.Smallint);
            typeMapSystemToSQLAdd.Add(typeof(UInt32), (int) NpgsqlDbType.Integer);
            typeMapSystemToSQLAdd.Add(typeof(UInt64), (int) NpgsqlDbType.Bigint);
            typeMapSystemToSQLAdd.Add(typeof(Byte[]), (int) NpgsqlDbType.Bytea);
            typeMapSystemToSQLAdd.Add(typeof(Char[]), (int) NpgsqlDbType.Varchar);
            typeMapSystemToSQLAdd.Add(typeof(Guid), (int)NpgsqlDbType.Uuid);

            Dictionary<int, Type> typeMapSQLToSystemAdd = new Dictionary<int, Type>();
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Bigint, typeof(Int64));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Bit, typeof(Boolean));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Boolean, typeof(Boolean));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Bytea, typeof(Byte[]));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Char, typeof(Char));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Date, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Double, typeof(Double));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Integer, typeof(Int32));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Money, typeof(Decimal));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Numeric, typeof(Decimal));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Real, typeof(Single));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Smallint, typeof(Int16));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Text, typeof(String));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Interval, typeof(TimeSpan));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Time, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Timestamp, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.TimestampTZ, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Uuid, typeof(Guid));
            typeMapSQLToSystemAdd.Add((int)NpgsqlDbType.Varchar, typeof(String));

            Dictionary<string, int> sqlSynonymsAdd = new Dictionary<string, int>();
            sqlSynonymsAdd.Add("bigint", (int)NpgsqlDbType.Bigint);
            sqlSynonymsAdd.Add("bigserial", (int)NpgsqlDbType.Bigint);
            sqlSynonymsAdd.Add("bit varying", (int)NpgsqlDbType.Bytea); // ??
            sqlSynonymsAdd.Add("bit", (int)NpgsqlDbType.Bit);
            sqlSynonymsAdd.Add("bool", (int)NpgsqlDbType.Boolean);
            sqlSynonymsAdd.Add("boolean", (int)NpgsqlDbType.Boolean);
            sqlSynonymsAdd.Add("box", (int)NpgsqlDbType.Box);
            sqlSynonymsAdd.Add("bytea", (int)NpgsqlDbType.Bytea);
            sqlSynonymsAdd.Add("char", (int)NpgsqlDbType.Char);
            sqlSynonymsAdd.Add("character varying", (int)NpgsqlDbType.Varchar);
            sqlSynonymsAdd.Add("character", (int)NpgsqlDbType.Char);
            sqlSynonymsAdd.Add("cidr", (int)NpgsqlDbType.Varchar); // ??
            sqlSynonymsAdd.Add("circle", (int)NpgsqlDbType.Circle);
            sqlSynonymsAdd.Add("date", (int)NpgsqlDbType.Date);
            sqlSynonymsAdd.Add("double precision", (int)NpgsqlDbType.Double);
            sqlSynonymsAdd.Add("float4", (int)NpgsqlDbType.Real);
            sqlSynonymsAdd.Add("float8", (int)NpgsqlDbType.Double);
            sqlSynonymsAdd.Add("inet", (int)NpgsqlDbType.Inet);
            sqlSynonymsAdd.Add("int", (int)NpgsqlDbType.Integer);
            sqlSynonymsAdd.Add("int2", (int)NpgsqlDbType.Smallint);
            sqlSynonymsAdd.Add("int4", (int)NpgsqlDbType.Integer);
            sqlSynonymsAdd.Add("int8", (int)NpgsqlDbType.Bigint);
            sqlSynonymsAdd.Add("integer", (int)NpgsqlDbType.Integer);
            sqlSynonymsAdd.Add("interval", (int)NpgsqlDbType.Interval);
            sqlSynonymsAdd.Add("line", (int)NpgsqlDbType.Line);
            sqlSynonymsAdd.Add("lseg", (int)NpgsqlDbType.LSeg);
            sqlSynonymsAdd.Add("macaddr", (int)NpgsqlDbType.Varchar); // ??
            sqlSynonymsAdd.Add("money", (int)NpgsqlDbType.Money);
            sqlSynonymsAdd.Add("numeric", (int)NpgsqlDbType.Numeric);
            sqlSynonymsAdd.Add("path", (int)NpgsqlDbType.Path);
            sqlSynonymsAdd.Add("point", (int)NpgsqlDbType.Point);
            sqlSynonymsAdd.Add("polygon", (int)NpgsqlDbType.Polygon);
            sqlSynonymsAdd.Add("real", (int)NpgsqlDbType.Real);
            sqlSynonymsAdd.Add("serial", (int)NpgsqlDbType.Integer);
            sqlSynonymsAdd.Add("serial4", (int)NpgsqlDbType.Integer);
            sqlSynonymsAdd.Add("serial8", (int)NpgsqlDbType.Bigint);
            sqlSynonymsAdd.Add("smallint", (int)NpgsqlDbType.Smallint);
            sqlSynonymsAdd.Add("text", (int)NpgsqlDbType.Text);
            sqlSynonymsAdd.Add("time with time zone", (int)NpgsqlDbType.TimeTZ);
            sqlSynonymsAdd.Add("time without time zone", (int)NpgsqlDbType.Time);
            sqlSynonymsAdd.Add("time", (int)NpgsqlDbType.Time);
            sqlSynonymsAdd.Add("timestamp with time zone", (int)NpgsqlDbType.TimestampTZ);
            sqlSynonymsAdd.Add("timestamp without time zone", (int)NpgsqlDbType.Timestamp);
            sqlSynonymsAdd.Add("timestamp", (int)NpgsqlDbType.Timestamp);
            sqlSynonymsAdd.Add("timestamptz", (int)NpgsqlDbType.TimestampTZ);
            sqlSynonymsAdd.Add("timetz", (int)NpgsqlDbType.TimeTZ);
            sqlSynonymsAdd.Add("uuid", (int)NpgsqlDbType.Uuid);
            sqlSynonymsAdd.Add("varbit", (int)NpgsqlDbType.Bytea); // ??
            sqlSynonymsAdd.Add("varchar", (int)NpgsqlDbType.Varchar);
            sqlSynonymsAdd.Add("xml", (int)NpgsqlDbType.Xml);

            foreach (KeyValuePair<Type, int> kv in typeMapSystemToSQLAdd)
            {
                if (!_typeMapSystemToSQL.ContainsKey(kv.Key))
                    _typeMapSystemToSQL.Add(kv.Key, kv.Value);
            }

            ReplaceType(typeof(DateTime), (int)(useTimeZone ? NpgsqlDbType.TimestampTZ : 
                NpgsqlDbType.Timestamp), _typeMapSystemToSQL);

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
            
            _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Bigint, "bigint");
            _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Bit, "boolean");
            _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Boolean, "boolean");
            _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Bytea, "bytea");
            _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Char, "character");
            _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Date, "date");
            _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Double, "double precision");
            _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Integer, "integer");
            _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Money, "money");

            _typeMapSQLCodeToSQL.Add("bigint", (int)NpgsqlDbType.Bigint);
            //_typeMapSQLCodeToSQL.Add("boolean", (int)NpgsqlDbType.Bit);
            _typeMapSQLCodeToSQL.Add("boolean", (int)NpgsqlDbType.Boolean);
            _typeMapSQLCodeToSQL.Add("bytea", (int)NpgsqlDbType.Bytea);
            _typeMapSQLCodeToSQL.Add("character", (int)NpgsqlDbType.Char);
            _typeMapSQLCodeToSQL.Add("date", (int)NpgsqlDbType.Date);
            _typeMapSQLCodeToSQL.Add("integer", (int)NpgsqlDbType.Integer);
            _typeMapSQLCodeToSQL.Add("money", (int)NpgsqlDbType.Money);
            
            if ((numericPrecision > 0) && (numericScale > 0))
            {
                _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Numeric, String.Format("numeric ({0},{1})", numericPrecision, numericScale));
                _typeMapSQLCodeToSQL.Add(String.Format("numeric ({0},{1})", numericPrecision, numericScale), (int)NpgsqlDbType.Numeric);
            }
            else if (numericPrecision > 0)
            {
                _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Numeric, String.Format("numeric ({0})", numericPrecision));
                _typeMapSQLCodeToSQL.Add(String.Format("numeric ({0})", numericPrecision), (int)NpgsqlDbType.Numeric);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Numeric, "numeric");
                _typeMapSQLCodeToSQL.Add("numeric", (int)NpgsqlDbType.Numeric);
            }
            _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Real, "real");
            _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Smallint, "smallint");
            _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Text, "text");

            _typeMapSQLCodeToSQL.Add("real", (int)NpgsqlDbType.Real);
            _typeMapSQLCodeToSQL.Add("smallint", (int)NpgsqlDbType.Smallint);
            _typeMapSQLCodeToSQL.Add("text", (int)NpgsqlDbType.Text);

            if (timePrecision > 0)
            {
                _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Time, String.Format("time{0} ({1})", timeZoneSuffix, timePrecision));
                _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Timestamp, String.Format("timestamp ({0})", timePrecision));
                _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.TimestampTZ, String.Format("timestamptz ({0})", timePrecision));
                
                _typeMapSQLCodeToSQL.Add(String.Format("time{0} ({1})", timeZoneSuffix, timePrecision), (int)NpgsqlDbType.Time);
                _typeMapSQLCodeToSQL.Add(String.Format("timestamp ({0})", timePrecision), (int)NpgsqlDbType.Timestamp);
                _typeMapSQLCodeToSQL.Add(String.Format("timestamptz ({0})", timePrecision), (int)NpgsqlDbType.TimestampTZ);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Time, String.Format("time{0}", timeZoneSuffix));
                _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Timestamp, "timestamp");
                _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.TimestampTZ, "timestamptz");
                
                _typeMapSQLCodeToSQL.Add(String.Format("time{0}", timeZoneSuffix), (int)NpgsqlDbType.Time);
                _typeMapSQLCodeToSQL.Add("timestamp", (int)NpgsqlDbType.Timestamp);
                _typeMapSQLCodeToSQL.Add("timestamptz", (int)NpgsqlDbType.TimestampTZ);
            }
            if (textLength > 0)
            {
                _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Varchar, String.Format("varchar ({0})", textLength));
                _typeMapSQLCodeToSQL.Add(String.Format("varchar ({0})", textLength), (int)NpgsqlDbType.Varchar);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)NpgsqlDbType.Varchar, "varchar");
                _typeMapSQLCodeToSQL.Add("varchar", (int)NpgsqlDbType.Varchar);
            }
        }

        private void SetPgClientEncoding()
        {
            if (String.IsNullOrEmpty(_encoding)) return;

            ConnectionState previousConnectionState = _connection.State;

            try
            {
                if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();

                _command.CommandText = "SET client_encoding TO " + ("'" + _encoding.Trim() + "'").Replace("''", "'");
                _command.ExecuteNonQuery();
            }
            catch { }
            finally { if (previousConnectionState != ConnectionState.Open) _connection.Clone(); }
        }
        
        #endregion
    }
}
