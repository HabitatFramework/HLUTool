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
using System.Data.Common;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Windows;

namespace HLU.Data.Connection
{
    class DbOleDb : DbBase
    {
        #region Private Members

        private string _errorMessage;
        private OleDbConnectionStringBuilder _connStrBuilder;
        private OleDbConnection _connection;
        private OleDbCommand _command;
        private OleDbDataAdapter _adapter;
        private OleDbCommandBuilder _commandBuilder;
        private OleDbTransaction _transaction;
        private Dictionary<Type, OleDbDataAdapter> _adaptersDic = new Dictionary<Type, OleDbDataAdapter>();

        HLU.UI.View.Connection.ViewConnectOleDb _connWindow;
        HLU.UI.ViewModel.ViewModelConnectOleDb _connViewModel;

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

        public DbOleDb(ref string connString, ref string defaultSchema, ref bool promptPwd, string pwdMask,
            bool useCommandBuilder, bool useColumnNames, bool isUnicode, bool useTimeZone, uint textLength, 
            uint binaryLength, uint timePrecision, uint numericPrecision, uint numericScale)
            : base(ref connString, ref defaultSchema, ref promptPwd, pwdMask, useCommandBuilder, useColumnNames, 
            isUnicode, useTimeZone, textLength, binaryLength, timePrecision, numericPrecision, numericScale)
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
                _adapter = new OleDbDataAdapter(_command);
                _commandBuilder = new OleDbCommandBuilder(_adapter);

                connString = MaskPassword(_connStrBuilder, pwdMask);
                defaultSchema = DefaultSchema;
            }
            catch { throw; }
        }
        
        #endregion

        #region DbBase Members

        #region Public Static

        public static Backends GetBackend(OleDbConnection cn)
        {
            ConnectionState previousConnectionState = cn.State;

            if (String.IsNullOrEmpty(cn.Provider) &&
                (previousConnectionState != ConnectionState.Open)) cn.Open();

            string provider = cn.Provider.ToLower();

            if ((cn.State == ConnectionState.Open) &&
                (previousConnectionState != ConnectionState.Open)) cn.Close();

            if (provider.StartsWith("microsoft.jet.oledb"))
                return Backends.Access;
            else if (provider.StartsWith("sqloledb"))
                return Backends.SqlServer;
            else if (provider.StartsWith("oraoledb"))
                return Backends.Oracle;
            else if (provider.StartsWith("postgresql"))
                return Backends.PostgreSql;
            else if (provider.StartsWith("ibmdadb2"))
                return Backends.DB2;
            else
                return Backends.UndeterminedOleDb;
        }

        public static Backends GetBackend(OleDbConnectionStringBuilder connStrBuilder)
        {
            if ((connStrBuilder == null) || String.IsNullOrEmpty(connStrBuilder.ConnectionString))
                return Backends.UndeterminedOdbc;

            try
            {
                OleDbConnection cn = new OleDbConnection(connStrBuilder.ConnectionString);
                return GetBackend(cn);
            }
            catch { return Backends.UndeterminedOleDb; }
        }

        public static Backends GetBackend(string connString)
        {
            if (String.IsNullOrEmpty(connString)) return Backends.UndeterminedOleDb;

            try
            {
                OleDbConnection cn = new OleDbConnection(connString);
                return GetBackend(cn);
            }
            catch { return Backends.UndeterminedOleDb; }
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
                                           DataType = r.Field<int>("DATA_TYPE")
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
                                                              DbToSystemType(dbCol.DataType) == dsCol.DataType
                                                              select dbCol
                                                 where dbCols.Count() == 0
                                                 select QuoteIdentifier(dsCol.ColumnName) + " (" +
                                                 ((OleDbType)SystemToDbType(dsCol.DataType) + ")").ToString()).ToArray();
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
                return new OleDbCommand();
        }

        public override IDbDataAdapter CreateAdapter()
        {
            return new OleDbDataAdapter();
        }

        public override IDbDataAdapter CreateAdapter<T>(T table)
        {
            if (table == null) table = new T();

            OleDbDataAdapter adapter;

            if (!_adaptersDic.TryGetValue(typeof(T), out adapter))
            {
                adapter = new OleDbDataAdapter();

                DataColumn[] pk = table.PrimaryKey;
                if ((pk == null) || (pk.Length == 0)) return null;

                DataTableMapping tableMapping = new DataTableMapping();
                tableMapping.SourceTable = table.TableName;
                tableMapping.DataSetTable = table.TableName;

                List<OleDbParameter> deleteParams = new List<OleDbParameter>();
                List<OleDbParameter> insertParams = new List<OleDbParameter>();
                List<OleDbParameter> updateParams = new List<OleDbParameter>();
                List<OleDbParameter> updateParamsOrig = new List<OleDbParameter>();
                
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
                OleDbType isNullType = (OleDbType)isNullTypeInt;

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

                    if (c.AllowDBNull || ((_backend == Backends.Access) && !pk.Contains(c)))
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
                        (OleDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Original, false));

                    insColParamName = ParameterName(_parameterPrefixCurr, c.ColumnName, i + _startParamNo);
                    insertParams.Add(CreateParameter(insColParamName,
                        (OleDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Current, false));

                    updColParamName = ParameterName(_parameterPrefixCurr, c.ColumnName, i + _startParamNo);
                    updateParams.Add(CreateParameter(updColParamName,
                        (OleDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Current, false));

                    updOrigParamName = ParameterName(_parameterPrefixOrig, c.ColumnName, i + _startParamNo + columnCount + nullParamCount);
                    updateParamsOrig.Add(CreateParameter(updOrigParamName,
                        (OleDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Original, false));

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

                adapter.SelectCommand = new OleDbCommand();
                adapter.SelectCommand.CommandType = CommandType.Text;
                adapter.SelectCommand.Connection = _connection;
                adapter.SelectCommand.CommandText = String.Format("SELECT {0} FROM {1}", sbTargetList, tableName);

                if (!_useCommandBuilder)
                {
                    adapter.DeleteCommand = new OleDbCommand();
                    adapter.DeleteCommand.CommandType = CommandType.Text;
                    adapter.DeleteCommand.Connection = _connection;
                    adapter.DeleteCommand.CommandText = String.Format("DELETE FROM {0} WHERE {1}", tableName, sbWhereDel);
                    adapter.DeleteCommand.Parameters.AddRange(deleteParams.ToArray());

                    adapter.UpdateCommand = new OleDbCommand();
                    adapter.UpdateCommand.Connection = _connection;
                    adapter.UpdateCommand.CommandType = CommandType.Text;
                    adapter.UpdateCommand.CommandText =
                        String.Format("UPDATE {0} SET {1} WHERE {2}", tableName, sbUpdSetList, sbWhereUpd);
                    adapter.UpdateCommand.Parameters.AddRange(updateParams.ToArray());

                    adapter.InsertCommand = new OleDbCommand();
                    adapter.InsertCommand.CommandType = CommandType.Text;
                    adapter.InsertCommand.Connection = _connection;
                    adapter.InsertCommand.CommandText = String.Format("INSERT INTO {0} ({1}) VALUES ({2})",
                        tableName, sbTargetList, sbInsValues);
                    adapter.InsertCommand.Parameters.AddRange(insertParams.ToArray());
                }
                else
                {
                    OleDbCommandBuilder cmdBuilder = new OleDbCommandBuilder(adapter);
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

        private OleDbParameter CreateParameter(string name, OleDbType type, ParameterDirection direction, 
            string srcColumn, DataRowVersion srcVersion, bool nullMapping)
        {
            OleDbParameter param = new OleDbParameter(name, type);
            param.Direction = direction;
            param.SourceColumn = srcColumn;
            param.SourceVersion = srcVersion;
            param.SourceColumnNullMapping = nullMapping;

            param.IsNullable = nullMapping;
            
            return param;
        }

        private OleDbParameter CreateParameter(string name, object value, ParameterDirection direction,
            string srcColumn, DataRowVersion srcVersion, bool nullMapping)
        {
            OleDbParameter param = new OleDbParameter(name, value);
            param.Direction = direction;
            param.SourceColumn = srcColumn;
            param.SourceVersion = srcVersion;
            param.SourceColumnNullMapping = nullMapping;

            param.IsNullable = nullMapping;

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
                OleDbDataAdapter adapter = UpdateAdapter(table);
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
                    _adapter = new OleDbDataAdapter(_command);
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
                OleDbDataAdapter adapter = UpdateAdapter(table);
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
                    _adapter = new OleDbDataAdapter(_command);
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

        public override int Update<T>(T table, string insertCommand, string updateCommand, string deleteCommand)
        {
            ConnectionState previousConnectionState = _connection.State;
            if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();

            if ((table == null) || (table.Rows.Count == 0)) return 0;
            if (_adapter == null) return -1;

            try
            {
                if (!String.IsNullOrEmpty(insertCommand)) 
                    _adapter.InsertCommand = new OleDbCommand(insertCommand);
                if (!String.IsNullOrEmpty(updateCommand)) 
                    _adapter.UpdateCommand = new OleDbCommand(updateCommand);
                if (!String.IsNullOrEmpty(deleteCommand)) 
                    _adapter.DeleteCommand = new OleDbCommand(deleteCommand);

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
            if ((_connection.State & ConnectionState.Open) != ConnectionState.Open) _connection.Open();

            if ((table == null) || (table.Rows.Count == 0)) return 0;

            try
            {
                OleDbDataAdapter adapter = UpdateAdapter(table);

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

                OleDbDataAdapter adapter = UpdateAdapter(table);

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

                OleDbDataAdapter adapter = UpdateAdapter(table);

                if (adapter != null)
                    return adapter.Update(rows);
                else
                    return -1;
            }
            catch { return -1; }
            finally { if (previousConnectionState == ConnectionState.Closed) _connection.Close(); }
        }

        private OleDbDataAdapter UpdateAdapter<T>(T table) where T : DataTable, new()
        {
            if (table == null) return null;

            OleDbDataAdapter adapter;
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
            get { return String.Empty; }
        }

        #region Browse Connection

        protected override void BrowseConnection()
        {
            try
            {
                _connWindow = new HLU.UI.View.Connection.ViewConnectOleDb();
                
                if ((_connWindow.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));

                _connWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;

                // create ViewModel to which main window binds
                _connViewModel = new HLU.UI.ViewModel.ViewModelConnectOleDb();
                _connViewModel.DisplayName = "OleDb Connection";

                // when ViewModel asks to be closed, close window
                _connViewModel.RequestClose +=
                    new HLU.UI.ViewModel.ViewModelConnectOleDb.RequestCloseEventHandler(_connViewModel_RequestClose);

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
                MessageBox.Show("OleDb Server responded with an error:\n\n" + ex.Message,
                     "OleDb Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
                // FIX: 029 Update DateTime fields as strings not numbers
                // Fractions of a second can cause rounding differences when
                // comparing DateTime fields later in some databases so use
                // DateTime strings not numbers containing fractions.
                string s = valueType == typeof(DateTime) ? ((DateTime)value).ToString("s").Replace("T", " ") : value.ToString();
                //---------------------------------------------------------------------
                switch ((OleDbType)colType)
                {
                    case OleDbType.BSTR:
                    case OleDbType.Char:
                    case OleDbType.LongVarChar:
                    case OleDbType.LongVarWChar:
                    case OleDbType.VarChar:
                    case OleDbType.VarWChar:
                    case OleDbType.WChar:
                        if (s.Length == 0) return StringLiteralDelimiter + StringLiteralDelimiter;
                        if (!s.StartsWith(StringLiteralDelimiter)) s = StringLiteralDelimiter + s;
                        if (!s.EndsWith(StringLiteralDelimiter)) s += StringLiteralDelimiter;
                        return s;
                    case OleDbType.Date:
                    case OleDbType.DBDate:
                    case OleDbType.DBTime:
                    case OleDbType.DBTimeStamp:
                    case OleDbType.Filetime:
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
            string sTimeZone = useTimeZone ? " WITH TIME ZONE" : String.Empty;

            GetMetaData(typeof(OleDbType), _connection, _transaction);

            Dictionary<Type, int> typeMapSystemToSQLAdd = new Dictionary<Type, int>();
            typeMapSystemToSQLAdd.Add(typeof(Object), (int)OleDbType.Variant);
            typeMapSystemToSQLAdd.Add(typeof(Boolean), (int)OleDbType.Boolean);
            typeMapSystemToSQLAdd.Add(typeof(SByte), (int)OleDbType.Integer);
            typeMapSystemToSQLAdd.Add(typeof(Byte), (int)OleDbType.TinyInt);
            typeMapSystemToSQLAdd.Add(typeof(Int16), (int)OleDbType.SmallInt);
            typeMapSystemToSQLAdd.Add(typeof(UInt16), (int)OleDbType.UnsignedSmallInt);
            typeMapSystemToSQLAdd.Add(typeof(Int32), (int)OleDbType.Integer);
            typeMapSystemToSQLAdd.Add(typeof(UInt32), (int)OleDbType.UnsignedInt);
            typeMapSystemToSQLAdd.Add(typeof(Int64), (int)OleDbType.BigInt);
            typeMapSystemToSQLAdd.Add(typeof(UInt64), (int)OleDbType.UnsignedBigInt);
            typeMapSystemToSQLAdd.Add(typeof(Single), (int)OleDbType.Single);
            typeMapSystemToSQLAdd.Add(typeof(Double), (int)OleDbType.Double);
            typeMapSystemToSQLAdd.Add(typeof(Decimal), (int)OleDbType.Decimal);
            typeMapSystemToSQLAdd.Add(typeof(DateTime), (int)OleDbType.Date);
            typeMapSystemToSQLAdd.Add(typeof(TimeSpan), (int)OleDbType.DBTime);
            typeMapSystemToSQLAdd.Add(typeof(Byte[]), (int)OleDbType.Binary);
            typeMapSystemToSQLAdd.Add(typeof(Guid), (int)OleDbType.Guid);
            if (isUnicode)
            {
                typeMapSystemToSQLAdd.Add(typeof(Char), (int)OleDbType.WChar);
                typeMapSystemToSQLAdd.Add(typeof(String), (int)OleDbType.VarWChar);
                typeMapSystemToSQLAdd.Add(typeof(Char[]), (int)OleDbType.VarWChar);
            }
            else
            {
                typeMapSystemToSQLAdd.Add(typeof(Char), (int)OleDbType.Char);
                typeMapSystemToSQLAdd.Add(typeof(String), (int)OleDbType.VarChar);
                typeMapSystemToSQLAdd.Add(typeof(Char[]), (int)OleDbType.VarChar);
            }

            Dictionary<int, Type> typeMapSQLToSystemAdd = new Dictionary<int, Type>();
            typeMapSQLToSystemAdd.Add((int)OleDbType.BigInt, typeof(Int64));
            typeMapSQLToSystemAdd.Add((int)OleDbType.Binary, typeof(Byte[]));
            typeMapSQLToSystemAdd.Add((int)OleDbType.Boolean, typeof(Boolean));
            typeMapSQLToSystemAdd.Add((int)OleDbType.BSTR, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OleDbType.Char, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OleDbType.Currency, typeof(Decimal));
            typeMapSQLToSystemAdd.Add((int)OleDbType.Date, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)OleDbType.DBDate, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)OleDbType.DBTime, typeof(TimeSpan));
            typeMapSQLToSystemAdd.Add((int)OleDbType.DBTimeStamp, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)OleDbType.Decimal, typeof(Decimal));
            typeMapSQLToSystemAdd.Add((int)OleDbType.Double, typeof(Double));
            typeMapSQLToSystemAdd.Add((int)OleDbType.Error, typeof(Exception));
            typeMapSQLToSystemAdd.Add((int)OleDbType.Filetime, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)OleDbType.Guid, typeof(Guid));
            typeMapSQLToSystemAdd.Add((int)OleDbType.IDispatch, typeof(Object));
            typeMapSQLToSystemAdd.Add((int)OleDbType.Integer, typeof(Int32));
            typeMapSQLToSystemAdd.Add((int)OleDbType.IUnknown, typeof(Object));
            typeMapSQLToSystemAdd.Add((int)OleDbType.LongVarBinary, typeof(Byte[]));
            typeMapSQLToSystemAdd.Add((int)OleDbType.LongVarChar, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OleDbType.LongVarWChar, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OleDbType.Numeric, typeof(Decimal));
            typeMapSQLToSystemAdd.Add((int)OleDbType.PropVariant, typeof(Object));
            typeMapSQLToSystemAdd.Add((int)OleDbType.Single, typeof(Single));
            typeMapSQLToSystemAdd.Add((int)OleDbType.SmallInt, typeof(Int16));
            typeMapSQLToSystemAdd.Add((int)OleDbType.TinyInt, typeof(SByte));
            typeMapSQLToSystemAdd.Add((int)OleDbType.UnsignedBigInt, typeof(UInt64));
            typeMapSQLToSystemAdd.Add((int)OleDbType.UnsignedInt, typeof(UInt32));
            typeMapSQLToSystemAdd.Add((int)OleDbType.UnsignedSmallInt, typeof(UInt16));
            typeMapSQLToSystemAdd.Add((int)OleDbType.UnsignedTinyInt, typeof(Byte));
            typeMapSQLToSystemAdd.Add((int)OleDbType.VarBinary, typeof(Byte[]));
            typeMapSQLToSystemAdd.Add((int)OleDbType.VarChar, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OleDbType.Variant, typeof(Object));
            typeMapSQLToSystemAdd.Add((int)OleDbType.VarNumeric, typeof(Decimal));
            typeMapSQLToSystemAdd.Add((int)OleDbType.VarWChar, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OleDbType.WChar, typeof(String));

            Dictionary<string, int> sqlSynonymsAdd = new Dictionary<string, int>();
            sqlSynonymsAdd.Add("bit", (int)OleDbType.Boolean);
            sqlSynonymsAdd.Add("character", (int)OleDbType.Char);
            sqlSynonymsAdd.Add("char", (int)OleDbType.Char);
            sqlSynonymsAdd.Add("date", (int)OleDbType.DBDate);
            sqlSynonymsAdd.Add("time", (int)OleDbType.DBTime);
            sqlSynonymsAdd.Add("time with time zone", (int)OleDbType.DBTime);
            sqlSynonymsAdd.Add("timestamp", (int)OleDbType.DBTimeStamp);
            sqlSynonymsAdd.Add("timestamp with time zone", (int)OleDbType.DBTimeStamp);
            sqlSynonymsAdd.Add("decimal", (int)OleDbType.Decimal);
            sqlSynonymsAdd.Add("dec", (int)OleDbType.Decimal);
            sqlSynonymsAdd.Add("double precision", (int)OleDbType.Double);
            sqlSynonymsAdd.Add("character (36)", (int)OleDbType.Guid);
            sqlSynonymsAdd.Add("integer", (int)OleDbType.Integer);
            sqlSynonymsAdd.Add("int", (int)OleDbType.Integer);
            sqlSynonymsAdd.Add("numeric", (int)OleDbType.Numeric);
            sqlSynonymsAdd.Add("float", (int)OleDbType.Single);
            sqlSynonymsAdd.Add("real", (int)OleDbType.Single);
            sqlSynonymsAdd.Add("smallint", (int)OleDbType.SmallInt);
            sqlSynonymsAdd.Add("character varying", (int)OleDbType.VarChar);
            sqlSynonymsAdd.Add("char varying", (int)OleDbType.VarChar);
            sqlSynonymsAdd.Add("varchar", (int)OleDbType.VarChar);
            sqlSynonymsAdd.Add("bit varying", (int)OleDbType.Variant);
            sqlSynonymsAdd.Add("national character varying", (int)OleDbType.VarWChar);
            sqlSynonymsAdd.Add("national char varying", (int)OleDbType.VarWChar);
            sqlSynonymsAdd.Add("nchar varying", (int)OleDbType.VarWChar);
            sqlSynonymsAdd.Add("national character", (int)OleDbType.WChar);
            sqlSynonymsAdd.Add("national char", (int)OleDbType.WChar);
            sqlSynonymsAdd.Add("nchar", (int)OleDbType.WChar);

            foreach (KeyValuePair<Type, int> kv in typeMapSystemToSQLAdd)
            {
                if (!_typeMapSystemToSQL.ContainsKey(kv.Key))
                    _typeMapSystemToSQL.Add(kv.Key, kv.Value);
            }

            if (isUnicode)
            {
                ReplaceType(typeof(Char), (int)OleDbType.WChar, _typeMapSystemToSQL);
                ReplaceType(typeof(String), (int)OleDbType.VarWChar, _typeMapSystemToSQL);
                ReplaceType(typeof(Char[]), (int)OleDbType.VarWChar, _typeMapSystemToSQL);
            }
            else
            {
                ReplaceType(typeof(Char), (int)OleDbType.Char, _typeMapSystemToSQL);
                ReplaceType(typeof(String), (int)OleDbType.VarChar, _typeMapSystemToSQL);
                ReplaceType(typeof(Char[]), (int)OleDbType.VarChar, _typeMapSystemToSQL);
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
            
            _typeMapSQLToSQLCode.Add((int)OleDbType.BigInt, "INTEGER");
            _typeMapSQLToSQLCode.Add((int)OleDbType.Boolean, "BIT");
            _typeMapSQLToSQLCode.Add((int)OleDbType.Char, "CHARACTER");
            _typeMapSQLToSQLCode.Add((int)OleDbType.Date, "DATE");
            _typeMapSQLToSQLCode.Add((int)OleDbType.DBDate, "DATE");
            _typeMapSQLToSQLCode.Add((int)OleDbType.DBTimeStamp, "TIMESTAMP");
            _typeMapSQLToSQLCode.Add((int)OleDbType.Double, "DOUBLE PRECISION");
            _typeMapSQLToSQLCode.Add((int)OleDbType.Filetime, "DATE");
            _typeMapSQLToSQLCode.Add((int)OleDbType.Guid, "CHARACTER (36)");
            _typeMapSQLToSQLCode.Add((int)OleDbType.Integer, "INTEGER");
            _typeMapSQLToSQLCode.Add((int)OleDbType.Single, "REAL");
            _typeMapSQLToSQLCode.Add((int)OleDbType.SmallInt, "INTEGER");
            _typeMapSQLToSQLCode.Add((int)OleDbType.TinyInt, "SMALLINT");
            _typeMapSQLToSQLCode.Add((int)OleDbType.UnsignedBigInt, "INTEGER");
            _typeMapSQLToSQLCode.Add((int)OleDbType.UnsignedInt, "INTEGER");
            _typeMapSQLToSQLCode.Add((int)OleDbType.UnsignedSmallInt, "SMALLINT");
            _typeMapSQLToSQLCode.Add((int)OleDbType.UnsignedTinyInt, "SMALLINT");
            _typeMapSQLToSQLCode.Add((int)OleDbType.WChar, "NATIONAL CHARACTER");

            //_typeMapSQLCodeToSQL.Add("INTEGER", (int)OleDbType.BigInt);
            _typeMapSQLCodeToSQL.Add("BIT", (int)OleDbType.Boolean);
            _typeMapSQLCodeToSQL.Add("CHARACTER", (int)OleDbType.Char);
            _typeMapSQLCodeToSQL.Add("DATE", (int)OleDbType.Date);
            //_typeMapSQLCodeToSQL.Add("DATE", (int)OleDbType.DBDate);
            _typeMapSQLCodeToSQL.Add("TIMESTAMP", (int)OleDbType.DBTimeStamp);
            _typeMapSQLCodeToSQL.Add("DOUBLE PRECISION", (int)OleDbType.Double);
            //_typeMapSQLCodeToSQL.Add("DATE", (int)OleDbType.Filetime);
            _typeMapSQLCodeToSQL.Add("CHARACTER (36)", (int)OleDbType.Guid);
            _typeMapSQLCodeToSQL.Add("INTEGER", (int)OleDbType.Integer);
            _typeMapSQLCodeToSQL.Add("REAL", (int)OleDbType.Single);
            //_typeMapSQLCodeToSQL.Add("INTEGER", (int)OleDbType.SmallInt);
            _typeMapSQLCodeToSQL.Add("SMALLINT", (int)OleDbType.TinyInt);
            //_typeMapSQLCodeToSQL.Add("INTEGER", (int)OleDbType.UnsignedBigInt);
            //_typeMapSQLCodeToSQL.Add("INTEGER", (int)OleDbType.UnsignedInt);
            //_typeMapSQLCodeToSQL.Add("SMALLINT", (int)OleDbType.UnsignedSmallInt);
            //_typeMapSQLCodeToSQL.Add("SMALLINT", (int)OleDbType.UnsignedTinyInt);
            _typeMapSQLCodeToSQL.Add("NATIONAL CHARACTER", (int)OleDbType.WChar);

            if (textLength > 0)
            {
                _typeMapSQLToSQLCode.Add((int)OleDbType.BSTR, String.Format("NATIONAL CHARACTER VARYING ({0})", textLength));
                _typeMapSQLToSQLCode.Add((int)OleDbType.LongVarChar, String.Format("CHARACTER VARYING ({0})", textLength));
                _typeMapSQLToSQLCode.Add((int)OleDbType.LongVarWChar, String.Format("NATIONAL CHARACTER VARYING ({0})", textLength));
                _typeMapSQLToSQLCode.Add((int)OleDbType.VarChar, String.Format("CHARACTER VARYING ({0})", textLength));
                _typeMapSQLToSQLCode.Add((int)OleDbType.VarWChar, String.Format("NATIONAL CHARACTER VARYING ({0})", textLength));

                //_typeMapSQLCodeToSQL.Add(String.Format("NATIONAL CHARACTER VARYING ({0})", textLength), (int)OleDbType.BSTR);
                //_typeMapSQLCodeToSQL.Add(String.Format("CHARACTER VARYING ({0})", textLength), (int)OleDbType.LongVarChar);
                //_typeMapSQLCodeToSQL.Add(String.Format("NATIONAL CHARACTER VARYING ({0})", textLength), (int)OleDbType.LongVarWChar);
                _typeMapSQLCodeToSQL.Add(String.Format("CHARACTER VARYING ({0})", textLength), (int)OleDbType.VarChar);
                _typeMapSQLCodeToSQL.Add(String.Format("NATIONAL CHARACTER VARYING ({0})", textLength), (int)OleDbType.VarWChar);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)OleDbType.BSTR, "NATIONAL CHARACTER VARYING");
                _typeMapSQLToSQLCode.Add((int)OleDbType.LongVarChar, "CHARACTER VARYING");
                _typeMapSQLToSQLCode.Add((int)OleDbType.LongVarWChar, "NATIONAL CHARACTER VARYING");
                _typeMapSQLToSQLCode.Add((int)OleDbType.VarChar, "CHARACTER VARYING");
                _typeMapSQLToSQLCode.Add((int)OleDbType.VarWChar, "NATIONAL CHARACTER VARYING");

                //_typeMapSQLCodeToSQL.Add("NATIONAL CHARACTER VARYING", (int)OleDbType.BSTR);
                //_typeMapSQLCodeToSQL.Add("CHARACTER VARYING", (int)OleDbType.LongVarChar);
                //_typeMapSQLCodeToSQL.Add("NATIONAL CHARACTER VARYING", (int)OleDbType.LongVarWChar);
                _typeMapSQLCodeToSQL.Add("CHARACTER VARYING", (int)OleDbType.VarChar);
                _typeMapSQLCodeToSQL.Add("NATIONAL CHARACTER VARYING", (int)OleDbType.VarWChar);
            }
            if ((numericPrecision > 0) && (numericScale > 0))
            {
                _typeMapSQLToSQLCode.Add((int)OleDbType.Currency, String.Format("DECIMAL ({0},{1})", numericPrecision, numericScale));
                //_typeMapSQLCodeToSQL.Add(String.Format("DECIMAL ({0},{1})", numericPrecision, numericScale), (int)OleDbType.Currency);
                _typeMapSQLToSQLCode.Add((int)OleDbType.Decimal, String.Format("DECIMAL ({0},{1})", numericPrecision, numericScale));
                _typeMapSQLCodeToSQL.Add(String.Format("DECIMAL ({0},{1})", numericPrecision, numericScale), (int)OleDbType.Decimal);
                _typeMapSQLToSQLCode.Add((int)OleDbType.Numeric, String.Format("NUMERIC ({0},{1})", numericPrecision, numericScale));
                _typeMapSQLCodeToSQL.Add(String.Format("NUMERIC ({0},{1})", numericPrecision, numericScale), (int)OleDbType.Numeric);
                _typeMapSQLToSQLCode.Add((int)OleDbType.VarNumeric, String.Format("NUMERIC ({0},{1})", numericPrecision, numericScale));
                //_typeMapSQLCodeToSQL.Add(String.Format("NUMERIC ({0},{1})", numericPrecision, numericScale), (int)OleDbType.VarNumeric);
            }
            else if (numericPrecision > 0)
            {
                _typeMapSQLToSQLCode.Add((int)OleDbType.Currency, String.Format("DECIMAL ({0})", numericPrecision));
                _typeMapSQLToSQLCode.Add((int)OleDbType.Decimal, String.Format("DECIMAL ({0})", numericPrecision));
                _typeMapSQLToSQLCode.Add((int)OleDbType.Numeric, String.Format("NUMERIC ({0})", numericPrecision));
                _typeMapSQLToSQLCode.Add((int)OleDbType.VarNumeric, String.Format("NUMERIC ({0})", numericPrecision));

                //_typeMapSQLCodeToSQL.Add(String.Format("DECIMAL ({0})", numericPrecision), (int)OleDbType.Currency);
                _typeMapSQLCodeToSQL.Add(String.Format("DECIMAL ({0})", numericPrecision), (int)OleDbType.Decimal);
                _typeMapSQLCodeToSQL.Add(String.Format("NUMERIC ({0})", numericPrecision), (int)OleDbType.Numeric);
                //_typeMapSQLCodeToSQL.Add(String.Format("NUMERIC ({0})", numericPrecision), (int)OleDbType.VarNumeric);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)OleDbType.Currency, "DECIMAL");
                _typeMapSQLToSQLCode.Add((int)OleDbType.Decimal, "DECIMAL");
                _typeMapSQLToSQLCode.Add((int)OleDbType.Numeric, "NUMERIC");
                _typeMapSQLToSQLCode.Add((int)OleDbType.VarNumeric, "NUMERIC");

                //_typeMapSQLCodeToSQL.Add("DECIMAL", (int)OleDbType.Currency);
                _typeMapSQLCodeToSQL.Add("DECIMAL", (int)OleDbType.Decimal);
                _typeMapSQLCodeToSQL.Add("NUMERIC", (int)OleDbType.Numeric);
                //_typeMapSQLCodeToSQL.Add("NUMERIC", (int)OleDbType.VarNumeric);
            }
            if (timePrecision > 0)
            {
                _typeMapSQLToSQLCode.Add((int)OleDbType.DBTime, String.Format("TIME ({0}){1}", timePrecision, sTimeZone));
                _typeMapSQLCodeToSQL.Add(String.Format("TIME ({0}){1}", timePrecision, sTimeZone), (int)OleDbType.DBTime);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)OleDbType.DBTime, String.Format("TIME{0}", sTimeZone));
                _typeMapSQLCodeToSQL.Add(String.Format("TIME{0}", sTimeZone), (int)OleDbType.DBTime);
            }
            if (binaryLength > 0)
            {
                _typeMapSQLToSQLCode.Add((int)OleDbType.Binary, String.Format("BIT VARYING ({0})", binaryLength));
                _typeMapSQLToSQLCode.Add((int)OleDbType.IDispatch, String.Format("BIT VARYING ({0})", binaryLength));
                _typeMapSQLToSQLCode.Add((int)OleDbType.LongVarBinary, String.Format("BIT VARYING ({0})", binaryLength));
                _typeMapSQLToSQLCode.Add((int)OleDbType.PropVariant, String.Format("BIT VARYING ({0})", binaryLength));
                _typeMapSQLToSQLCode.Add((int)OleDbType.VarBinary, String.Format("BIT VARYING ({0})", binaryLength));
                _typeMapSQLToSQLCode.Add((int)OleDbType.Variant, String.Format("BIT VARYING ({0})", binaryLength));

                //_typeMapSQLCodeToSQL.Add(String.Format("BIT VARYING ({0})", binaryLength), (int)OleDbType.Binary);
                //_typeMapSQLCodeToSQL.Add(String.Format("BIT VARYING ({0})", binaryLength), (int)OleDbType.IDispatch);
                //_typeMapSQLCodeToSQL.Add(String.Format("BIT VARYING ({0})", binaryLength), (int)OleDbType.LongVarBinary);
                //_typeMapSQLCodeToSQL.Add(String.Format("BIT VARYING ({0})", binaryLength), (int)OleDbType.PropVariant);
                _typeMapSQLCodeToSQL.Add(String.Format("BIT VARYING ({0})", binaryLength), (int)OleDbType.VarBinary);
                //_typeMapSQLCodeToSQL.Add(String.Format("BIT VARYING ({0})", binaryLength), (int)OleDbType.Variant);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)OleDbType.Binary, "BIT VARYING");
                _typeMapSQLToSQLCode.Add((int)OleDbType.IDispatch, "BIT VARYING");
                _typeMapSQLToSQLCode.Add((int)OleDbType.LongVarBinary, "BIT VARYING");
                _typeMapSQLToSQLCode.Add((int)OleDbType.PropVariant, "BIT VARYING");
                _typeMapSQLToSQLCode.Add((int)OleDbType.VarBinary, "BIT VARYING");
                _typeMapSQLToSQLCode.Add((int)OleDbType.Variant, "BIT VARYING");

                //_typeMapSQLCodeToSQL.Add("BIT VARYING", (int)OleDbType.Binary);
                //_typeMapSQLCodeToSQL.Add("BIT VARYING", (int)OleDbType.IDispatch);
                //_typeMapSQLCodeToSQL.Add("BIT VARYING", (int)OleDbType.LongVarBinary);
                //_typeMapSQLCodeToSQL.Add("BIT VARYING", (int)OleDbType.PropVariant);
                _typeMapSQLCodeToSQL.Add("BIT VARYING", (int)OleDbType.VarBinary);
                //_typeMapSQLCodeToSQL.Add("BIT VARYING", (int)OleDbType.Variant);
            }
        }

        private void SetDefaults()
        {
            _backend = GetBackend(_connStrBuilder);

            switch (_backend)
            {
                case Backends.Access:
                    _quotePrefix = "[";
                    _quoteSuffix = "]";
                    _stringLiteralDelimiter = "\"";
                    _dateLiteralPrefix = "#";
                    _dateLiteralSuffix = "#";
                    _wildcardSingleMatch = "_";
                    _wildcardManyMatch = "%";
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
