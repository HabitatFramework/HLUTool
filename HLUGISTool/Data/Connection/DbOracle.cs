using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Windows;
using Oracle.DataAccess.Client;
using System.Text;

namespace HLU.Data.Connection
{
    class DbOracle : DbBase
    {
        #region Private Members
        
        private string _errorMessage;
        private OracleConnectionStringBuilder _connStrBuilder;
        private OracleConnection _connection;
        private OracleCommand _command;
        private OracleDataAdapter _adapter;
        private OracleCommandBuilder _commandBuilder;
        private OracleTransaction _transaction;
        private Dictionary<Type, OracleDataAdapter> _adaptersDic = new Dictionary<Type, OracleDataAdapter>();

        private HLU.UI.View.Connection.ViewConnectOracle _connWindow;
        private HLU.UI.ViewModel.ViewModelConnectOracle _connViewModel;

        #endregion

        #region Constructor

        public DbOracle(ref string connString, ref string defaultSchema, ref bool promptPwd,
            string pwdMask, bool useCommandBuilder, bool useColumnNames, bool isUnicode, 
            bool useTimeZone, uint textLength, uint binaryLength, uint timePrecision, 
            uint numericPrecision, uint numericScale)
            : base(ref connString, ref defaultSchema, ref promptPwd, pwdMask, useCommandBuilder, 
            useColumnNames, isUnicode, useTimeZone, textLength, binaryLength, timePrecision, 
            numericPrecision, numericScale)
        {
            if (String.IsNullOrEmpty(ConnectionString)) throw (new Exception("No connection string"));

            try
            {
                Login("User ID", ConnectionString, ref promptPwd, ref _connStrBuilder, ref _connection);

                PopulateTypeMaps(IsUnicode, UseTimeZone, TextLength, BinaryLength,
                    TimePrecision, NumericPrecision, NumericScale);

                _command = _connection.CreateCommand();
                _adapter = new OracleDataAdapter(_command);
                _commandBuilder = new OracleCommandBuilder(_adapter);

                connString = MaskPassword(_connStrBuilder, pwdMask);
                defaultSchema = DefaultSchema;

                _startParamNo = 0;
                _parameterPrefixCurr = "cur_";
                _parameterPrefixNull = "ind_";
                _parameterPrefixOrig = "ori_";
            }
            catch { throw; }
        }
        
        #endregion

        #region Public Static

        public static Dictionary<string, string> GetConnectionStrings(DataTable dataSources)
        {
            return (from r in dataSources.AsEnumerable()
                    select BuildConnectionString(r.Field<string>("InstanceName"), 
                    r.Field<string>("Protocol"), r.Field<string>("ServerName"), 
                    r.Field<string>("Port"), r.Field<string>("ServiceName"))
                    ).ToDictionary(kv => kv.Key, kv => kv.Value);
        }

        public static KeyValuePair<string, string> BuildConnectionString(string instanceName, string protocol,
            string serverName, string port, string serviceName)
        {
            return new KeyValuePair<string, string>(instanceName, 
                String.Format("(DESCRIPTION = (ADDRESS_LIST = (ADDRESS = (PROTOCOL = {0})" +
                "(HOST = {1})(PORT = {2}))) (CONNECT_DATA = (SERVICE_NAME = {3})))", protocol,
                serverName, port, serviceName));
        }

        public static string GetUserId(string userIDstring)
        {
            if (!String.IsNullOrEmpty(userIDstring))
            {
                if (userIDstring.StartsWith("\"") && userIDstring.EndsWith("\""))
                    userIDstring = userIDstring.Remove(userIDstring.Length - 1, 1).Remove(0, 1);
                else
                    userIDstring = userIDstring.ToUpper();
            }
            return userIDstring;
        }

        public static string GetUserId(DbConnectionStringBuilder connStrBuilder)
        {
            if (connStrBuilder == null) return String.Empty;
            object userID;
            if (connStrBuilder.TryGetValue("USER ID", out userID))
                return GetUserId(userID.ToString());
            else
                return String.Empty;
        }

        #endregion

        #region DbBase Members

        #region Public Members

        public override Backends Backend { get { return Backends.Oracle; } }

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
                                       let numScale = r["SCALE"]
                                       where tableName == t.TableName
                                       select new
                                       {
                                           TableName = tableName,
                                           ColumnName = r.Field<string>("COLUMN_NAME"),
                                           ColumnLength = r.Field<decimal>("LENGTHINCHARS"),
                                           NumericScale = numScale != DBNull.Value ? r.Field<decimal>("SCALE") : 0,
                                           DataType = r.Field<string>("DATATYPE")
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
                                                              TypeMatch(SystemDataType(dbCol.DataType), dbCol.ColumnLength, 
                                                                        dbCol.NumericScale, dsCol.DataType, dsCol.MaxLength)
                                                              select dbCol
                                                 where dbCols.Count() == 0
                                                 select QuoteIdentifier(dsCol.ColumnName) + " (" +
                                                 ((OracleDbType)SystemToDbType(dsCol.DataType) + ")").ToString()).ToArray();
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

        private bool TypeMatch(Type dbColSysType, decimal dbColLength, decimal dbColNumScale, Type dsColType, int dsColLength)
        {
            TypeCode dbColSysTypeCode = Type.GetTypeCode(dbColSysType);
            TypeCode dsColTypeCode = Type.GetTypeCode(dsColType);

            if (dbColSysTypeCode == dsColTypeCode) return true;

            TypeCode[] floatingPoint = new TypeCode[] { TypeCode.Decimal, TypeCode.Double };

            switch (dsColTypeCode)
            {
                case TypeCode.Boolean:
                    return (dbColSysTypeCode == TypeCode.String && dbColLength == 5);
                case TypeCode.Char:
                    return (dsColLength == -1 || dbColLength <= dsColLength) && dbColSysTypeCode == TypeCode.String;
                case TypeCode.Decimal:
                    return dbColNumScale > 0 && Array.IndexOf(floatingPoint, dbColSysTypeCode) != 1;
                case TypeCode.Double:
                    return Array.IndexOf(floatingPoint, dbColSysTypeCode) != -1;
                case TypeCode.Int16:
                case TypeCode.UInt16:
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                    return dbColNumScale == 0 && dbColSysTypeCode == TypeCode.Decimal;
                case TypeCode.Object:
                case TypeCode.Single:
                    return dbColNumScale > 0 && Array.IndexOf(floatingPoint, dbColSysTypeCode) != 1;
            }

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
                return new OracleCommand();
        }

        public override IDbDataAdapter CreateAdapter()
        {
            return new OracleDataAdapter();
        }

        public override IDbDataAdapter CreateAdapter<T>(T table)
        {
            if (table == null) table = new T();

            OracleDataAdapter adapter;

            if (!_adaptersDic.TryGetValue(typeof(T), out adapter))
            {
                adapter = new OracleDataAdapter();

                DataColumn[] pk = table.PrimaryKey;
                if ((pk == null) || (pk.Length == 0)) return null;

                DataTableMapping tableMapping = new DataTableMapping();
                tableMapping.SourceTable = table.TableName; // "Table";
                tableMapping.DataSetTable = table.TableName; // "Exports";

                List<OracleParameter> deleteParams = new List<OracleParameter>();
                List<OracleParameter> insertParams = new List<OracleParameter>();
                List<OracleParameter> updateParams = new List<OracleParameter>();
                List<OracleParameter> updateParamsOrig = new List<OracleParameter>();

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
                OracleDbType isNullType = (OracleDbType)isNullTypeInt;

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
                        (OracleDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Original, false));

                    insColParamName = ParameterName(_parameterPrefixCurr, c.ColumnName, i + _startParamNo);
                    insertParams.Add(CreateParameter(insColParamName,
                        (OracleDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Current, false));

                    updColParamName = ParameterName(_parameterPrefixCurr, c.ColumnName, i + _startParamNo);
                    updateParams.Add(CreateParameter(updColParamName,
                        (OracleDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Current, false));

                    updOrigParamName = ParameterName(_parameterPrefixOrig, c.ColumnName, i + _startParamNo + columnCount + nullParamCount);
                    updateParamsOrig.Add(CreateParameter(updOrigParamName,
                        (OracleDbType)colType, ParameterDirection.Input, c.ColumnName, DataRowVersion.Original, false));

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

                adapter.SelectCommand = new OracleCommand();
                adapter.SelectCommand.CommandType = CommandType.Text;
                adapter.SelectCommand.Connection = _connection;
                adapter.SelectCommand.CommandText = String.Format("SELECT {0} FROM {1}", sbTargetList, tableName);

                if (!_useCommandBuilder)
                {
                    adapter.DeleteCommand = new OracleCommand();
                    adapter.DeleteCommand.CommandType = CommandType.Text;
                    adapter.DeleteCommand.Connection = _connection;
                    adapter.DeleteCommand.CommandText = String.Format("DELETE FROM {0} WHERE {1}", tableName, sbWhereDel);
                    adapter.DeleteCommand.Parameters.AddRange(deleteParams.ToArray());

                    adapter.UpdateCommand = new OracleCommand();
                    adapter.UpdateCommand.Connection = _connection;
                    adapter.UpdateCommand.CommandType = CommandType.Text;
                    adapter.UpdateCommand.CommandText =
                        String.Format("UPDATE {0} SET {1} WHERE {2}", tableName, sbUpdSetList, sbWhereUpd);
                    adapter.UpdateCommand.Parameters.AddRange(updateParams.ToArray());

                    adapter.InsertCommand = new OracleCommand();
                    adapter.InsertCommand.CommandType = CommandType.Text;
                    adapter.InsertCommand.Connection = _connection;
                    adapter.InsertCommand.CommandText = String.Format("INSERT INTO {0} ({1}) VALUES ({2})",
                        tableName, sbTargetList, sbInsValues);
                    adapter.InsertCommand.Parameters.AddRange(insertParams.ToArray());
                }
                else
                {
                    OracleCommandBuilder cmdBuilder = new OracleCommandBuilder(adapter);
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

        private OracleParameter CreateParameter(string name, OracleDbType type, ParameterDirection direction,
            string srcColumn, DataRowVersion srcVersion, bool nullMapping)
        {
            OracleParameter param = new OracleParameter(name, type);
            param.Direction = direction;
            param.SourceColumn = srcColumn;
            param.SourceVersion = srcVersion;
            param.SourceColumnNullMapping = nullMapping;
            return param;
        }

        private OracleParameter CreateParameter(string name, object value, ParameterDirection direction,
            string srcColumn, DataRowVersion srcVersion, bool nullMapping)
        {
            OracleParameter param = new OracleParameter(name, value);
            param.Direction = direction;
            param.SourceColumn = srcColumn;
            param.SourceVersion = srcVersion;
            param.SourceColumnNullMapping = nullMapping;
            return param;
        }

        protected override string ParameterName(string prefix, string columnName, int paramNo)
        {
            return ParameterPrefix + prefix +
                (columnName.Length < 7 ? columnName : columnName.Substring(0, 7)) + "_p" + paramNo;
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
                OracleDataAdapter adapter = UpdateAdapter(table);
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
                    _adapter = new OracleDataAdapter(_command);

                    // potential data loss with Oracle types: NUMBER, DATE, all Timestamp types, and INTERVAL DAY TO SECOND
                    // _adapter.SafeMapping.Add("ColumnName", typeof(byte[])); _adapter.SafeMapping.Add("ColumnName", typeof(string));
                    // _adapter.SafeMapping.Add("*", typeof(byte[])); _adapter.SafeMapping.Add("*", typeof(string));
                    
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
                OracleDataAdapter adapter = UpdateAdapter(table);
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
                    _adapter = new OracleDataAdapter(_command);

                    // potential data loss with Oracle types: NUMBER, DATE, all Timestamp types, and INTERVAL DAY TO SECOND
                    // _adapter.SafeMapping.Add("ColumnName", typeof(byte[])); _adapter.SafeMapping.Add("ColumnName", typeof(string));
                    // _adapter.SafeMapping.Add("*", typeof(byte[])); _adapter.SafeMapping.Add("*", typeof(string));
                    
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
                bool newCommand = false;

                if ((newCommand = !String.IsNullOrEmpty(insertCommand)))
                    _adapter.InsertCommand = new OracleCommand(insertCommand);
                if ((newCommand |= !String.IsNullOrEmpty(updateCommand)))
                    _adapter.UpdateCommand = new OracleCommand(updateCommand);
                if ((newCommand |= !String.IsNullOrEmpty(deleteCommand)))
                    _adapter.DeleteCommand = new OracleCommand(deleteCommand);

                if (newCommand)
                {
                    // potential data loss with Oracle types: NUMBER, DATE, all Timestamp types, and INTERVAL DAY TO SECOND
                    // _adapter.SafeMapping.Add("ColumnName", typeof(byte[])); _adapter.SafeMapping.Add("ColumnName", typeof(string));
                    // _adapter.SafeMapping.Add("*", typeof(byte[])); _adapter.SafeMapping.Add("*", typeof(string));
                }

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
                OracleDataAdapter adapter = UpdateAdapter(table);

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

                OracleDataAdapter adapter = UpdateAdapter(table);

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

                OracleDataAdapter adapter = UpdateAdapter(table);

                if (adapter != null)
                    return adapter.Update(rows);
                else
                    return -1;
            }
            catch { return -1; }
            finally { if (previousConnectionState == ConnectionState.Closed) _connection.Close(); }
        }

        private OracleDataAdapter UpdateAdapter<T>(T table) where T : DataTable, new()
        {
            if (table == null) return null;

            OracleDataAdapter adapter;
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
            get { return ":"; }
        }

        #region Browse Connection
        
        protected override void BrowseConnection()
        {
            try
            {
                _connWindow = new HLU.UI.View.Connection.ViewConnectOracle();
                if ((_connWindow.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));
                _connWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;

                // create ViewModel to which main window binds
                _connViewModel = new HLU.UI.ViewModel.ViewModelConnectOracle();
                _connViewModel.DisplayName = "Oracle Connection";

                // when ViewModel asks to be closed, close window
                _connViewModel.RequestClose +=
                    new HLU.UI.ViewModel.ViewModelConnectOracle.RequestCloseEventHandler(_connViewModel_RequestClose);

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
                MessageBox.Show("Oracle Server responded with an error:\n\n" + ex.Message,
                     "Oracle Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
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

        public override string QuotePrefix { get { return "\""; } }

        public override string QuoteSuffix { get { return "\""; } }

        public override string StringLiteralDelimiter { get { return "'"; } }

        public override string DateLiteralPrefix { get { return "'"; } }

        public override string DateLiteralSuffix { get { return "'"; } }

        public override string WildcardSingleMatch { get { return "_"; } }

        public override string WildcardManyMatch { get { return "%"; } }

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
                switch ((Oracle.DataAccess.Client.OracleDbType)colType)
                {
                    case OracleDbType.Char:
                    case OracleDbType.Clob:
                    case OracleDbType.NChar:
                    case OracleDbType.NClob:
                    case OracleDbType.NVarchar2:
                    case OracleDbType.Varchar2:
                        if (s.Length == 0) return StringLiteralDelimiter + StringLiteralDelimiter;
                        if (!s.StartsWith(StringLiteralDelimiter)) s = StringLiteralDelimiter + s;
                        if (!s.EndsWith(StringLiteralDelimiter)) s += StringLiteralDelimiter;
                        return s;
                    case OracleDbType.Date:
                    case OracleDbType.IntervalDS:
                    case OracleDbType.IntervalYM:
                    case OracleDbType.TimeStamp:
                    case OracleDbType.TimeStampLTZ:
                    case OracleDbType.TimeStampTZ:
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
            GetMetaData(typeof(OracleDbType), _connection, _transaction);

            Dictionary<Type, int> typeMapSystemToSQLAdd = new Dictionary<Type, int>();
            typeMapSystemToSQLAdd.Add(typeof(Byte), (int)OracleDbType.Byte);
            typeMapSystemToSQLAdd.Add(typeof(Char), (int)OracleDbType.Char);
            typeMapSystemToSQLAdd.Add(typeof(DateTime), (int)(useTimeZone ? OracleDbType.TimeStampTZ : OracleDbType.TimeStamp));
            typeMapSystemToSQLAdd.Add(typeof(TimeSpan), (int)OracleDbType.IntervalDS);
            typeMapSystemToSQLAdd.Add(typeof(Decimal), (int)OracleDbType.Decimal);
            typeMapSystemToSQLAdd.Add(typeof(Double), (int)OracleDbType.Double);
            typeMapSystemToSQLAdd.Add(typeof(Int16), (int)OracleDbType.Int16);
            typeMapSystemToSQLAdd.Add(typeof(Int32), (int)OracleDbType.Int32);
            typeMapSystemToSQLAdd.Add(typeof(Int64), (int)OracleDbType.Int64);
            typeMapSystemToSQLAdd.Add(typeof(Object), (int)OracleDbType.Blob);
            typeMapSystemToSQLAdd.Add(typeof(SByte), (int)OracleDbType.Int16);
            typeMapSystemToSQLAdd.Add(typeof(Single), (int)OracleDbType.Single);
            typeMapSystemToSQLAdd.Add(typeof(String), (int)(isUnicode ? OracleDbType.NVarchar2 : OracleDbType.Varchar2));
            typeMapSystemToSQLAdd.Add(typeof(UInt16), (int)OracleDbType.Int16);
            typeMapSystemToSQLAdd.Add(typeof(UInt32), (int)OracleDbType.Int32);
            typeMapSystemToSQLAdd.Add(typeof(UInt64), (int)OracleDbType.Int64);
            typeMapSystemToSQLAdd.Add(typeof(Byte[]), (int)OracleDbType.Blob);
            typeMapSystemToSQLAdd.Add(typeof(Char[]), (int)OracleDbType.Clob);
            typeMapSystemToSQLAdd.Add(typeof(Guid), (int)OracleDbType.NVarchar2);
            typeMapSystemToSQLAdd.Add(typeof(Boolean), (int)OracleDbType.Char);

            Dictionary<int, Type> typeMapSQLToSystemAdd = new Dictionary<int, Type>();
            typeMapSQLToSystemAdd.Add((int)OracleDbType.BFile, typeof(object));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.Blob, typeof(object));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.Byte, typeof(Byte));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.Char, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.Clob, typeof(object));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.Date, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.Decimal, typeof(Decimal));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.Double, typeof(Double));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.Int16, typeof(Int16));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.Int32, typeof(Int32));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.Int64, typeof(Int64));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.IntervalDS, typeof(TimeSpan));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.IntervalYM, typeof(Int64));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.Long, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.LongRaw, typeof(Byte[]));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.NChar, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.NClob, typeof(object));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.NVarchar2, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.Object, typeof(object));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.Raw, typeof(Byte[]));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.RefCursor, typeof(object));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.Single, typeof(Single));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.TimeStamp, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.TimeStampLTZ, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.TimeStampTZ, typeof(DateTime));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.Varchar2, typeof(String));
            typeMapSQLToSystemAdd.Add((int)OracleDbType.XmlType, typeof(string));

            Dictionary<string, int> sqlSynonymsAdd = new Dictionary<string, int>();
            sqlSynonymsAdd.Add("bfile", (int)OracleDbType.BFile);
            sqlSynonymsAdd.Add("blob", (int)OracleDbType.Blob);
            sqlSynonymsAdd.Add("character", (int)OracleDbType.Char);
            sqlSynonymsAdd.Add("char", (int)OracleDbType.Char);
            sqlSynonymsAdd.Add("clob", (int)OracleDbType.Clob);
            sqlSynonymsAdd.Add("date", (int)OracleDbType.Date);
            sqlSynonymsAdd.Add("double precision", (int)OracleDbType.Double);
            sqlSynonymsAdd.Add("binary_double", (int)OracleDbType.Double);
            sqlSynonymsAdd.Add("float", (int)OracleDbType.Decimal);
            sqlSynonymsAdd.Add("binary_float", (int)OracleDbType.Decimal);
            sqlSynonymsAdd.Add("real", (int)OracleDbType.Decimal);
            sqlSynonymsAdd.Add("smallint", (int)OracleDbType.Int16);
            sqlSynonymsAdd.Add("integer", (int)OracleDbType.Int32);
            sqlSynonymsAdd.Add("int", (int)OracleDbType.Int32);
            sqlSynonymsAdd.Add("pls_integer", (int)OracleDbType.Int32);
            sqlSynonymsAdd.Add("binary_integer", (int)OracleDbType.Int32);
            sqlSynonymsAdd.Add("interval day to second", (int)OracleDbType.IntervalDS);
            sqlSynonymsAdd.Add("interval year to month", (int)OracleDbType.IntervalYM);
            sqlSynonymsAdd.Add("long raw", (int)OracleDbType.LongRaw);
            sqlSynonymsAdd.Add("long", (int)OracleDbType.Varchar2);
            sqlSynonymsAdd.Add("long varchar", (int)OracleDbType.Varchar2);
            sqlSynonymsAdd.Add("nchar", (int)OracleDbType.NChar);
            sqlSynonymsAdd.Add("national char", (int)OracleDbType.NChar);
            sqlSynonymsAdd.Add("national character", (int)OracleDbType.NChar);
            sqlSynonymsAdd.Add("nclob", (int)OracleDbType.NClob);
            sqlSynonymsAdd.Add("number", (int)OracleDbType.Decimal);
            sqlSynonymsAdd.Add("numeric", (int)OracleDbType.Decimal);
            sqlSynonymsAdd.Add("decimal", (int)OracleDbType.Decimal);
            sqlSynonymsAdd.Add("dec", (int)OracleDbType.Decimal);
            sqlSynonymsAdd.Add("national char varying", (int)OracleDbType.NVarchar2);
            sqlSynonymsAdd.Add("national character varying", (int)OracleDbType.NVarchar2);
            sqlSynonymsAdd.Add("nchar varying", (int)OracleDbType.NVarchar2);
            sqlSynonymsAdd.Add("nvarchar2", (int)OracleDbType.NVarchar2);
            sqlSynonymsAdd.Add("raw", (int)OracleDbType.Raw);
            sqlSynonymsAdd.Add("rowid", (int)OracleDbType.Int64);
            sqlSynonymsAdd.Add("urowid", (int)OracleDbType.Int64);
            sqlSynonymsAdd.Add("timestamp", (int)OracleDbType.TimeStamp);
            sqlSynonymsAdd.Add("timestamp with local time zone", (int)OracleDbType.TimeStampLTZ);
            sqlSynonymsAdd.Add("timestamp with time zone", (int)OracleDbType.TimeStampTZ);
            sqlSynonymsAdd.Add("character varying", (int)OracleDbType.Varchar2);
            sqlSynonymsAdd.Add("char varying", (int)OracleDbType.Varchar2);
            sqlSynonymsAdd.Add("varchar", (int)OracleDbType.Varchar2);
            sqlSynonymsAdd.Add("varchar2", (int)OracleDbType.Varchar2);

            foreach (KeyValuePair<Type, int> kv in typeMapSystemToSQLAdd)
            {
                if (!_typeMapSystemToSQL.ContainsKey(kv.Key))
                    _typeMapSystemToSQL.Add(kv.Key, kv.Value);
            }

            ReplaceType(typeof(DateTime), (int)(useTimeZone ? OracleDbType.TimeStampTZ : 
                OracleDbType.TimeStamp), _typeMapSystemToSQL);
            ReplaceType(typeof(String), (int)(isUnicode ? OracleDbType.NVarchar2 : 
                OracleDbType.Varchar2), _typeMapSystemToSQL);

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

            _typeMapSQLToSQLCode.Add((int)OracleDbType.BFile, "BFILE");
            _typeMapSQLToSQLCode.Add((int)OracleDbType.Blob, "BLOB");
            _typeMapSQLToSQLCode.Add((int)OracleDbType.Byte, "SMALLINT");
            _typeMapSQLToSQLCode.Add((int)OracleDbType.Char, "CHAR");
            _typeMapSQLToSQLCode.Add((int)OracleDbType.Clob, "CLOB");
            _typeMapSQLToSQLCode.Add((int)OracleDbType.Date, "DATE");
            _typeMapSQLToSQLCode.Add((int)OracleDbType.Double, "DOUBLE PRECISION");
            _typeMapSQLToSQLCode.Add((int)OracleDbType.Int16, "SMALLINT");
            _typeMapSQLToSQLCode.Add((int)OracleDbType.Int32, "INTEGER");
            _typeMapSQLToSQLCode.Add((int)OracleDbType.LongRaw, "LONG RAW");
            _typeMapSQLToSQLCode.Add((int)OracleDbType.NClob, "NCLOB");
            _typeMapSQLToSQLCode.Add((int)OracleDbType.NVarchar2, String.Format("NVARCHAR2 ({0})", textLength));
            _typeMapSQLToSQLCode.Add((int)OracleDbType.Raw, binaryLength > 0 ? String.Format("RAW ({0})", binaryLength) : "RAW");
            _typeMapSQLCodeToSQL.Add(binaryLength > 0 ? String.Format("RAW ({0})", binaryLength) : "RAW", (int)OracleDbType.Raw);

            _typeMapSQLCodeToSQL.Add("BFILE", (int)OracleDbType.BFile);
            _typeMapSQLCodeToSQL.Add("BLOB", (int)OracleDbType.Blob);
            _typeMapSQLCodeToSQL.Add("SMALLINT", (int)OracleDbType.Byte);
            _typeMapSQLCodeToSQL.Add("CHAR", (int)OracleDbType.Char);
            _typeMapSQLCodeToSQL.Add("CLOB", (int)OracleDbType.Clob);
            _typeMapSQLCodeToSQL.Add("DATE", (int)OracleDbType.Date);
            _typeMapSQLCodeToSQL.Add("DOUBLE PRECISION", (int)OracleDbType.Double);
            //_typeMapSQLCodeToSQL.Add("SMALLINT", (int)OracleDbType.Int16);
            _typeMapSQLCodeToSQL.Add("INTEGER", (int)OracleDbType.Int32);
            _typeMapSQLCodeToSQL.Add("LONG RAW", (int)OracleDbType.LongRaw);
            _typeMapSQLCodeToSQL.Add("NCLOB", (int)OracleDbType.NClob);
            _typeMapSQLCodeToSQL.Add(String.Format("NVARCHAR2 ({0})", textLength), (int)OracleDbType.NVarchar2);

            if ((numericPrecision > 0) && (numericScale > 0))
            {
                _typeMapSQLToSQLCode.Add((int)OracleDbType.Decimal, String.Format("NUMBER ({0},{1})", numericPrecision, numericScale));
                _typeMapSQLCodeToSQL.Add(String.Format("NUMBER ({0},{1})", numericPrecision, numericScale), (int)OracleDbType.Decimal);
            }
            else if (numericPrecision > 0)
            {
                //_typeMapSQLToSQLCode.Add((int)OracleDbType.Decimal, String.Format("FLOAT ({0})", numericPrecision));
                _typeMapSQLCodeToSQL.Add(String.Format("FLOAT ({0})", numericPrecision), (int)OracleDbType.Decimal);

                _typeMapSQLToSQLCode.Add((int)OracleDbType.Decimal, String.Format("NUMBER ({0})", numericPrecision));
                _typeMapSQLCodeToSQL.Add(String.Format("NUMBER ({0})", numericPrecision), (int)OracleDbType.Decimal);
            }
            else
            {
                //_typeMapSQLToSQLCode.Add((int)OracleDbType.Decimal, "FLOAT");
                _typeMapSQLCodeToSQL.Add("FLOAT", (int)OracleDbType.Decimal);

                _typeMapSQLToSQLCode.Add((int)OracleDbType.Decimal, "NUMBER");
                _typeMapSQLCodeToSQL.Add("NUMBER", (int)OracleDbType.Decimal);
            }
            if (timePrecision > 0)
            {
                _typeMapSQLToSQLCode.Add((int)OracleDbType.IntervalDS,
                    String.Format("INTERVAL DAY ({0}) TO SECOND ({1})", timePrecision, numericPrecision));
                _typeMapSQLToSQLCode.Add((int)OracleDbType.IntervalYM,
                    String.Format("INTERVAL YEAR ({0}) TO MONTH", timePrecision));
                _typeMapSQLToSQLCode.Add((int)OracleDbType.TimeStamp, String.Format("TIMESTAMP ({0})", timePrecision));
                _typeMapSQLToSQLCode.Add((int)OracleDbType.TimeStampLTZ,
                    String.Format("TIMESTAMP ({0}) WITH LOCAL TIME ZONE", timePrecision));
                _typeMapSQLToSQLCode.Add((int)OracleDbType.TimeStampTZ,
                    String.Format("TIMESTAMP ({0}) WITH TIME ZONE", timePrecision));

                _typeMapSQLCodeToSQL.Add(String.Format("INTERVAL DAY ({0}) TO SECOND ({1})", timePrecision,
                    numericPrecision), (int)OracleDbType.IntervalDS);
                _typeMapSQLCodeToSQL.Add(String.Format("INTERVAL YEAR ({0}) TO MONTH", timePrecision),
                    (int)OracleDbType.IntervalYM);
                _typeMapSQLCodeToSQL.Add(String.Format("TIMESTAMP ({0})", timePrecision), (int)OracleDbType.TimeStamp);
                _typeMapSQLCodeToSQL.Add(String.Format("TIMESTAMP ({0}) WITH LOCAL TIME ZONE", timePrecision),
                    (int)OracleDbType.TimeStampLTZ);
                _typeMapSQLCodeToSQL.Add(String.Format("TIMESTAMP ({0}) WITH TIME ZONE", timePrecision),
                    (int)OracleDbType.TimeStampTZ);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)OracleDbType.IntervalDS, numericPrecision > 0 ?
                    String.Format("INTERVAL DAY TO SECOND ({0})", numericPrecision) : "INTERVAL DAY TO SECOND");
                _typeMapSQLToSQLCode.Add((int)OracleDbType.IntervalYM, "INTERVAL YEAR TO MONTH");
                _typeMapSQLToSQLCode.Add((int)OracleDbType.TimeStamp, "TIMESTAMP");
                _typeMapSQLToSQLCode.Add((int)OracleDbType.TimeStampLTZ, "TIMESTAMP WITH LOCAL TIME ZONE");
                _typeMapSQLToSQLCode.Add((int)OracleDbType.TimeStampTZ, "TIMESTAMP WITH TIME ZONE");

                _typeMapSQLCodeToSQL.Add(numericPrecision > 0 ? String.Format("INTERVAL DAY TO SECOND ({0})",
                    numericPrecision) : "INTERVAL DAY TO SECOND", (int)OracleDbType.IntervalDS);
                _typeMapSQLCodeToSQL.Add("INTERVAL YEAR TO MONTH", (int)OracleDbType.IntervalYM);
                _typeMapSQLCodeToSQL.Add("TIMESTAMP", (int)OracleDbType.TimeStamp);
                _typeMapSQLCodeToSQL.Add("TIMESTAMP WITH LOCAL TIME ZONE", (int)OracleDbType.TimeStampLTZ);
                _typeMapSQLCodeToSQL.Add("TIMESTAMP WITH TIME ZONE", (int)OracleDbType.TimeStampTZ);
            }
            if (textLength > 0)
            {
                _typeMapSQLToSQLCode.Add((int)OracleDbType.NChar, String.Format("NCHAR({0})", textLength));
                _typeMapSQLToSQLCode.Add((int)OracleDbType.Varchar2, String.Format("VARCHAR2({0})", textLength));

                _typeMapSQLCodeToSQL.Add(String.Format("VARCHAR2 ({0})", textLength), (int)OracleDbType.Varchar2);
                _typeMapSQLCodeToSQL.Add(String.Format("NCHAR ({0})", textLength), (int)OracleDbType.NChar);
            }
            else
            {
                _typeMapSQLToSQLCode.Add((int)OracleDbType.NChar, "NCHAR");
                _typeMapSQLToSQLCode.Add((int)OracleDbType.Varchar2, "VARCHAR2");

                _typeMapSQLCodeToSQL.Add("VARCHAR2", (int)OracleDbType.Varchar2);
                _typeMapSQLCodeToSQL.Add("NCHAR", (int)OracleDbType.NChar);
            }
        }
        
        #endregion
    }
}
