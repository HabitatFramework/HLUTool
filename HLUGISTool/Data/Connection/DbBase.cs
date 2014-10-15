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
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;

namespace HLU.Data.Connection
{
    #region Enums

    public enum ConnectionTypes
    {
        ODBC = 2, 
        OleDb = 4, 
        Oracle = 8, 
        PostgreSQL = 16, 
        SQLServer = 32, 
        Unknown = 0
    }

    public enum Backends : int
    {
        Undetermined = 0,
        Access = 2,
        SqlServer = 4,
        Oracle = 8,
        PostgreSql = 16,
        DB2 = 32,
        UndeterminedOdbc = 64,
        UndeterminedOleDb = 128
    }

    #endregion

    abstract class DbBase : HLU.SqlBuilder
    {
        #region Fields
        
        private string _connectionString;
        private string _defaultSchema;
        private string _errorMessage;
        private string _pwd;
        protected Dictionary<string, string> _replaceDataTypes;
        protected Regex _sqlTypeRegex;
        
        private HLU.UI.View.Connection.ViewPassword _pwdWindow;
        private HLU.UI.ViewModel.ViewModelPassword _pwdViewModel;
        
        #endregion

        #region Constructor

        protected DbBase(ref string connString, ref string defaultSchema, ref bool promptPwd, string pwdMask, 
            bool useCommandBuilder, bool useColumnNames, bool isUnicode, bool useTimeZone, uint textLength, 
            uint binaryLength, uint timePrecision, uint numericPrecision, uint numericScale)
        {
            try
            {
                if (!String.IsNullOrEmpty(connString))
                {
                    ConnectionString = connString;
                    _defaultSchema = defaultSchema; // set by BrowseConnection
                }
                else
                {
                    BrowseConnection();
                }
                _useCommandBuilder = useCommandBuilder;
                _useColumnNames = useColumnNames;
                _isUnicode = isUnicode;
                _useTimeZone = useTimeZone;
                _textLength = textLength;
                _binaryLength = binaryLength;
                _timePrecision = timePrecision;
                _numericPrecision = numericPrecision;
                _numericScale = numericScale;
                _sqlTypeRegex = new Regex(@"\s*\(\s*[0-9]+(\s*,\s*[0-9]+\s*)*\)");
            }
            catch { throw; }
        }
        
        #endregion

        #region Public Static

        public static Backends GetBackend(string connString, ConnectionTypes connType)
        {
            switch (connType)
            {
                case ConnectionTypes.ODBC:
                    return DbOdbc.GetBackend(connString);
                case ConnectionTypes.OleDb:
                    return DbOleDb.GetBackend(connString);
                default:
                    return Backends.Undetermined;
            }
        }

        public static string GetDefaultSchema(Backends backend, 
            DbConnectionStringBuilder connStrBuilder, List<string> schemata)
        {
            switch (backend)
            {
                case Backends.Access:
                    return null;
                case Backends.PostgreSql:
                    return "public";
                case Backends.SqlServer:
                    return "dbo";
                case Backends.Oracle:
                    if ((connStrBuilder != null) && (connStrBuilder.ContainsKey("USER ID")))
                    {
                        string userIDstring = DbOracle.GetUserId(connStrBuilder);
                        if ((schemata != null) && (schemata.IndexOf(userIDstring) != -1))
                            return userIDstring;
                    }
                    return null;
                default:
                    if (connStrBuilder != null)
                    {
                        object userID;
                        if ((connStrBuilder.TryGetValue("UID", out userID)) ||
                            (connStrBuilder.TryGetValue("User ID", out userID)))
                        {
                            string userIDstring = userID.ToString();
                            if ((schemata != null) && (schemata.IndexOf(userIDstring) != -1))
                                return userIDstring;
                        }
                    }
                    return null;
            }
        }

        public static bool HasPassword(DbConnectionStringBuilder connStringBuilder)
        {
            if ((connStringBuilder == null) || IsIntegratedSecurity(connStringBuilder) || 
                !HasPasswordKey(connStringBuilder)) return false;

            object pwd = null;
            connStringBuilder.TryGetValue("Password", out pwd);
            return !String.IsNullOrEmpty(pwd.ToString());
        }

        public static string MaskPassword(DbConnectionStringBuilder connStringBuilder, string maskString)
        {
            if (connStringBuilder == null) return String.Empty;

            if (IsIntegratedSecurity(connStringBuilder) || !HasPasswordKey(connStringBuilder))
                return connStringBuilder.ConnectionString;

            DbConnectionStringBuilder tmpConnStrBuilder =
                new DbConnectionStringBuilder(connStringBuilder is OdbcConnectionStringBuilder);
            tmpConnStrBuilder.ConnectionString = connStringBuilder.ConnectionString;
            tmpConnStrBuilder.Remove("Password");
            tmpConnStrBuilder.Add("Password", maskString);

            return tmpConnStrBuilder.ConnectionString;
        }

        private static bool IsIntegratedSecurity(DbConnectionStringBuilder connStringBuilder)
        {
            if (connStringBuilder == null) return false;

            object integratedSecurity;
           
            if (connStringBuilder.TryGetValue("Integrated Security", out integratedSecurity))
            {
                if (integratedSecurity is String)
                {
                    string s = integratedSecurity.ToString().ToLower();
                    return s == "true" || s == "yes" || s == "SSPI";
                }
                else if (integratedSecurity is bool)
                {
                    return (bool)integratedSecurity;
                }
            }

            return false;
        }

        private static bool HasPasswordKey(DbConnectionStringBuilder connStringBuilder)
        {
            if (connStringBuilder == null)
                return false;
            else
                return connStringBuilder.ContainsKey("Password");
        }

        #endregion

        #region Public

        public string ConnectionString
        { 
            get { return _connectionString; }
            protected set { if (!String.IsNullOrEmpty(value)) { _connectionString = value; } }
        }

        public string DefaultSchema
        {
            get { return String.IsNullOrEmpty(_defaultSchema) ? null : _defaultSchema; }
            set { _defaultSchema = String.IsNullOrEmpty(value) ? null : value; }
        }

        public bool IsUnicode { get { return _isUnicode; } }

        public bool UseTimeZone { get { return _useTimeZone; } }

        public uint TextLength { get { return _textLength; } }

        public uint BinaryLength { get { return _binaryLength; } }

        public uint TimePrecision { get { return _timePrecision; } }

        public uint NumericPrecision { get { return _numericPrecision; } }

        public uint NumericScale { get { return _numericScale; } }

        public string RestrictionNameCatalog { get { return _restrictionNameCatalog; } }

        public string RestrictionNameSchema { get { return _restrictionNameSchema; } }

        public string RestrictionNameTable { get { return _restrictionNameTable; } }

        public string RestrictionNameColumn { get { return _restrictionNameColumn; } }

        public string BackendDataType(Type systemType)
        {
            return SqlToSqlCodeType(SystemToSqlType(systemType));
        }

        public Type SystemDataType(String backendType)
        {
            try
            {
                backendType = _sqlTypeRegex.Replace(backendType, "").ToLowerInvariant();

                int tsql = -1;
                if (!_sqlSynonyms.TryGetValue(backendType, out tsql))
                    tsql = _typeMapSQLCodeToSQL.AsEnumerable()
                        .SingleOrDefault(t => _sqlTypeRegex.Replace(t.Key, "").ToLowerInvariant() == backendType).Value;

                Type tsys;
                if (_typeMapSQLToSystem.TryGetValue(tsql, out tsys)) return tsys;
            }
            catch { }

            return (Type)Type.Missing;
        }

        public string ErrorMessage { get { return _errorMessage; } }

        public string QualifyTableName(string tableName)
        {
            if (String.IsNullOrEmpty(tableName))
                return String.Empty;
            else if (String.IsNullOrEmpty(_defaultSchema))
                return QuoteIdentifier(tableName);
            else
                return QuoteIdentifier(_defaultSchema) + "." + QuoteIdentifier(tableName);
        }

        public bool FillSchema<T>(SchemaType schemaType, ref T table) where T : DataTable, new()
        {
            if (table == null) table = new T();
            return FillSchema<T>(schemaType, "SELECT * FROM " + table.TableName, ref table);
        }

        public int FillTable<T>(ref T table) where T : DataTable, new()
        {
            if (table == null) table = new T();
            return FillTable<T>("SELECT * FROM " + QuoteIdentifier(table.TableName), ref table);
        }

        public DataTable GetSchema<C, T>(string collectionName, string restrictionName, 
            string restrictionValue, C connection, T transaction)
            where C : DbConnection
            where T : DbTransaction
        {
            string[] restrictionNames;
            if (_schemaRestrictions.TryGetValue(collectionName, out restrictionNames))
            {
                string[] restrictions = new string[restrictionNames.Length];
                int restrictionPosition = Array.IndexOf(restrictionNames, restrictionName);
                if (restrictionPosition != -1)
                {
                    restrictions[restrictionPosition] = restrictionValue;
                    return GetSchema<C, T>(collectionName, restrictions, connection, transaction);
                }
            }

            return null;
        }

        public DataTable GetSchema<C, T>(string collectionName, string[] restrictionValues, 
            C connection, T transaction)
            where C : DbConnection
            where T : DbTransaction
        {
            if (transaction != null) return null;

            try
            {
                ConnectionState previousConnectionState = connection.State;
                if ((connection.State & ConnectionState.Open) != ConnectionState.Open) connection.Open();

                DataTable dt = null;

                if (String.IsNullOrEmpty(collectionName))
                {
                    dt = connection.GetSchema();
                }
                else if (restrictionValues == null)
                {
                    dt = connection.GetSchema(collectionName);
                }
                else
                {
                    string[] restrictionNames;
                    if (!_schemaRestrictions.TryGetValue(collectionName, out restrictionNames) ||
                        (restrictionValues.Length == restrictionNames.Length))
                    {
                        dt = connection.GetSchema(collectionName, restrictionValues);
                    }
                    else
                    {
                        string[] restrictions = new string[restrictionNames.Length];
                        if (restrictionValues.Length < restrictionNames.Length)
                            Array.Copy(restrictionValues, 0, restrictions, 0, restrictionValues.Length);
                        else if (restrictionValues.Length > restrictionNames.Length)
                            Array.Copy(restrictionValues, 0, restrictions, 0, restrictionNames.Length);
                        dt = connection.GetSchema(collectionName, restrictions);
                    }
                }

                if (previousConnectionState == ConnectionState.Closed) connection.Close();

                return dt;
            }
            catch { return null; }
        }

        /// <summary>
        /// Count the number database rows that match the list of
        /// WHERE conditions.
        /// </summary>
        /// <param name="targetColumns">The target database columns.</param>
        /// <param name="whereConds">The list of where conds.</param>
        /// <returns>An integer of the number of rows matching the SQL.</returns>
        public int SqlCount(DataColumn[] targetColumns, List<SqlFilterCondition> whereConds)
        {
            if ((targetColumns == null) || (targetColumns.Length == 0)) return 0;

            try
            {
                bool qualifyColumns = QualifyColumnNames(targetColumns);
                bool additionalTables;
                string fromList = FromList(true, targetColumns, true, ref whereConds, out additionalTables);
                qualifyColumns |= additionalTables;
                StringBuilder sbCommandText = new StringBuilder("SELECT COUNT(*) AS N");
                sbCommandText.Append(fromList);
                sbCommandText.Append(WhereClause(true, true, qualifyColumns, whereConds));

                object result = ExecuteScalar(sbCommandText.ToString(), 0, CommandType.Text);

                int numRows = 0;
                if (result != null) numRows = Convert.ToInt32(result);

                return numRows;
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return 0;
            }
        }

        /// <summary>
        /// Count the number database rows that match the list of
        /// WHERE conditions.
        /// </summary>
        /// <param name="targetTables">The target database tables.</param>
        /// <param name="whereConds">The list of where conds.</param>
        /// <returns>An integer of the number of rows matching the SQL.</returns>
        public int SqlCount(DataTable[] targetTables, List<SqlFilterCondition> whereConds)
        {
            if ((targetTables == null) || (targetTables.Length == 0) ||
                (targetTables[0].Columns.Count == 0)) return 0;

            try
            {
                bool qualifyColumns = targetTables.Length > 1;
                bool additionalTables;
                string fromList = FromList(true, true, targetTables, ref whereConds, out additionalTables);
                qualifyColumns |= additionalTables;
                StringBuilder sbCommandText = new StringBuilder("SELECT COUNT(*) AS N");
                sbCommandText.Append(fromList);
                sbCommandText.Append(WhereClause(true, true, qualifyColumns, whereConds));

                object result = ExecuteScalar(sbCommandText.ToString(), 0, CommandType.Text);

                int numRows = 0;
                if (result != null) numRows = Convert.ToInt32(result);

                return numRows;
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return 0;
            }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR5 (Select by attributes interface)
        // Count the number of database rows using a WHERE statement
        // based on both a list of conditions and a free-text string.
        // 
        /// <summary>
        /// Count the number database rows that match the list of
        /// WHERE conditions and string of WHERE clauses.
        /// </summary>
        /// <param name="targetTables">The target database tables.</param>
        /// <param name="whereConds">The list of where conds.</param>
        /// <param name="sqlWhereClause">The string of where clauses.</param>
        /// <returns>An integer of the number of rows matching the SQL.</returns>
        public int SqlCount(DataTable[] targetTables, List<SqlFilterCondition> whereConds, string sqlWhereClause)
        {
            if ((targetTables == null) || (targetTables.Length == 0)) return 0;

            try
            {
                // Determine if the column names need qualifiying.
                bool qualifyColumns = targetTables.Length > 1;

                // Create a string of the tables to query based on the the
                // target columns to select and the list of from tables.
                bool additionalTables;
                List<SqlFilterCondition> fromConds = new List<SqlFilterCondition>();
                DataColumn[] targetColumns = new DataColumn[0];
                string fromList = FromList(true, true, targetTables, ref whereConds, out additionalTables);

                // Force the column names to be qualified if there are any
                // additional tables.
                qualifyColumns |= additionalTables;

                // Build a sql command.
                StringBuilder sbCommandText = new StringBuilder("SELECT COUNT(*) AS N");

                // Append the tables to select from.
                sbCommandText.Append(fromList);

                // Append the where clauses relating to the from table joins.
                string fromClause = WhereClause(true, true, qualifyColumns, whereConds);
                sbCommandText.Append(fromClause);

                // Append any additional where clauses passed.
                if (string.IsNullOrEmpty(fromClause))
                    sbCommandText.Append(" WHERE (").Append(sqlWhereClause).Append(")");
                else
                    sbCommandText.Append(" AND (").Append(sqlWhereClause).Append(")");

                // Execute the sql command to count the number of records.
                object result = ExecuteScalar(sbCommandText.ToString(), 0, CommandType.Text);

                int numRows = 0;
                if (result != null) numRows = Convert.ToInt32(result);

                return numRows;
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return 0;
            }
        }
        //---------------------------------------------------------------------

        #endregion

        #region Protected

        protected bool _isUnicode;
        
        protected bool _useTimeZone;
        
        protected uint _textLength;
        
        protected uint _binaryLength;
        
        protected uint _timePrecision;
        
        protected uint _numericPrecision;
        
        protected uint _numericScale;
        
        protected bool _useCommandBuilder;
        
        protected bool _useColumnNames;

        protected int _startParamNo = 1;

        protected string _restrictionNameCatalog = "Catalog";
        
        protected string _restrictionNameSchema = "Schema";
        
        protected string _restrictionNameTable = "Table";
        
        protected string _restrictionNameColumn = "Column";

        protected string _parameterPrefixCurr = "";

        protected string _parameterPrefixOrig = "Original_";

        protected string _parameterPrefixNull = "IsNull_";
        
        protected Dictionary<int, string> _typeMapSQLToSQLCode;

        protected Dictionary<string, int> _typeMapSQLCodeToSQL;

        protected Dictionary<string, int> _sqlSynonyms;

        protected Dictionary<string, string[]> _schemaRestrictions;

        protected Dictionary<string, string> ReplaceDataTypes
        {
            get
            {
                if (_replaceDataTypes == null)
                {
                    _replaceDataTypes = new Dictionary<string, string>();
                    _replaceDataTypes.Add("System.long", "System.Int64");
                    _replaceDataTypes.Add("sql_variant", "Variant");
                    _replaceDataTypes.Add("Short", "SmallInt");
                    _replaceDataTypes.Add("Long", "BigInt");
                    _replaceDataTypes.Add("Bit", "Boolean");
                    _replaceDataTypes.Add("LongBinary", "LongVarBinary");
                    _replaceDataTypes.Add("LongText", "LongWVarChar");
                }
                return _replaceDataTypes;
            }
        }

        protected void GetMetaData<C, T>(Type enumType, C connection, T transaction)
            where C : DbConnection
            where T : DbTransaction
        {
            _typeMapSQLToSystem = new Dictionary<int, Type>();
            _typeMapSystemToSQL = new Dictionary<Type, int>();
            _sqlSynonyms = new Dictionary<string, int>();

            DataTable metaDataCollections = GetSchema(DbMetaDataCollectionNames.MetaDataCollections,
                null, connection, transaction);

            if (metaDataCollections == null) return;

            if ((metaDataCollections.AsEnumerable().Count(r => r.Field<string>(DbMetaDataColumnNames.CollectionName) ==
                DbMetaDataCollectionNames.Restrictions) == 1))
            {
                DataTable restrictions = GetSchema(DbMetaDataCollectionNames.Restrictions, null, connection, transaction);

                if (restrictions != null)
                {
                    string numRestCol = DbMetaDataColumnNames.NumberOfRestrictions;
                    if (!restrictions.Columns.Contains(numRestCol)) numRestCol = "RestrictionNumber";

                    _schemaRestrictions = (from r in restrictions.AsEnumerable()
                                           let collName = r.Field<string>(DbMetaDataColumnNames.CollectionName)
                                           group r by r.Field<string>(DbMetaDataColumnNames.CollectionName) into collGroup
                                           select new
                                           {
                                               key = collGroup.Key,
                                               value = collGroup.Select(n => n.Field<string>("RestrictionName")).ToArray()
                                           }
                                           ).ToDictionary(kv => kv.key, kv => kv.value);

                    string[] restrictionNames;
                    if (_schemaRestrictions.TryGetValue("Columns", out restrictionNames))
                    {
                        switch (restrictionNames.Length)
                        {
                            case 4:
                                _restrictionNameCatalog = restrictionNames[0];
                                _restrictionNameSchema = restrictionNames[1];
                                _restrictionNameTable = restrictionNames[2];
                                _restrictionNameColumn = restrictionNames[3];
                                break;
                            case 3:
                                _restrictionNameSchema = restrictionNames[0];
                                _restrictionNameTable = restrictionNames[1];
                                _restrictionNameColumn = restrictionNames[2];
                                break;
                        }
                    }
                }
            }

            if ((metaDataCollections.AsEnumerable().Count(r => r.Field<string>(DbMetaDataColumnNames.CollectionName) ==
                DbMetaDataCollectionNames.DataTypes) == 1))
            {
                DataTable dataTypes = GetSchema(DbMetaDataCollectionNames.DataTypes, null, connection, transaction);

                if (dataTypes != null)
                {
                    if (dataTypes.Columns.Contains(DbMetaDataColumnNames.ProviderDbType))
                    {
                        _typeMapSQLToSystem = (from rd in
                                                   (from r in dataTypes.AsEnumerable()
                                                    where r[DbMetaDataColumnNames.ProviderDbType] != DBNull.Value &&
                                                         r[DbMetaDataColumnNames.DataType] != DBNull.Value
                                                    group r by r.Field<int>(DbMetaDataColumnNames.ProviderDbType) into g
                                                    select g.First())
                                               let dbTypeCode = rd.Field<int>(DbMetaDataColumnNames.ProviderDbType)
                                               let dataTypeStr = rd.Field<string>(DbMetaDataColumnNames.DataType)
                                               where Enum.IsDefined(enumType, dbTypeCode)
                                               let dataType = Type.GetType(CleanDataType(dataTypeStr))
                                               select new KeyValuePair<int, Type>(dbTypeCode, dataType))
                                               .ToDictionary(kv => kv.Key, kv => kv.Value);
                    }
                    _typeMapSystemToSQL = new Dictionary<Type, int>();
                    int sysType;
                    foreach (KeyValuePair<int, Type> kv in _typeMapSQLToSystem)
                    {
                        if (!_typeMapSystemToSQL.TryGetValue(kv.Value, out sysType))
                            _typeMapSystemToSQL.Add(kv.Value, kv.Key);
                    }

                    _sqlSynonyms = (from r in dataTypes.AsEnumerable()
                                    let dataTypeStr = r.Field<string>(DbMetaDataColumnNames.DataType)
                                    where dataTypeStr != null
                                    let dataType = CleanDataType(dataTypeStr)
                                    select new
                                    {
                                        key = r.Field<string>(DbMetaDataColumnNames.TypeName).ToLower(),
                                        value = SystemToDbType(Type.GetType(dataType))
                                    }).ToDictionary(kv => kv.key, kv => kv.value);
                }
            }
        }

        protected int EnumValue(Type enumType, string typeName, bool ignoreCase)
        {
            try
            {
                return (int)Enum.Parse(enumType, typeName, ignoreCase);
            }
            catch { return -1; }
        }

        protected string CleanDataType(string dataType)
        {
            string test;
            if (ReplaceDataTypes.TryGetValue(dataType, out test))
                return test;
            else
                return dataType;
        }

        protected String DbTypeToString(int dbTypeCode)
        {
            string typeName;
            if (_typeMapSQLToSQLCode.TryGetValue(dbTypeCode, out typeName))
                return typeName;
            else
                return null;
        }

        protected void ReplaceType(Type sysType, int dbTypeNew, Dictionary<Type, int> typeDictionary)
        {
            int dbTypeOld;
            if (typeDictionary.TryGetValue(sysType, out dbTypeOld) && (dbTypeOld != dbTypeNew))
            {
                typeDictionary.Remove(sysType);
                typeDictionary.Add(sysType, dbTypeNew);
            }
        }

        protected int SQLCodeToSQLType(string sqlType)
        {
            sqlType = sqlType.ToLower();

            int typeCode;
            if (_sqlSynonyms.TryGetValue(sqlType, out typeCode))
                return typeCode;
            else
                return -1;
        }

        #region Login

        protected void Login<B, C>(string userNameLabel, string connectionString, 
            ref bool promptPwd, ref B connectionStringBuilder, ref C connection)
            where B : DbConnectionStringBuilder, new()
            where C : DbConnection, new()
        {
            connectionStringBuilder = new B();
            connectionStringBuilder.ConnectionString = connectionString;

            if (!promptPwd)
            {
                promptPwd = HasPassword(connectionStringBuilder);
                connection = new C();
                connection.ConnectionString = connectionStringBuilder.ConnectionString;
            }
            else
            {
                for (int i = 0; i < 3; i++)
                {
                    PromptPassword(userNameLabel, ref connectionStringBuilder);
                    connection = new C();
                    connection.ConnectionString = connectionStringBuilder.ConnectionString;
                    try
                    {
                        connection.Open();
                        connection.Close();
                        break;
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message != "cancelled")
                            break;
                        else if (i < 2)
                            MessageBox.Show(ex.Message, "Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        else
                            throw;
                    }
                }
            }
        }

        protected void PromptPassword<T>(string userLabel, ref T connStrBuilder) 
            where T : DbConnectionStringBuilder
        {
            if (connStrBuilder == null) return;

            string connType = Enum.GetName(typeof(Backends), this.Backend).Replace("Undetermined", "");
            
            try
            {
                if (connStrBuilder.ContainsKey("Password"))
                    connStrBuilder.Remove("Password");

                _pwdWindow = new HLU.UI.View.Connection.ViewPassword();
                if ((_pwdWindow.Owner = App.GetActiveWindow()) == null)
                    throw (new Exception("No parent window loaded"));

                // create ViewModel to which main window binds
                _pwdViewModel = new HLU.UI.ViewModel.ViewModelPassword();
                object dataSource;
                if (connStrBuilder.TryGetValue("Data Source", out dataSource) ||
                    connStrBuilder.TryGetValue("DataSource", out dataSource) ||
                    connStrBuilder.TryGetValue("Host", out dataSource))
                    _pwdViewModel.DisplayName = dataSource.ToString();
                else
                    _pwdViewModel.DisplayName = connType + " Connection";
                object userName;
                if (connStrBuilder.TryGetValue("UID", out userName) || connStrBuilder.TryGetValue("USER ID", out userName))
                {
                    _pwdViewModel.UserText = userName.ToString();
                    _pwdViewModel.UserLabel = userLabel;
                }
                else
                {
                    _pwdWindow.Height -= _pwdWindow.GridUser.Height;
                    _pwdWindow.GridUser.Visibility = Visibility.Collapsed;
                }

                // when ViewModel asks to be closed, close window
                _pwdViewModel.RequestClose +=
                    new HLU.UI.ViewModel.ViewModelPassword.RequestCloseEventHandler(_pwdViewModel_RequestClose);

                // allow all controls in window to bind to ViewModel by setting DataContext
                _pwdWindow.DataContext = _pwdViewModel;

                _pwdWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                _pwdWindow.Topmost = true;

                // show window
                _pwdWindow.ShowDialog();

                // throw error if connection failed
                if (!String.IsNullOrEmpty(_errorMessage))
                    throw (new Exception(_errorMessage));
                else if (!String.IsNullOrEmpty(_pwd))
                    connStrBuilder.Add("Password", _pwd);
                else
                    throw new Exception("cancelled");
            }
            catch (Exception ex)
            {
                if (ex.Message != "cancelled")
                    MessageBox.Show("Server responded with an error:\n\n" + ex.Message,
                        connType + " Error", MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
            finally { _pwd = null; }
        }

        protected void _pwdViewModel_RequestClose(string password, string errorMsg)
        {
            _pwdViewModel.RequestClose -= _pwdViewModel_RequestClose;
            _pwdWindow.Close();

            if (!String.IsNullOrEmpty(errorMsg))
            {
                _errorMessage = errorMsg;
            }
            else if (!String.IsNullOrEmpty(password))
            {
                _pwd = password;
            }
        }

        #endregion
       
        #endregion

        #region Private Methods

        protected int SystemToSqlType(Type tsys)
        {
            int tsql;
            if (_typeMapSystemToSQL.TryGetValue(tsys, out tsql))
                return tsql;
            else
                return (-1);
        }

        protected Type SqlToSystemType(int tsql)
        {
            Type tsys;
            if (_typeMapSQLToSystem.TryGetValue(tsql, out tsys))
                return tsys;
            else
                return (Type)Type.Missing;
        }

        protected string SqlToSqlCodeType(int tsql)
        {
            string tcode;
            _typeMapSQLToSQLCode.TryGetValue(tsql, out tcode);
            return tcode;
        }

        #endregion

        #region Public Override
        
        public override string QuoteIdentifier(string identifier)
        {
            if (!String.IsNullOrEmpty(identifier))
            {
                if (!identifier.StartsWith(QuotePrefix)) identifier = identifier.Insert(0, QuotePrefix);
                if (!identifier.EndsWith(QuoteSuffix)) identifier += QuoteSuffix;
            }
            return identifier;
        }

        public override string TargetList(DataColumn[] targetColumns, bool quoteIdentifiers, 
            bool checkQualify, ref bool qualifyColumns, out DataTable resultTable)
        {
            resultTable = new DataTable();

            if ((targetColumns == null) || (targetColumns.Length == 0)) return String.Empty; ;

            StringBuilder sbTargetList = new StringBuilder();

            try
            {
                if (checkQualify) qualifyColumns = QualifyColumnNames(targetColumns);

                string columnAlias;
                foreach (DataColumn c in targetColumns)
                {
                    if (qualifyColumns)
                    {
                        columnAlias = ColumnAlias(c);
                        if (quoteIdentifiers)
                            sbTargetList.Append(String.Format(",{0}.{1} AS {2}", QuoteIdentifier(c.Table.TableName),
                                QuoteIdentifier(c.ColumnName), QuoteIdentifier(columnAlias)));
                        else
                            sbTargetList.Append(String.Format(",{0}.{1} AS {2}", c.Table.TableName, c.ColumnName, columnAlias));
                        resultTable.Columns.Add(new DataColumn(columnAlias, c.DataType));
                    }
                    else
                    {
                        if (quoteIdentifiers)
                            sbTargetList.Append(String.Format(",{0}", QuoteIdentifier(c.ColumnName)));
                        else
                            sbTargetList.Append(String.Format(",{0}", c.ColumnName));
                        resultTable.Columns.Add(new DataColumn(c.ColumnName, c.DataType));
                    }
                }
                sbTargetList.Remove(0, 1);
            }
            catch { }

            return sbTargetList.ToString();
        }

        /// <summary>
        /// Select database records using a SQL statement based on an array
        /// of target columns to select, a list of tables to select from, and
        /// a list of where conditions.
        /// </summary>
        /// <param name="selectDistinct">if set to <c>true</c> select only DISTINCT values.</param>
        /// <param name="targetColumns">The target columns to select.</param>
        /// <param name="whereConds">The list of where conds to apply.</param>
        /// <returns></returns>
        public override DataTable SqlSelect(bool selectDistinct, DataColumn[] targetColumns, List<SqlFilterCondition> whereConds)
        {
            if ((targetColumns == null) || (targetColumns.Length == 0)) return new DataTable();

            try
            {
                DataTable resultTable = null;
                bool qualifyColumns = QualifyColumnNames(targetColumns);
                bool additionalTables;
                string fromList = FromList(true, targetColumns, true, ref whereConds, out additionalTables);
                qualifyColumns |= additionalTables;
                StringBuilder sbCommandText = new StringBuilder(selectDistinct ? "SELECT DISTINCT " : "SELECT ");
                sbCommandText.Append(TargetList(targetColumns, true, false, ref qualifyColumns, out resultTable));
                sbCommandText.Append(fromList);
                sbCommandText.Append(WhereClause(true, true, qualifyColumns, whereConds));

                FillTable<DataTable>(sbCommandText.ToString(), ref resultTable);

                return resultTable;
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return new DataTable();
            }
        }

        /// <summary>
        /// Select database records using a SQL statement based on an array
        /// of target tables to select from, a list of tables to select from, and
        /// a list of where conditions.
        /// </summary>
        /// <param name="selectDistinct">if set to <c>true</c> select only DISTINCT values.</param>
        /// <param name="targetTables">The target tables to select from.</param>
        /// <param name="whereConds">The list of where conds to apply.</param>
        /// <returns></returns>
        public override DataTable SqlSelect(bool selectDistinct, DataTable[] targetTables, List<SqlFilterCondition> whereConds)
        {
            if ((targetTables == null) || (targetTables.Length == 0) ||
                (targetTables[0].Columns.Count == 0)) return new DataTable();

            try
            {
                DataTable resultTable = null;
                bool qualifyColumns = targetTables.Length > 1;
                bool additionalTables;
                string fromList = FromList(true, true, targetTables, ref whereConds, out additionalTables);
                qualifyColumns |= additionalTables;
                StringBuilder sbCommandText = new StringBuilder(selectDistinct ? "SELECT DISTINCT " : "SELECT ");
                sbCommandText.Append(TargetList(targetTables, true, ref qualifyColumns, out resultTable));
                sbCommandText.Append(fromList);
                sbCommandText.Append(WhereClause(true, true, qualifyColumns, whereConds));

                FillTable<DataTable>(sbCommandText.ToString(), ref resultTable);

                return resultTable;
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return new DataTable();
            }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR5 (Select by attributes interface)
        // Count the number of database rows using a WHERE statement
        // based on both a list of conditions and a free-text string.
        // 
        /// <summary>
        /// Select database records using a SQL statement based on an array
        /// of target columns to select, a list of tables to select from, and
        /// a string of where clauses.
        /// </summary>
        /// <param name="selectDistinct">if set to <c>true</c> select only DISTINCT values.</param>
        /// <param name="targetColumns">The target columns to select.</param>
        /// <param name="sqlFromTables">The tables to select from.</param>
        /// <param name="sqlWhereClause">The where clauses to apply.</param>
        /// <returns></returns>
        public DataTable SqlSelect(bool selectDistinct, DataColumn[] targetColumns, List<DataTable> sqlFromTables, string sqlWhereClause)
        {
            if ((targetColumns == null) || (targetColumns.Length == 0)) return new DataTable();

            try
            {
                // Declare a new empty result data table.
                DataTable resultTable = null;

                // Determine if the column names need qualifiying.
                bool qualifyColumns = QualifyColumnNames(targetColumns);

                // Create a string of the tables to query based on the the
                // target columns to select and the list of from tables.
                bool additionalTables;
                List<SqlFilterCondition> fromConds = new List<SqlFilterCondition>();
                string fromList = FromList(true, true, targetColumns, sqlFromTables, ref fromConds, out additionalTables);

                // Force the column names to be qualified if there are any
                // additional tables.
                qualifyColumns |= additionalTables;

                // Build a sql command.
                StringBuilder sbCommandText = new StringBuilder(selectDistinct ? "SELECT DISTINCT " : "SELECT ");

                // Append the columns to be selected.
                sbCommandText.Append(TargetList(targetColumns, true, false, ref qualifyColumns, out resultTable));

                // Append the tables to select from.
                sbCommandText.Append(fromList);

                // Append the where clauses relating to the from table joins.
                string fromClause = WhereClause(true, true, qualifyColumns, fromConds);
                sbCommandText.Append(fromClause);

                // Append any additional where clauses passed.
                if (string.IsNullOrEmpty(fromClause))
                    sbCommandText.Append(" WHERE (").Append(sqlWhereClause).Append(")");
                else
                    sbCommandText.Append(" AND (").Append(sqlWhereClause).Append(")");

                // Fill the result table using the sql command.
                FillTable<DataTable>(sbCommandText.ToString(), ref resultTable);

                return resultTable;
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return new DataTable();
            }
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // CHANGED: CR5 (Select by attributes interface)
        // Execute the SQL statement to check if it is valid and
        // see if it returns at least one record.
        // 
        /// <summary>
        /// SQLs the validate.
        /// </summary>
        /// <param name="targetColumns">The target columns.</param>
        /// <param name="sqlFromTables">The SQL from tables.</param>
        /// <param name="sqlWhereClause">The SQL where clause.</param>
        /// <returns></returns>
        public string SqlValidate(DataColumn[] targetColumns, List<DataTable> sqlFromTables, string sqlWhereClause)
        {
            if ((targetColumns == null) || (targetColumns.Length == 0)) return "Error verifying Sql";

            try
            {
                // Declare a new empty result data table.
                DataTable resultTable = null;

                // Determine if the column names need qualifiying.
                bool qualifyColumns = QualifyColumnNames(targetColumns);

                // Create a string of the tables to query based on the the
                // target columns to select and the list of from tables.
                bool additionalTables;
                List<SqlFilterCondition> fromConds = new List<SqlFilterCondition>();
                string fromList = FromList(true, true, targetColumns, sqlFromTables, ref fromConds, out additionalTables);

                // Force the column names to be qualified if there are any
                // additional tables.
                qualifyColumns |= additionalTables;

                // Build two sql commands.
                StringBuilder sbCommandText = new StringBuilder("SELECT TOP 1 ");

                // Append the columns to be selected.
                string targetList = TargetList(targetColumns, true, false, ref qualifyColumns, out resultTable);
                sbCommandText.Append(targetList);

                // Append the tables to select from.
                sbCommandText.Append(fromList);

                // Append the where clauses relating to the from table joins.
                string fromClause = WhereClause(true, true, qualifyColumns, fromConds);
                sbCommandText.Append(fromClause);

                // Append any additional where clauses passed.
                if (string.IsNullOrEmpty(fromClause))
                    sbCommandText.Append(" WHERE ").Append(sqlWhereClause);
                else
                    sbCommandText.Append(" AND (").Append(sqlWhereClause).Append(")");

                // Execute the sql command to check it is valid.
                bool valid = false;
                try
                {
                    valid = ValidateQuery(sbCommandText.ToString(), 0, CommandType.Text);
                }
                catch (Exception ex)
                {
                    throw ex;
                }

                // If the sql is not valid then return error.
                if (!valid) return "Sql is invalid";

                // Fill the result table using the sql command.
                FillTable<DataTable>(sbCommandText.ToString(), ref resultTable);
                int numRows = 0;
                if (resultTable != null) numRows = resultTable.Rows.Count;

                return numRows.ToString();
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return _errorMessage;
            }
        }
        //---------------------------------------------------------------------

        #endregion

        #region Public Abstract
        
        public abstract Backends Backend { get; }

        public abstract DbConnectionStringBuilder ConnectionStringBuilder { get; }

        public abstract IDbConnection Connection { get; }

        public abstract bool FillSchema<T>(SchemaType schemaType, string sql, ref T table) where T : DataTable, new();

        public abstract int FillTable<T>(string sql, ref T table) where T : DataTable, new();
        
        public abstract IDbTransaction Transaction { get; }

        public abstract IDbCommand CreateCommand();

        public abstract IDbDataAdapter CreateAdapter();

        public abstract IDbDataAdapter CreateAdapter<T>(T table) where T : DataTable, new();

        public abstract int Update<T>(T table) where T : DataTable, new();

        public abstract int Update<T, R>(R[] rows) where T : DataTable, new() where R : DataRow;

        public abstract int Update<T>(T dataSet, string sourceTable) where T : DataSet;

        public abstract bool BeginTransaction(bool commitPrevious, IsolationLevel isolationLevel);

        public abstract bool CommitTransaction();

        public abstract bool RollbackTransaction();

        public abstract IDataReader ExecuteReader(string sql, int commandTimeout, CommandType commandType);

        public abstract int ExecuteNonQuery(string sql, int commandTimeout, CommandType commandType);

        public abstract object ExecuteScalar(string sql, int commandTimeout, CommandType commandType);

        public abstract bool ValidateQuery(string sql, int commandTimeout, CommandType commandType);

        public abstract int Update<T>(T table, string insertCommand, string updateCommand, string deleteCommand) where T : DataTable;

        public abstract bool ContainsDataSet(DataSet ds, out string errorMessage);

        public bool CreateTable(DataTable adoTable)
        {
            try
            {
                StringBuilder sql = new StringBuilder();
                foreach (DataColumn c in adoTable.Columns)
                {
                    int dbColTypeInt;
                    string dbColTypeString;
                    if (_typeMapSystemToSQL.TryGetValue(c.DataType, out dbColTypeInt) &&
                        _typeMapSQLToSQLCode.TryGetValue(dbColTypeInt, out dbColTypeString))
                    {
                        if ((c.DataType == typeof(string)) && (c.MaxLength != -1))
                            dbColTypeString = dbColTypeString.Replace("(" + TextLength + ")", "(" + c.MaxLength + ")");
                        sql.Append(String.Format(", {0} {1} {2}", QuoteIdentifier(c.ColumnName),
                            dbColTypeString, c.AllowDBNull ? "NULL" : "NOT NULL"));
                    }
                }

                StringBuilder primaryKey = new StringBuilder();
                foreach (DataColumn c in adoTable.PrimaryKey)
                {
                    primaryKey.Append(String.Format(", {0}", QuoteIdentifier(c.ColumnName)));
                }
                if (primaryKey.Length > 0)
                    primaryKey.Remove(0, 2).Insert(0, String.Format(", CONSTRAINT {0} PRIMARY KEY (",
                        QuoteIdentifier("pk__" + adoTable.TableName))).Append(")");

                if (sql.Length > 0)
                    sql.Remove(0, 2).Insert(0, String.Format("CREATE TABLE {0} (",
                        QualifyTableName(adoTable.TableName))).Append(primaryKey.Length > 0 ?
                        primaryKey.ToString() : String.Empty).Append(")");

                int returnVal = ExecuteNonQuery(sql.ToString(), Connection.ConnectionTimeout, CommandType.Text);

                return returnVal != -1;
            }
            catch { return false; }
        }

        #endregion

        #region Protected Abstract
        
        protected abstract void BrowseConnection();

        protected abstract string ParameterPrefix { get; }

        protected abstract string ParameterName(string prefix, string columnName, int paramNo);

        protected abstract string ParameterMarker(string parameterName);

        #endregion
    }
}
