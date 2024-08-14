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
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using HLU.Data.Connection;
using Npgsql;

namespace HLU.UI.ViewModel
{
    class ViewModelConnectPgSql : ViewModelBase, IDataErrorInfo
    {
        #region private Members

        private IntPtr _windowHandle;
        private string _displayName;
        private RelayCommand _okCommand;
        private RelayCommand _cancelCommand;
        private string[] _sslModes;
        private string[] _encodings;
        private string _encoding = "<default>";
        private string[] _databases = new string[] { };
        private List<String> _schemata = new List<String>();

        private NpgsqlConnectionStringBuilder _connStrBuilder;

        #endregion

        #region Constructor

        public ViewModelConnectPgSql()
        {
            _connStrBuilder = new NpgsqlConnectionStringBuilder();
            _connStrBuilder.Host = "localhost";
            _connStrBuilder.Port = 5432;
            _connStrBuilder.SslMode = Npgsql.SslMode.Prefer;
        }

        #endregion

        #region Connection String Builder

        public NpgsqlConnectionStringBuilder ConnectionStringBuilder { get { return _connStrBuilder; } }

        #endregion

        #region Display Name

        public override string DisplayName
        {
            get { return _displayName; }
            set { _displayName = value; }
        }

        #endregion

        #region Window Title

        public override string WindowTitle { get { return DisplayName; } }

        #endregion

        #region RequestClose

        // declare the delegate since using non-generic pattern
        public delegate void RequestCloseEventHandler(string connString, string encoding, string defaultSchema, string errorMsg);

        // declare the event
        public event RequestCloseEventHandler RequestClose;

        #endregion

        #region Ok Command

        /// <summary>
        /// Create Ok button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand OkCommand
        {
            get
            {
                if (_okCommand == null)
                {
                    Action<object> okAction = new Action<object>(this.OkCommandClick);
                    _okCommand = new RelayCommand(okAction, param => this.CanOk);
                }

                return _okCommand;
            }
        }

        /// <summary>
        /// Handles event when Ok button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void OkCommandClick(object param)
        {
            NpgsqlConnection cn;

            try
            {
                cn = new NpgsqlConnection(_connStrBuilder.ConnectionString);

                cn.Open();
                cn.Close();

                this.RequestClose(_connStrBuilder.ConnectionString, _encoding != _encodings[0] ? 
                    _encoding : null, _connStrBuilder.SearchPath.Split(',')[0], null);
            }
            catch (NpgsqlException exNpgsql)
            {
                MessageBox.Show("PostgreSQL Server responded with an error:\n\n" + exNpgsql.Message,
                     "PostgreSQL Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally { cn = null; }
        }

        /// <summary>
        /// Determines whether the Ok button is enabled
        /// To be enabled the following must be true:
        /// server name and database must be set; if windows authentication is not set then a username and password are required.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanOk 
        {
            get
            {
                return !(String.IsNullOrEmpty(_connStrBuilder.Host) || (_connStrBuilder.Port == 0) || 
                    String.IsNullOrEmpty(_connStrBuilder.Database) || String.IsNullOrEmpty(_connStrBuilder.UserName) || 
                    String.IsNullOrEmpty(_connStrBuilder.SearchPath)); 
            }
        }

        #endregion

        #region Cancel Command

        /// <summary>
        /// Create Cancel button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand CancelCommand
        {
            get
            {
                if (_cancelCommand == null)
                {
                    Action<object> cancelAction = new Action<object>(this.CancelCommandClick);
                    _cancelCommand = new RelayCommand(cancelAction);
                }

                return _cancelCommand;
            }
        }

        /// <summary>
        /// Handles event when Cancel button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void CancelCommandClick(object param)
        {
            this.RequestClose(null, null, null, null);
        }

        #endregion

        #region Host

        public string Host
        {
            get { return _connStrBuilder.Host; }
            set
            {
                if (!String.IsNullOrEmpty(value) && (value != _connStrBuilder.Host))
                    _connStrBuilder.Host = value;
            }
        }

        public int Port
        {
            get { return _connStrBuilder.Port; }
            set { if (value != _connStrBuilder.Port) _connStrBuilder.Port = value; }
        }

        public string[] SslModes
        {
            get 
            {
                if (_sslModes == null)
                    _sslModes = new string[] { "Allow", "Disable", "Prefer", "Require" };
                return _sslModes;
            }
            set { }
        }

        public string SslMode
        {
            get { return Enum.GetName(typeof(Npgsql.SslMode), _connStrBuilder.SslMode); }
            set
            {
                if (Enum.IsDefined(typeof(Npgsql.SslMode), value))
                {
                    _connStrBuilder.SslMode = (Npgsql.SslMode)Enum.Parse(typeof(Npgsql.SslMode), value);
                    if (_connStrBuilder.SslMode == Npgsql.SslMode.Require)
                        _connStrBuilder.SSL = true;
                    else if (_connStrBuilder.SslMode == Npgsql.SslMode.Disable)
                        _connStrBuilder.SSL = false;
                }

            }
        }

        #endregion

        #region Database

        public string[] Databases
        {
            get { return _databases; }
            set { }
        }

        public string Database
        {
            get { return _connStrBuilder.Database; }
            set
            {
                if (!String.IsNullOrEmpty(value) && (value != _connStrBuilder.Database))
                {
                    _connStrBuilder.Database = value;
                    LoadSchemata();
                }
            }
        }

        private void LoadDatabases()
        {
            string[] databaseList = new string[] { };
            NpgsqlConnection cn = null;
           
            try
            {
                if ((_connStrBuilder != null) && !String.IsNullOrEmpty(_connStrBuilder.Host))
                {
                    cn = new NpgsqlConnection(_connStrBuilder.ConnectionString);
                    cn.Open();

                    DataTable databases = cn.GetSchema("Databases", null);
                    databaseList = (from r in databases.AsEnumerable()
                                    let dbName = r.Field<string>("database_name")
                                    where !dbName.ToLower().StartsWith("template")
                                    select dbName).OrderBy(t => t).ToArray();
                }
            }
            catch (Exception ex)
            {
                if (ex is NpgsqlException)
                    MessageBox.Show(ex.Message, "PostgreSQL Error", MessageBoxButton.OK, MessageBoxImage.Error);
                else
                    MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if ((cn != null) && (cn.State != ConnectionState.Closed)) cn.Close();

                _databases = databaseList;
                OnPropertyChanged("Databases");

                if (_databases.Length == 1) _connStrBuilder.Database = _databases[0];
                OnPropertyChanged("Database");
            }
        }

        #endregion

        #region Encoding

        public string[] Encodings
        {
            get
            {
                if (_encodings == null)
                {
                    _encodings = new string[] { "<default>", "BIG5", "EUC_CN", "EUC_JP", "EUC_KR", "EUC_TW",
                        "GB18030", "GBK", "ISO_8859_5", "ISO_8859_6", "ISO_8859_7", "ISO_8859_8", "JOHAB",
                        "KOI8", "LATIN1", "LATIN2", "LATIN3", "LATIN4", "LATIN5", "LATIN6", "LATIN7", "LATIN8",
                        "LATIN9", "LATIN10", "MULE_INTERNAL", "SJIS", "SQL_ASCII", "UHC", "UTF8", "WIN866",
                        "WIN874", "WIN1250", "WIN1251", "WIN1252", "WIN1253", "WIN1254", "WIN1255", "WIN1256",
                        "WIN1257","WIN1258" };
                }
                return _encodings; 
            }
            set { }
        }

        public string Encoding
        {
            get { return _encoding; }
            set
            {
                if (!String.IsNullOrEmpty(value) && (value != _encoding))
                    _encoding = value;
            }
        }

        #endregion

        #region Authentication

        public string UserName
        {
            get { return _connStrBuilder.UserName; }
            set
            {
                if (!String.IsNullOrEmpty(value) && (value != _connStrBuilder.UserName))
                    _connStrBuilder.UserName = value;
            }
        }

        public string Password
        {
            get { return _connStrBuilder.Password; }
            set { if (value != _connStrBuilder.Password) _connStrBuilder.Password = value; }
        }

        #endregion

        #region Default Schema

        public string[] Schemata
        {
            get { return _schemata.ToArray(); }
            set { }
        }

        public string SearchPath
        {
            get { return _connStrBuilder.SearchPath; }
            set { if (value != _connStrBuilder.SearchPath) _connStrBuilder.SearchPath = value; }
        }

        private void LoadSchemata()
        {
            List<String> schemaList = new List<String>();
            NpgsqlConnection cn = null;

            try
            {
                if ((_connStrBuilder != null) && !String.IsNullOrEmpty(_connStrBuilder.Host))
                {
                    cn = new NpgsqlConnection(_connStrBuilder.ConnectionString);
                    cn.Open();

                    NpgsqlCommand cmd = cn.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT schema_name FROM information_schema.schemata" +
                                        " WHERE schema_name !~* '^(pg|information)_'" +
                                        " AND catalog_name = '" + _connStrBuilder.Database + "'";
                    NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(cmd);
                    DataTable dbTable = new DataTable();
                    
                    try
                    {
                        adapter.Fill(dbTable);
                        schemaList = (from r in dbTable.AsEnumerable()
                                      let schemaName = r.Field<string>("schema_name")
                                      select schemaName).OrderBy(s => s).ToList();
                        _connStrBuilder.SearchPath = DbBase.GetDefaultSchema(Backends.PostgreSql, _connStrBuilder, schemaList);
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                if (ex is NpgsqlException)
                    MessageBox.Show(ex.Message, "PostgreSQL Error", MessageBoxButton.OK, MessageBoxImage.Error);
                else
                    MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if ((cn != null) && (cn.State != ConnectionState.Closed)) cn.Close();

                _schemata = schemaList;
                OnPropertyChanged("Schemata");

                if (_schemata.Count == 1) _connStrBuilder.SearchPath = _schemata[0];
                OnPropertyChanged("SearchPath");
            }
        }

        #endregion

        #region View Events

        public void ViewEvents(IntPtr windowHandle, string propertyName)
        {
            if (windowHandle != IntPtr.Zero) _windowHandle = windowHandle;

            switch (propertyName)
            {
                case "Database":
                    LoadDatabases();
                    break;
                case "SearchPath":
                    LoadSchemata();
                    break;
            }
        }

        #endregion

        #region IDataErrorInfo Members

        string IDataErrorInfo.Error
        {
            get
            {
                StringBuilder error = new StringBuilder();

                if (String.IsNullOrEmpty(_connStrBuilder.Host))
                    error.Append(", host name");
                if (_connStrBuilder.Port == 0)
                    error.Append(", port");
                if (String.IsNullOrEmpty(_connStrBuilder.Database))
                    error.Append(", database name");
                if (String.IsNullOrEmpty(_connStrBuilder.UserName))
                    error.Append(", user name");
                if (String.IsNullOrEmpty(_connStrBuilder.SearchPath))
                    error.Append(", search path");

                if (error.Length > 0)
                    return error.Remove(0, 1).Insert(0, "Please provide").ToString();
                else
                    return null;
            }
        }

        string IDataErrorInfo.this[string columnName]
        {
            get
            {
                string error = null;

                switch (columnName)
                {
                    case "Host":
                        if (String.IsNullOrEmpty(_connStrBuilder.Host))
                            error = "Error: You must provide a host name";
                        break;
                    case "Port":
                        if (_connStrBuilder.Port == 0)
                            error = "Error: You must provide a port";
                        break;
                    case "Database":
                        if (String.IsNullOrEmpty(_connStrBuilder.Database))
                            error = "Error: You must provide a database name";
                        break;
                    case "UserName":
                        if (String.IsNullOrEmpty(_connStrBuilder.UserName))
                            error = "Error: You must provide a user name";
                        break;
                    case "SearchPath":
                        if (String.IsNullOrEmpty(_connStrBuilder.SearchPath))
                            error = "Error: You must provide a search path";
                        break;
                }

                // dirty commands registered with CommandManager so they are queried to see if they can execute now
                CommandManager.InvalidateRequerySuggested();

                return error;
            }
        }

        #endregion
    }
}
