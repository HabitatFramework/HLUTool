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
using System.Data.Sql;
using System.Data.SqlClient;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using HLU.Data.Connection;
using HLU.Properties;

namespace HLU.UI.ViewModel
{
    class ViewModelConnectSqlServer : ViewModelBase, IDataErrorInfo
    {
        #region private Members

        private IntPtr _windowHandle;
        private string _displayName;
        private RelayCommand _okCommand;
        private RelayCommand _cancelCommand;
        private List<String> _servers;
        private List<String> _databases = new List<String>();
        private List<String> _schemata = new List<String>();
        private string _defaultSchema;
        private SqlConnectionStringBuilder _connStrBuilder;

        #endregion

        #region Constructor

        public ViewModelConnectSqlServer()
        {
            _connStrBuilder = new SqlConnectionStringBuilder();
        }

        #endregion

        #region Connection String Builder

        public SqlConnectionStringBuilder ConnectionStringBuilder { get { return _connStrBuilder; } }

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
        public delegate void RequestCloseEventHandler(string connString, string defaultSchema, string errorMsg);

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
            SqlConnection cn;

            try
            {
                cn = new SqlConnection(_connStrBuilder.ConnectionString);

                cn.Open();
                cn.Close();

                _connStrBuilder.PersistSecurityInfo = Settings.Default.DbConnectionPersistSecurityInfo;
                
                this.RequestClose(_connStrBuilder.ConnectionString, _defaultSchema, null);
            }
            catch (SqlException exSql)
            {
                MessageBox.Show("SQL Server responded with an error:\n\n" + exSql.Message,
                     "SQL Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
                return !(String.IsNullOrEmpty(_connStrBuilder.DataSource) || String.IsNullOrEmpty(_connStrBuilder.InitialCatalog) ||
                    (!_connStrBuilder.IntegratedSecurity && (String.IsNullOrEmpty(_connStrBuilder.Password) ||
                    String.IsNullOrEmpty(_connStrBuilder.UserID))) && !String.IsNullOrEmpty(_defaultSchema));
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
            this.RequestClose(null, null, null);
        }

        #endregion

        #region Server

        public string Server
        {
            get { return _connStrBuilder.DataSource; }
            set
            {
                if (!String.IsNullOrEmpty(value) && (value != _connStrBuilder.DataSource))
                    _connStrBuilder.DataSource = value;
            }
        }

        public string[] Servers
        {
            get
            {
                if (_servers == null) _servers = LoadServers();
                return _servers.ToArray();
            }
            set { }
        }

        private List<string> LoadServers()
        {
            try
            {
                // Retrieve enumerator instance and then the data
                SqlDataSourceEnumerator instance = SqlDataSourceEnumerator.Instance;
                DataTable table = instance.GetDataSources();
                List<string> serverList = new List<string>();

                // Display contents of table
                foreach (DataRow row in table.Rows)
                    serverList.Add(row["ServerName"].ToString() + @"\" + row["InstanceName"].ToString());

                return serverList;
            }
            catch (Exception ex)
            {
                MessageBox.Show("SQL Server responded with an error: " + ex.Message,
                    "SQL Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return new List<string>();
            }
        }

        #endregion

        #region Authentication

        public bool WindowsAuthentication
        {
            get { return _connStrBuilder.IntegratedSecurity; }
            set
            {
                _connStrBuilder.IntegratedSecurity = value;
                OnPropertyChanged("SQLServerAuthentication");
                OnPropertyChanged("WindowsAuthentication");
                OnPropertyChanged("Username");
            }
        }

        public bool SQLServerAuthentication
        {
            get { return !_connStrBuilder.IntegratedSecurity; }
            set
            {
                _connStrBuilder.IntegratedSecurity = !value;
                OnPropertyChanged("SQLServerAuthentication");
                OnPropertyChanged("WindowsAuthentication");
                OnPropertyChanged("Username");
            }
        }

        public string Username
        {
            get { return _connStrBuilder.UserID; }
            set { if (value != _connStrBuilder.UserID) _connStrBuilder.UserID = value; }
        }

        public string Password
        {
            get { return _connStrBuilder.Password; }
            set { if (value != _connStrBuilder.Password) _connStrBuilder.Password = value; }
        }

        #endregion

        #region Database

        public string Database
        {
            get { return _connStrBuilder.InitialCatalog; }
            set
            {
                if (!String.IsNullOrEmpty(value) && (value != _connStrBuilder.InitialCatalog))
                {
                    _connStrBuilder.InitialCatalog = value;
                    LoadSchemata();
                }
            }
        }

        public string[] Databases
        {
            get { return _databases.ToArray(); }
            set { }
        }

        private void LoadDatabases()
        {
            try
            {
                if (_connStrBuilder != null)
                {
                    SqlConnection cn = new SqlConnection(_connStrBuilder.ConnectionString);

                    List<String> DatabaseList = new List<String>();
                    cn.Open();

                    DataTable dbTable = cn.GetSchema("Databases");
                    _databases = (from r in dbTable.AsEnumerable()
                                  let tableName = r.Field<string>("database_name")
                                  select tableName).OrderBy(t => t).ToList();

                    OnPropertyChanged("Databases");
                }
            }
            catch (Exception ex)
            {
                _databases = new List<String>();
                MessageBox.Show("SQL Server responded with an error: " + ex.Message,
                    "SQL Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Default Schema

        public string[] Schemata
        {
            get { return _schemata.ToArray(); }
            set { }
        }

        public string DefaultSchema
        {
            get { return _defaultSchema; }
            set
            {
                if (!String.IsNullOrEmpty(value) && (value != _defaultSchema))
                    _defaultSchema = value;
            }
        }

        private void LoadSchemata()
        {
            List<String> schemaList = new List<String>();
            SqlConnection cn = null;

            try
            {
                if ((_connStrBuilder != null) && !String.IsNullOrEmpty(_connStrBuilder.InitialCatalog))
                {
                    cn = new SqlConnection(_connStrBuilder.ConnectionString);
                    cn.Open();

                    DataTable dbTable = cn.GetSchema("Users");
                    schemaList = (from r in dbTable.AsEnumerable()
                                  let schemaName = r.Field<string>("user_name")
                                  select schemaName).OrderBy(s => s).ToList();
                    _defaultSchema = DbBase.GetDefaultSchema(Backends.SqlServer, _connStrBuilder, schemaList);
                }
            }
            catch (Exception ex)
            {
                if (ex is SqlException)
                    MessageBox.Show(ex.Message, "SQL Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
                else
                    MessageBox.Show("SQL Server responded with an error: " + ex.Message,
                    "SQL Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if ((cn != null) && (cn.State != ConnectionState.Closed)) cn.Close();

                _schemata = schemaList;
                OnPropertyChanged("Schemata");

                if (_schemata.Count == 1) _defaultSchema = _schemata[0];
                OnPropertyChanged("DefaultSchema");
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
            }
        }

        #endregion

        #region IDataErrorInfo Members

        public string Error
        {
            get
            {
                string error = null;

                if (String.IsNullOrEmpty(_connStrBuilder.DataSource) || String.IsNullOrEmpty(_connStrBuilder.InitialCatalog))
                    error = "Error: You must provide at least server name and database";


                if (!_connStrBuilder.IntegratedSecurity && String.IsNullOrEmpty(_connStrBuilder.UserID))
                    error = "Error: You must provide user id (and usually password) if using SQL Server authentication";

                return error;
            }
        }

        public string this[string columnName]
        {
            get
            {
                string error = null;

                switch (columnName)
                {
                    case "Server":
                        if (String.IsNullOrEmpty(_connStrBuilder.DataSource))
                            error = "Error: You must choose a server";
                        break;
                    case "Username":
                        if ((!_connStrBuilder.IntegratedSecurity) && (String.IsNullOrEmpty(_connStrBuilder.UserID)))
                            error = "Error: You must provide a user id";
                        break;
                    case "Database":
                        if (String.IsNullOrEmpty(_connStrBuilder.InitialCatalog))
                            error = "Error: You must choose a database";
                        break;
                    case "DefaultSchema":
                        if (String.IsNullOrEmpty(_defaultSchema))
                            error = "Error: You must choose a default schema";
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
