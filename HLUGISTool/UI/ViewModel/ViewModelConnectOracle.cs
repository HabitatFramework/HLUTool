// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2013 Andy Foy
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
using HLU.Properties;
using Oracle.DataAccess.Client;

namespace HLU.UI.ViewModel
{
    class ViewModelConnectOracle : ViewModelBase, IDataErrorInfo
    {
        internal enum DBAPrivilege
        {
            Normal, SYSDBA, SYSOPER
        }

        #region private Members

        private IntPtr _windowHandle;
        private string _displayName;
        private RelayCommand _okCommand;
        private RelayCommand _cancelCommand;
        private Dictionary<string, string> _dataSourcesDic;
        private string[] _dataSources;
        private List<String> _schemata = new List<string>();
        private string _defaultSchema;
        private OracleConnectionStringBuilder _connStrBuilder;

        #endregion

        #region Constructor

        public ViewModelConnectOracle()
        {
            _connStrBuilder = new OracleConnectionStringBuilder();
        }

        #endregion

        #region Connection String Builder

        public OracleConnectionStringBuilder ConnectionStringBuilder { get { return _connStrBuilder; } }

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
        public delegate void RequestCloseEventHandler(string connString, string defaultSchema, string errMsg);

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
            OracleConnection cn;

            try
            {
                _connStrBuilder.PersistSecurityInfo = Settings.Default.DbConnectionPersistSecurityInfo;

                cn = new OracleConnection(_connStrBuilder.ConnectionString);

                cn.Open();
                cn.Close();

                this.RequestClose(_connStrBuilder.ConnectionString, _defaultSchema, null);
            }
            catch (OracleException exOra)
            {
                MessageBox.Show("Oracle Server responded with an error:\n\n" + exOra.Message,
                     "Oracle Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
                //return !((String.IsNullOrEmpty(_connStrBuilder.DataSource)) || 
                //    (!_connStrBuilder.IntegratedSecurity && String.IsNullOrEmpty(_connStrBuilder.UserID)) || 
                //    String.IsNullOrEmpty(_defaultSchema));
                return !((String.IsNullOrEmpty(_connStrBuilder.DataSource)) || String.IsNullOrEmpty(_connStrBuilder.UserID) ||
                    String.IsNullOrEmpty(_defaultSchema));
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

        #region View Events

        public void ViewEvents(IntPtr windowHandle, string propertyName)
        {
            if (windowHandle != IntPtr.Zero) _windowHandle = windowHandle;

            switch (propertyName)
            {
                case "DefaultSchema":
                    LoadSchemata();
                    break;
            }
        }

        #endregion

        #region Data Source

        public string[] DataSources
        {
            get 
            {
                if (_dataSources == null)
                {
                    OracleClientFactory factory = new OracleClientFactory();
                    if (factory.CanCreateDataSourceEnumerator)
                    {
                        System.Data.Common.DbDataSourceEnumerator dataSourceEnumarator = factory.CreateDataSourceEnumerator();
                        DataTable dt = dataSourceEnumarator.GetDataSources();
                        _dataSourcesDic = DbOracle.GetConnectionStrings(dt);
                        _dataSources = _dataSourcesDic.Keys.ToArray();
                        OnPropertyChanged("DataSources");
                    }
                    else
                    {
                        _dataSources = new string[] { };
                    }
                }
                return _dataSources;
            }
            set { }
        }

        public string DataSource
        {
            get { return _connStrBuilder.DataSource; }
            set
            {
                if (!String.IsNullOrEmpty(value) && (value != _connStrBuilder.DataSource))
                    _connStrBuilder.DataSource = value;
            }
        }

        #endregion

        #region Authentication

        public string UserID
        {
            get { return _connStrBuilder.UserID; }
            set
            {
                if (!String.IsNullOrEmpty(value) && (value != _connStrBuilder.UserID))
                {
                    _connStrBuilder.UserID = value;
                    if (String.IsNullOrEmpty(_defaultSchema) && (_schemata != null) && (_schemata.Count > 0))
                    {
                        _defaultSchema = _connStrBuilder.UserID;
                        OnPropertyChanged("DefaultSchema");
                    }
                    if (DbOracle.GetUserId(_connStrBuilder.UserID) == "SYS")
                    {
                        _connStrBuilder.DBAPrivilege = DBAPrivilege.SYSDBA.ToString();
                        OnPropertyChanged("DBAPrivilegeOption");
                    }
                }
            }
        }

        public DBAPrivilege[] DBAPrivilegeOptions
        {
            get { return (DBAPrivilege[])Enum.GetValues(typeof(DBAPrivilege)); }
            set { }
        }

        public DBAPrivilege DBAPrivilegeOption
        {
            get
            {
                if (String.IsNullOrEmpty(_connStrBuilder.DBAPrivilege)) return DBAPrivilege.Normal;
                object newValue = Enum.Parse(typeof(DBAPrivilege), _connStrBuilder.DBAPrivilege);
                if (newValue != null)
                    return (DBAPrivilege)newValue;
                else
                    return DBAPrivilege.Normal;
            }
            set
            {
                if (value == DBAPrivilege.Normal)
                    _connStrBuilder.DBAPrivilege = String.Empty;
                else
                    _connStrBuilder.DBAPrivilege = Enum.GetName(typeof(DBAPrivilege), value);
            }
        }

        public string Password
        {
            get { return _connStrBuilder.Password; }
            set
            {
                if (!String.IsNullOrEmpty(value) && (value != _connStrBuilder.Password))
                    _connStrBuilder.Password = value;
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
            set { if (value != _defaultSchema) _defaultSchema = value; }
        }

        private void LoadSchemata()
        {
            List<String> schemaList = new List<String>();
            OracleConnection cn = null;

            try
            {
                if ((_connStrBuilder != null) && !String.IsNullOrEmpty(_connStrBuilder.DataSource))
                {
                    cn = new OracleConnection(_connStrBuilder.ConnectionString);
                    cn.Open();

                    OracleCommand cmd = cn.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT username FROM all_users";
                    OracleDataAdapter adapter = new OracleDataAdapter(cmd);
                    DataTable dbTable = new DataTable();

                    try
                    {
                        adapter.Fill(dbTable);
                        schemaList = (from r in dbTable.AsEnumerable()
                                      let schemaName = r.Field<string>("username")
                                      select schemaName).OrderBy(s => s).ToList();
                        _defaultSchema = DbBase.GetDefaultSchema(Backends.Oracle, _connStrBuilder, schemaList);
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                if (ex is OracleException)
                    MessageBox.Show(ex.Message, "Oracle Error", MessageBoxButton.OK, MessageBoxImage.Error);
                else
                    MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
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

        #region IDataErrorInfo Members

        string IDataErrorInfo.Error
        {
            get
            {
                StringBuilder error = new StringBuilder();

                if (String.IsNullOrEmpty(_connStrBuilder.DataSource))
                    error.Append(", data source");
                //if (!_connStrBuilder.IntegratedSecurity && String.IsNullOrEmpty(_connStrBuilder.UserID))
                //    error.Append(", user ID");
                if (String.IsNullOrEmpty(_connStrBuilder.UserID)) error.Append(", user ID");
                if (String.IsNullOrEmpty(_defaultSchema)) error.Append(", default schema");

                if (error.Length > 0)
                    return error.Remove(0, 1).Insert(0, "Please provide a ").ToString();
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
                    case "DataSource":
                        if (String.IsNullOrEmpty(_connStrBuilder.DataSource)) 
                            error = "Please provide a data source";
                        break;
                    case "UserID":
                        //if (!_connStrBuilder.IntegratedSecurity && String.IsNullOrEmpty(_connStrBuilder.UserID))
                        if (String.IsNullOrEmpty(_connStrBuilder.UserID)) error = "Please provide a user ID";
                        break;
                    case "DefaultSchema":
                        if (String.IsNullOrEmpty(_defaultSchema))
                            error = "Please provide a default schema";
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
