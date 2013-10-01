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
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using HLU.Data.Connection;
using Microsoft.Win32;

namespace HLU.UI.ViewModel
{
    class ViewModelConnectOdbc : ViewModelBase, IDataErrorInfo
    {
        #region private Members

        private IntPtr _windowHandle;
        private string _displayName;
        private RelayCommand _okCommand;
        private RelayCommand _cancelCommand;
        private RelayCommand _manageDsnCommand;
        private string[] _dsnList;
        private bool _userDsn = true;
        private bool _systemDsn;
        private Backends _backend = Backends.UndeterminedOdbc;
        private string _defaultSchema;
        private List<String> _schemata;

        private OdbcConnectionStringBuilder _connStrBuilder;

        #endregion

        #region Constructor

        public ViewModelConnectOdbc()
        {
            _connStrBuilder = new OdbcConnectionStringBuilder();
        }

        #endregion

        #region Connection String Builder

        public OdbcConnectionStringBuilder ConnectionStringBuilder { get { return _connStrBuilder; } }

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
            OdbcConnection cn;

            try
            {
                cn = new OdbcConnection(_connStrBuilder.ConnectionString);

                cn.Open();
                cn.Close();

                if (DbOdbc.GetBackend(cn) == Backends.Access) _defaultSchema = String.Empty;

                this.RequestClose(_connStrBuilder.ConnectionString, _defaultSchema, null);
            }
            catch (OdbcException exOdbc)
            {
                MessageBox.Show("ODBC Server responded with an error:\n\n" + exOdbc.Message,
                     "ODBC Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
                return !String.IsNullOrEmpty(_connStrBuilder.ConnectionString) &&
                    ((DbOdbc.GetBackend(_connStrBuilder) == Backends.Access) || !String.IsNullOrEmpty(_defaultSchema));
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

        #region Manage DSN Command

        /// <summary>
        /// Create Manage DSN button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand ManageDsnCommand
        {
            get
            {
                if (_manageDsnCommand == null)
                {
                    Action<object> manageDsnAction = new Action<object>(this.ManageDsnCommandClick);
                    _manageDsnCommand = new RelayCommand(manageDsnAction);
                }

                return _manageDsnCommand;
            }
        }

        /// <summary>
        /// Handles event when Manage DSN button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void ManageDsnCommandClick(object param)
        {
            //DispatcherHelper.DoEvents();
            OdbcCP32 odbccp32 = new OdbcCP32();
            bool result = odbccp32.ManageDatasources(_windowHandle);
            OnPropertyChanged("DsnList");
        }

        #endregion

        #region DSN

        public string[] DsnList
        {
            get
            {
                RegistryKey rk = _userDsn ? Registry.CurrentUser : Registry.LocalMachine;
                RegistryKey sk = rk.OpenSubKey(@"SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources");
                if (sk != null)
                    _dsnList = sk.GetValueNames();
                else
                    _dsnList = new string[0];
                return _dsnList; 
            }
            set { }
        }

        public string Dsn
        {
            get { return _connStrBuilder.Dsn; }
            set
            {
                if (!String.IsNullOrEmpty(value) && (value != _connStrBuilder.Dsn))
                {
                    _connStrBuilder.Dsn = value;
                    OnPropertyChanged("SupportsSchemata");
                }
            }
        }

        public bool UserDsn
        {
            get { return _userDsn; }
            set
            {
                _userDsn = value;
                _systemDsn = !value;
                OnPropertyChanged("UserDsn");
                OnPropertyChanged("SystemDsn");
                OnPropertyChanged("DsnList");
            }
        }

        public bool SystemDsn
        {
            get { return _systemDsn; }
            set
            {
                _systemDsn = value;
                _userDsn = !value;
                OnPropertyChanged("UserDsn");
                OnPropertyChanged("SystemDsn");
                OnPropertyChanged("DsnList");
            }
        }

        #endregion

        #region Default Schema

        public bool SupportsSchemata
        {
            get
            {
                LoadSchemata();
                return _backend != Backends.Access;
            }
        }

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
            OdbcConnection cn = null;

            try
            {
                if ((_connStrBuilder != null) && !String.IsNullOrEmpty(_connStrBuilder.ConnectionString))
                {
                    cn = new OdbcConnection(_connStrBuilder.ConnectionString);
                    cn.Open();

                    _backend = DbOdbc.GetBackend(cn);

                    if (_backend != Backends.Access)
                    {
                        OdbcCommand cmd = cn.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "SELECT SCHEMA_NAME FROM information_schema.schemata" +
                                            " WHERE SCHEMA_NAME <> 'INFORMATION_SCHEMA'";
                        OdbcDataAdapter adapter = new OdbcDataAdapter(cmd);
                        DataTable dbTable = new DataTable();
                        try
                        {
                            adapter.Fill(dbTable);
                            schemaList = (from r in dbTable.AsEnumerable()
                                          let schemaName = r.Field<string>("SCHEMA_NAME")
                                          select schemaName).OrderBy(s => s).ToList();
                            _defaultSchema = DbBase.GetDefaultSchema(_backend, _connStrBuilder, schemaList);
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                _connStrBuilder.Dsn = String.Empty;
                if (ex is OdbcException)
                    MessageBox.Show(ex.Message, "ODBC Error", MessageBoxButton.OK, MessageBoxImage.Error);
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

        #region View Events

        public void ViewEvents(IntPtr windowHandle, string propertyName)
        {
            if (windowHandle != IntPtr.Zero) _windowHandle = windowHandle;
        }

        #endregion

        #region IDataErrorInfo Members

        string IDataErrorInfo.Error
        {
            get
            {
                StringBuilder error = new StringBuilder();

                if (String.IsNullOrEmpty(_connStrBuilder.Dsn))
                    error.Append(", a data source");
                if ((_backend != Backends.Access) && String.IsNullOrEmpty(_defaultSchema))
                    error.Append(" , default schema");

                if (error.Length > 1)
                    return error.Remove(0, 1).Insert(0, "Please choose ").ToString();
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
                    case "Dsn":
                        if (String.IsNullOrEmpty(_connStrBuilder.Dsn)) 
                            error = "Please choose a data source";
                        break;
                    case "DefaultSchema":
                        if ((_backend != Backends.Access) && String.IsNullOrEmpty(_defaultSchema))
                            error = "Please choose a default schema";
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
