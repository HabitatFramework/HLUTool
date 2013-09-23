using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using HLU.Data.Connection;
using HLU.Properties;
using Microsoft.Win32;
using MSDASC;

namespace HLU.UI.ViewModel
{
    class ViewModelConnectOleDb : ViewModelBase, IDataErrorInfo
    {
        #region private Members

        private IntPtr _windowHandle;
        private string _displayName;
        private RelayCommand _okCommand;
        private RelayCommand _cancelCommand;
        private RelayCommand _editConnCommand;
        private RelayCommand _createConnCommand;
        private RelayCommand _browseConnCommand;
        private ADODB.Connection _connAdo;
        private Backends _backend = Backends.UndeterminedOleDb;
        private List<String> _schemata = new List<string>();
        private string _defaultSchema;

        private OleDbConnectionStringBuilder _connStrBuilder;

        #endregion

        #region Constructor

        public ViewModelConnectOleDb()
        {
            _connStrBuilder = new OleDbConnectionStringBuilder();
        }

        #endregion

        #region Connection String Builder

        public OleDbConnectionStringBuilder ConnectionStringBuilder { get { return _connStrBuilder; } }

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
            OleDbConnection cn;

            try
            {
                _connStrBuilder.PersistSecurityInfo = Settings.Default.DbConnectionPersistSecurityInfo;

                cn = new OleDbConnection(_connStrBuilder.ConnectionString);

                cn.Open();
                cn.Close();

                if (DbOleDb.GetBackend(cn) == Backends.Access) _defaultSchema = String.Empty;

                this.RequestClose(_connStrBuilder.ConnectionString, _defaultSchema, null);
            }
            catch (OleDbException exOleDb)
            {
                MessageBox.Show("OleDb Server responded with an error:\n\n" + exOleDb.Message,
                     "OleDb Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
                    (IsMsAccess(_connAdo) || !String.IsNullOrEmpty(_defaultSchema));
            }
        }

        private bool IsMsAccess(ADODB.Connection connection)
        {
            if (connection == null)
                return false;
            else
                return connection.Provider.ToLower().StartsWith("microsoft.jet.oledb");
        }

        private bool IsSqlServer(ADODB.Connection connection)
        {
            if (connection == null)
                return false;
            else
                return connection.Provider.ToLower().StartsWith("sqloledb");
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

        #region Create Connection Command

        /// <summary>
        /// Create Create Connection button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand CreateConnCommand
        {
            get
            {
                if (_createConnCommand == null)
                {
                    Action<object> createConnAction = new Action<object>(this.CreateConnCommandClick);
                    _createConnCommand = new RelayCommand(createConnAction);
                }

                return _createConnCommand;
            }
        }

        /// <summary>
        /// Handles event when Create Connection button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void CreateConnCommandClick(object param)
        {
            DataLinks udl = new DataLinksClass();
            udl.hWnd = _windowHandle != IntPtr.Zero ? _windowHandle.ToInt32() : 
                new WindowInteropHelper(App.Current.MainWindow).Handle.ToInt32();
            _connAdo = udl.PromptNew() as ADODB.Connection;

            if ((_connAdo != null) && TestConnection(_connAdo.ConnectionString))
            {
                _connStrBuilder.ConnectionString = _connAdo.ConnectionString;
                OnPropertyChanged("ConnectionString");
                OnPropertyChanged("SupportsSchemata");
            }
        }

        #endregion

        #region Browse Connection Command

        /// <summary>
        /// Create Browse Connection button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand BrowseConnCommand
        {
            get
            {
                if (_browseConnCommand == null)
                {
                    Action<object> browseConnAction = new Action<object>(this.BrowseConnCommandClick);
                    _browseConnCommand = new RelayCommand(browseConnAction);
                }

                return _browseConnCommand;
            }
        }

        /// <summary>
        /// Handles event when Browse Connection button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void BrowseConnCommandClick(object param)
        {
            OpenFileDialog openFileDlg = new OpenFileDialog();
            openFileDlg.Filter = "Microsoft Data Links (*.udl)|*.udl";
            openFileDlg.Multiselect = false;
            openFileDlg.RestoreDirectory = true;

            if (openFileDlg.ShowDialog() != true) return;

            string testString = "File Name = " + openFileDlg.FileName;
            if (TestConnection(testString))
            {
                _connStrBuilder.ConnectionString = testString;
                OnPropertyChanged("ConnectionString");
                OnPropertyChanged("SupportsSchemata");
            }
        }

        #endregion

        #region Edit Connection Command

        /// <summary>
        /// Create Edit Connection button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand EditConnCommand
        {
            get
            {
                if (_editConnCommand == null)
                {
                    Action<object> editConnAction = new Action<object>(this.EditConnCommandClick);
                    _editConnCommand = new RelayCommand(editConnAction, param => this.CanEditConn);
                }

                return _editConnCommand;
            }
        }

        /// <summary>
        /// Handles event when Edit Connection button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void EditConnCommandClick(object param)
        {
            try
            {
                object adoDbConn = (object)_connAdo;
                
                DataLinks udl = new DataLinksClass();
                udl.hWnd = _windowHandle != IntPtr.Zero ? _windowHandle.ToInt32() :
                    new WindowInteropHelper(App.Current.MainWindow).Handle.ToInt32();

                if (udl.PromptEdit(ref adoDbConn))
                {
                    if (TestConnection(_connAdo.ConnectionString))
                    {
                        _connStrBuilder.ConnectionString = _connAdo.ConnectionString;
                        OnPropertyChanged("ConnectionString");
                        OnPropertyChanged("SupportsSchemata");
                    }
                }
            }
            catch { _connAdo = null; }
        }

        /// <summary>
        /// Determines whether the Edit Connection button is enabled
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanEditConn { get { return _connAdo != null; } }

        #endregion

        #region Connection String

        public string ConnectionString
        {
            get { return HLU.Data.Connection.DbBase.MaskPassword(_connStrBuilder, Resources.PasswordMaskString); }
            set { }
        }

        private bool TestConnection(string connectionString)
        {
            bool success = true;

            try
            {
                OleDbConnection cn = new OleDbConnection(connectionString);
                cn.Open();
                cn.Close();

                _connAdo = new ADODB.Connection();
                _connAdo.ConnectionString = connectionString;
                _connAdo.Provider = cn.Provider;
            }
            catch (OleDbException exOleDb)
            {
                success = false;
                MessageBox.Show("OleDb Server responded with an error:\n\n" + exOleDb.Message,
                    "OleDb Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return success;
        }

        #endregion

        #region Default Schema

        public bool SupportsSchemata
        {
            get
            {
                LoadSchemata();
                return !IsMsAccess(_connAdo);
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
            OleDbConnection cn = null;

            try
            {
                if ((_connStrBuilder != null) && !String.IsNullOrEmpty(_connStrBuilder.ConnectionString))
                {
                    cn = new OleDbConnection(_connStrBuilder.ConnectionString);
                    cn.Open();

                    _backend = DbOleDb.GetBackend(cn);

                    if (_backend != Backends.Access)
                    {
                        OleDbCommand cmd = cn.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "SELECT SCHEMA_NAME FROM INFORMATION_SCHEMA.SCHEMATA" +
                                            " WHERE SCHEMA_NAME <> 'INFORMATION_SCHEMA'";
                        OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);
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
                _connStrBuilder.ConnectionString = String.Empty;
                if (ex is OleDbException)
                    MessageBox.Show(ex.Message, "OleDb Error", MessageBoxButton.OK, MessageBoxImage.Error);
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

                if (String.IsNullOrEmpty(_connStrBuilder.ConnectionString))
                    error.Append(", connection");
                if ((_connAdo != null) && !_connAdo.Provider.StartsWith("Microsoft.Jet.OLEDB") && 
                    String.IsNullOrEmpty(_defaultSchema)) error.Append(", default schema");

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
                    case "ConnectionString":
                        if (String.IsNullOrEmpty(_connStrBuilder.ConnectionString))
                            error = "Please create a connection";
                        break;
                    case "DefaultSchema":
                        if ((_connAdo != null) && !_connAdo.Provider.StartsWith("Microsoft.Jet.OLEDB") && 
                            String.IsNullOrEmpty(_defaultSchema)) error = "Please provide a default schema";
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
