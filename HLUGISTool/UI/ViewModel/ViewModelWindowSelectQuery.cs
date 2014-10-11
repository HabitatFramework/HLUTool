// HLUTool is used to view and maintain habitat and land use GIS data.
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
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using HLU.Data.Model;
using HLU.Data.Connection;
using HLU.UI.ViewModel;
using HLU.GISApplication;
using HLU.Properties;
using HLU.UI.UserControls;

namespace HLU.UI.ViewModel
{
    class ViewModelWindowSelectQuery : ViewModelBase, IDataErrorInfo
    {
        public static HluDataSet HluDatasetStatic = null;

        #region Fields

        private ICommand _getValuesCommand;

        private ICommand _clearCommand;
        private ICommand _verifyCommand;
        private ICommand _loadCommand;
        private ICommand _saveCommand;

        private ICommand _okCommand;
        private ICommand _cancelCommand;

        private ICommand _addTableCommand;
        private ICommand _addColumnCommand;
        private ICommand _addOperatorCommand;
        private ICommand _addValueCommand;

        private string _displayName = "Advanced Query Builder";
        private Cursor _cursorType = Cursors.Arrow;

        private Dictionary<string, DataTable> _tables;
        private DataTable _table;
        private DataColumn _column;
        private DbBase _db;

        private string[] _comparisonOperators;
        private string _comparisonOperator;
        private Dictionary<string, object> _queryValues;
        private string _queryValueText;
        private object _queryValue;

        private string _sqlFromTables;
        private string _sqlWhereClause;

        private string _descriptionFieldName = Settings.Default.LutDescriptionFieldName;
        private int _descriptionFieldOrdinal = Settings.Default.LutDescriptionFieldOrdinal;
        private Regex _queryValueRegex = new Regex(@"\s+:\s+", RegexOptions.IgnoreCase); // @"\A(?<code>[^:\s]+)\s+:\s+(?<desc>[^:]+)\z", RegexOptions.IgnoreCase);

        #endregion

        #region Constructor

        public ViewModelWindowSelectQuery(HluDataSet hluDataset, DbBase hluDatabase)
        {
            HluDatasetStatic = hluDataset;
            _db = hluDatabase;
        }

        #endregion

        #region ViewModelBase Members

        public override string DisplayName
        {
            get { return _displayName; }
            set { _displayName = value; }
        }

        public override string WindowTitle
        {
            get { return _displayName; }
        }

        public Cursor CursorType
        {
            get { return _cursorType; }
        }

        #endregion

        #region RequestClose

        // declare the delegate since using non-generic pattern
        public delegate void RequestCloseEventHandler(string sqlFromTables, string sqlWhereClause);

        // declare the event
        public event RequestCloseEventHandler RequestClose;

        #endregion

        #region GetValues Command

        /// <summary>
        /// Create GetValues button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand GetValuesCommand
        {
            get
            {
                if (_getValuesCommand == null)
                {
                    Action<object> getValuesAction = new Action<object>(this.GetValuesCommandClick);
                    _getValuesCommand = new RelayCommand(getValuesAction, param => this.CanGetValues);
                }

                return _getValuesCommand;
            }
        }

        /// <summary>
        /// Handles event when GetValues button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void GetValuesCommandClick(object param)
        {
            // Return null if there is no table or column selected or
            // if the comparison operator = 'IS NULL' or 'IS NOT NULL'
            if ((Table != null) && (Column != null))
            {
                // Show the wait cursor whilst loading the values.
                _cursorType = Cursors.Wait;
                OnPropertyChanged("CursorType");


                try
                {
                    IDataReader dataReader = _db.ExecuteReader(String.Format(
                        "SELECT DISTINCT {0} FROM {1}",
                        _db.QuoteIdentifier(Column.ColumnName),
                        _db.QualifyTableName(Table.TableName)),
                        _db.Connection.ConnectionTimeout, CommandType.Text);

                    if (dataReader == null) throw new Exception(String.Format("Error reading values from {0}.{1}", Table.TableName, Column.ColumnName));

                    Dictionary<string, object> q = new Dictionary<string, object>();
                    while (dataReader.Read())
                        q.Add(dataReader.GetValue(0).ToString(), dataReader.GetValue(0));

                    dataReader.Close();

                    _queryValues = q;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    //throw;
                }


                //DataView view = new DataView(Table);
                //DataTable distinctValues = new DataTable();
                //distinctValues = view.ToTable(true, Column.ColumnName);
                //var q = distinctValues.AsEnumerable();
                //_queryValues = q.ToDictionary(r => r[0].ToString(), r => r[0]);

            }
            else
                _queryValues = null;

            OnPropertyChanged("QueryValues");
            OnPropertyChanged("QueryValueIsEnabled");

        }

        /// <summary>
        /// Determine if the GetValues button can be clicked.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanGetValues
        {
            get
            {
                return (Column != null && QueryValues == null);
            }
        }

        #endregion

        #region Clear Command

        /// <summary>
        /// Create Clear button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand ClearCommand
        {
            get
            {
                if (_clearCommand == null)
                {
                    Action<object> clearAction = new Action<object>(this.ClearCommandClick);
                    _clearCommand = new RelayCommand(clearAction, param => this.CanClear);
                }

                return _clearCommand;
            }
        }

        /// <summary>
        /// Handles event when Clear button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void ClearCommandClick(object param)
        {
            SqlFromTables = null;
            SqlWhereClause = null;
            OnPropertyChanged("SqlFromTables");
            OnPropertyChanged("SqlWhereClause");
        }

        /// <summary>
        /// Determine if the Clear button can be clicked.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanClear { get { return (!String.IsNullOrEmpty(SqlFromTables) || !String.IsNullOrEmpty(SqlWhereClause)); } }

        #endregion

        #region Verify Command

        /// <summary>
        /// Create Verify button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand VerifyCommand
        {
            get
            {
                if (_verifyCommand == null)
                {
                    Action<object> verifyAction = new Action<object>(this.VerifyCommandClick);
                    _verifyCommand = new RelayCommand(verifyAction, param => this.CanVerify);
                }

                return _verifyCommand;
            }
        }

        /// <summary>
        /// Handles event when Verify button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void VerifyCommandClick(object param)
        {
        }

        /// <summary>
        /// Determine if the Verify button can be clicked.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanVerify { get { return (!String.IsNullOrEmpty(SqlFromTables) && !String.IsNullOrEmpty(SqlWhereClause)); } }

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
            this.RequestClose(SqlFromTables, SqlWhereClause);
        }

        /// <summary>
        /// Determine if the Ok button can be clicked.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanOk { get { return String.IsNullOrEmpty(Error); } }

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
            this.RequestClose(null, null);
        }

        #endregion

        #region Load Command

        /// <summary>
        /// Create Load button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand LoadCommand
        {
            get
            {
                if (_loadCommand == null)
                {
                    Action<object> loadAction = new Action<object>(this.LoadCommandClick);
                    _loadCommand = new RelayCommand(loadAction);
                }

                return _loadCommand;
            }
        }

        /// <summary>
        /// Handles event when Load button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void LoadCommandClick(object param)
        {
            LoadSQLQuery();
        }

        #endregion

        #region Save Command

        /// <summary>
        /// Create Save button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand SaveCommand
        {
            get
            {
                if (_saveCommand == null)
                {
                    Action<object> saveAction = new Action<object>(this.SaveCommandClick);
                    _saveCommand = new RelayCommand(saveAction);
                }

                return _saveCommand;
            }
        }

        /// <summary>
        /// Handles event when Save button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void SaveCommandClick(object param)
        {
            SaveSQLQuery();
        }

        #endregion

        #region Table

        public Dictionary<string, DataTable> Tables
        {
            get
            {
                if ((_tables == null) && ((ViewModelWindowSelectQuery.HluDatasetStatic != null)))
                {
                    _tables = ViewModelWindowSelectQuery.HluDatasetStatic.incid.ChildRelations
                        .Cast<DataRelation>().ToDictionary(r => r.ChildTable.TableName, r => r.ChildTable);
                    _tables.Add("incid", ViewModelWindowSelectQuery.HluDatasetStatic.incid);
                }
                return _tables;
            }
            set { }
        }

        public DataTable Table
        {
            get { return _table; }
            set
            {
                _table = value;
                OnPropertyChanged("Table");
                OnPropertyChanged("Columns");
                OnPropertyChanged("ColumnIsEnabled");
            }
        }

        #endregion

        #region Column

        public Dictionary<string, DataColumn> Columns
        {
            get
            {
                if (Table != null)
                    return Table.Columns.Cast<DataColumn>().ToDictionary(c => c.ColumnName, c => c);
                else
                    return null;
            }
            set { }
        }

        public DataColumn Column
        {
            get { return _column; }
            set
            { 
                DataColumn column;
                column = value;

                if ((Table != null) && (column != null) &&
                    Table.Columns.Contains(column.ColumnName))
                {
                    _column = column;
                    OnPropertyChanged("Column");
                    _queryValues = null;
                    OnPropertyChanged("QueryValues");
                    OnPropertyChanged("QueryValueIsEnabled");
                }
            }
        }

        public bool ColumnIsEnabled
        {
            get { return Table != null; }
        }

        #endregion

        #region Operator

        public string[] ComparisonOperators
        {
            get 
            {
                if (_comparisonOperators == null)
                {
                    _comparisonOperators = new string[] { "=", "<", ">", "<=", ">=", "<>", "(", ")", "AND", "OR",
                        "contains", "begins with", "ends with", "IS NULL", "IS NOT NULL", "LIKE", "IN ()", "NOT IN ()" };
                }
                return _comparisonOperators;
            }
        }

        public string ComparisonOperator
        {
            get { return _comparisonOperator; }
            set
            {
                _comparisonOperator = value;
                OnPropertyChanged("ComparisonOperator");
                OnPropertyChanged("QueryValues");
                OnPropertyChanged("QueryValueIsEnabled");
            }
        }

        public bool ComparisonOperatorIsEnabled
        {
            get { return Column != null; }
        }

        #endregion

        #region Value

        public Dictionary<string, object> QueryValues
        {
            get
            {
                if (_queryValues == null)
                {
                    // Return null if there is no table or column selected.
                    if ((Table == null) || (Column == null))
                    {
                        _queryValues = null;
                        return _queryValues;
                    }

                    // Show the wait cursor whilst loading the values.
                    _cursorType = Cursors.Wait;
                    OnPropertyChanged("CursorType");

                    // Find the related lookup tables for the selected table and column
                    IEnumerable<DataRelation> parentRelations = Table.ParentRelations.Cast<DataRelation>();
                    IEnumerable<DataRelation> lutRelations = parentRelations.Where(r => r.ChildTable == Table &&
                        r.ParentTable.TableName.StartsWith("lut_", StringComparison.CurrentCultureIgnoreCase) &&
                        r.ChildColumns.Length == 1 && r.ChildColumns.Contains(Column));

                    DataRelation lutRelation;

                    // If there is only one related lookup table then load the
                    // column and description values.
                    if (lutRelations.Count() == 1)
                    {
                        lutRelation = lutRelations.ElementAt(0);
                        DataTable lut = lutRelation.ParentTable;
                        DataColumn lutColumn = lutRelation.ParentColumns[0];
                        DataColumn testColumn = null;
                        DataColumn descriptionColumn = lut.Columns.Contains(_descriptionFieldName) ?
                            lut.Columns[_descriptionFieldName] : (testColumn =
                            lut.Columns.Cast<DataColumn>().FirstOrDefault(c => c.ColumnName.IndexOf("_name",
                                StringComparison.CurrentCultureIgnoreCase) != -1)) != null ? testColumn :
                            _descriptionFieldOrdinal < lut.Columns.Count ?
                            lut.Columns[_descriptionFieldOrdinal] : lutColumn.Ordinal < lut.Columns.Count - 1 ?
                            lut.Columns[lutColumn.Ordinal + 1] : lutColumn;

                        // If the select table is the 'incid_bap' table then only
                        // 'PHAP' values from the related lookup table (lut_habitat_type).
                        var q = lut is HluDataSet.lut_habitat_typeDataTable && Table is HluDataSet.incid_bapDataTable ?
                            lut.AsEnumerable().Where(r => Regex.IsMatch(r[lut.PrimaryKey[0].Ordinal].ToString(), @"\APHAP")) :
                            lut.AsEnumerable();

                        _queryValues = q.ToDictionary(r => r[lutColumn].ToString() + (descriptionColumn != lutColumn ?
                            " : " + r[descriptionColumn].ToString() : String.Empty), r => r[lutColumn]);
                        return _queryValues;
                    }
                    else
                    {
                        _queryValues = null;
                    }
                }

                return _queryValues;
            }
            set
            {
                _queryValues = value;
            }
        }

        public object QueryValue
        {
            get { return _queryValue != null ? _queryValue.ToString() : null; }
            set
            {
                //if (String.IsNullOrEmpty(_queryValueText) || (_queryValueRegex.Split(_queryValueText).Length > 1))
                    _queryValue = value;

                _cursorType = Cursors.Arrow;
                OnPropertyChanged("CursorType");
                OnPropertyChanged("QueryValue");
                OnPropertyChanged("QueryValueText");
            }
        }

        //public string QueryValueText
        //{
        //    get { return _queryValue != null ? _queryValue : null; }
        //    set
        //    {
        //        _cursorType = Cursors.Arrow;
        //        OnPropertyChanged("CursorType");

        //        if ((_queryValueText == null) || !_queryValueText.Equals(value))
        //        {
        //            if (!String.IsNullOrEmpty(value))
        //                _queryValueText = _queryValueRegex.Split(value)[0];
        //            else
        //                _queryValueText = value;

        //            _queryValue = _queryValueText;
        //            OnPropertyChanged("QueryValueText");
        //        }
        //    }
        //}

        public bool QueryValueIsEnabled
        {
            get { return _queryValues != null; }
        }

        #endregion

        #region GetValues

        #endregion

        #region Add Buttons

        #region Add Table

        /// <summary>
        /// Create AddTable button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand AddTableCommand
        {
            get
            {
                if (_addTableCommand == null)
                {
                    Action<object> addTableAction = new Action<object>(this.AddTableCommandClick);
                    _addTableCommand = new RelayCommand(addTableAction, param => this.CanAddTable);
                }

                return _addTableCommand;
            }
        }

        /// <summary>
        /// Handles event when AddTable button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void AddTableCommandClick(object param)
        {
            if (string.IsNullOrEmpty(SqlFromTables) || SqlFromTables.TrimEnd(' ').EndsWith(","))
                SqlFromTables += Table.TableName;
            else
                SqlFromTables += String.Concat(", ", Table.TableName);

            OnPropertyChanged("SqlFromTables");
        }

        /// <summary>
        /// Determine if the AddTable button can be clicked.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanAddTable
        { 
            get { return (Table != null); } 
        }

        #endregion

        #region Add Column

        /// <summary>
        /// Create AddColumn button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand AddColumnCommand
        {
            get
            {
                if (_addColumnCommand == null)
                {
                    Action<object> addColumnAction = new Action<object>(this.AddColumnCommandClick);
                    _addColumnCommand = new RelayCommand(addColumnAction, param => this.CanAddColumn);
                }

                return _addColumnCommand;
            }
        }

        /// <summary>
        /// Handles event when AddColumn button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void AddColumnCommandClick(object param)
        {
            if (string.IsNullOrEmpty(SqlWhereClause) || SqlWhereClause.EndsWith(" "))
                SqlWhereClause += _db.QuoteIdentifier(Table.TableName) + "." + _db.QuoteIdentifier(Column.ColumnName);
            else
                SqlWhereClause += String.Concat(" ", _db.QuoteIdentifier(Table.TableName), ".", _db.QuoteIdentifier(Column.ColumnName));

            OnPropertyChanged("SqlWhereClause");
        }

        /// <summary>
        /// Determine if the AddColumn button can be clicked.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanAddColumn
        { 
            get { return Column != null; } 
        }

        #endregion

        #region Add Operator

        /// <summary>
        /// Create AddOperator button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand AddOperatorCommand
        {
            get
            {
                if (_addOperatorCommand == null)
                {
                    Action<object> addOperatorAction = new Action<object>(this.AddOperatorCommandClick);
                    _addOperatorCommand = new RelayCommand(addOperatorAction, param => this.CanAddOperator);
                }

                return _addOperatorCommand;
            }
        }

        /// <summary>
        /// Handles event when AddOperator button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void AddOperatorCommandClick(object param)
        {
            if (string.IsNullOrEmpty(SqlWhereClause) || SqlWhereClause.EndsWith(" "))
                SqlWhereClause += ComparisonOperator;
            else
                SqlWhereClause += String.Concat(" ", ComparisonOperator);

            OnPropertyChanged("SqlWhereClause");
        }

        /// <summary>
        /// Determine if the AddOperator button can be clicked.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanAddOperator
        {
            get { return (ComparisonOperator != null); } 
        }

        #endregion

        #region Add Value

        /// <summary>
        /// Create AddValue button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand AddValueCommand
        {
            get
            {
                if (_addValueCommand == null)
                {
                    Action<object> addValueAction = new Action<object>(this.AddValueCommandClick);
                    _addValueCommand = new RelayCommand(addValueAction, param => this.CanAddValue);
                }

                return _addValueCommand;
            }
        }

        /// <summary>
        /// Handles event when AddValue button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void AddValueCommandClick(object param)
        {
            if (string.IsNullOrEmpty(SqlWhereClause) || SqlWhereClause.EndsWith(" "))
                SqlWhereClause += _db.QuoteValue(QueryValue);
            else
                SqlWhereClause += String.Concat(" ", _db.QuoteValue(QueryValue));


            OnPropertyChanged("SqlWhereClause");
        }

        /// <summary>
        /// Determine if the AddValue button can be clicked.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>        
        private bool CanAddValue
        { 
            get { return QueryValue != null; } 
        }

        #endregion

        #endregion

        #region SQL

        public string SqlFromTables
        {
            get { return _sqlFromTables; }
            set { _sqlFromTables = value; }
        }

        public string SqlWhereClause
        {
            get { return _sqlWhereClause; }
            set { _sqlWhereClause = value; }
        }

        #endregion

        #region Load/Save

        public bool LoadSQLQuery()
        {
            // Load the query dialog from file
            string filePath = Settings.Default.SqlPath;
            string fileName = null;

            if (!Directory.Exists(filePath))
                filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // Load the file contents into the dialog
            if (LoadSQLQueryFile(ref filePath, ref fileName))
            {
            }

            // If the existing file path does not exist
            // then save the new path in the settings
            if (!Directory.Exists(filePath) && filePath != null)
            {
                Settings.Default.SqlPath = filePath;
                Settings.Default.Save();
            }
            return true;
        }

        private bool LoadSQLQueryFile(ref string queryPath, ref string queryFile)
        {
            try
            {
                OpenFileDialog openFileDlg = new OpenFileDialog();
                openFileDlg.Title = "Load Dialog from Query File";
                openFileDlg.Filter = "HLUTool SQL Query (*.hsq)|*.hsq";
                openFileDlg.CheckPathExists = true;
                openFileDlg.CheckFileExists = true;
                openFileDlg.ValidateNames = true;
                openFileDlg.Multiselect = false;
                openFileDlg.RestoreDirectory = false;
                openFileDlg.InitialDirectory = queryPath;

                bool? ok = openFileDlg.ShowDialog();
                if (ok == true)
                {
                    queryFile = openFileDlg.FileName;
                    queryPath = Path.GetDirectoryName(queryFile);
                }
                else
                {
                    queryFile = null;
                    queryPath = null;
                }

                return (ok == true);
            }
            catch { return false; }
        }

        public bool SaveSQLQuery()
        {
            // Save the query dialog to file
            string filePath = Settings.Default.SqlPath;
            string fileName = null;

            if (!Directory.Exists(filePath))
                filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // Save the dialog contents to the file
            if (SaveSQLQueryFile(ref filePath, ref fileName))
            {
            }

            // If the existing file path does not exist
            // then save the new path in the settings
            if (!Directory.Exists(filePath) && filePath != null)
            {
                Settings.Default.SqlPath = filePath;
                Settings.Default.Save();
            }
            return true;
        }

        public bool SaveSQLQueryFile(ref string queryPath, ref string queryFile)
        {
            try
            {
                SaveFileDialog saveFileDlg = new SaveFileDialog();
                saveFileDlg.Title = "Save Dialog to Query File";
                saveFileDlg.Filter = "HLUTool SQL Query (*.hsq)|*.hsq";
                saveFileDlg.CheckPathExists = true;
                saveFileDlg.RestoreDirectory = false;
                saveFileDlg.InitialDirectory = queryPath;

                bool ?ok = saveFileDlg.ShowDialog();
                if (ok == true)
                {
                    if (!File.Exists(saveFileDlg.FileName))
                    {
                        queryFile = saveFileDlg.FileName;
                    }
                    else
                    {
                        queryFile = null;
                    }
                }
                else
                {
                    queryFile = null;
                }

                return (ok == true);
            }
            catch { return false; }
        }

        #endregion

        #region IDataErrorInfo Members

        public string Error
        {
            get 
            {
                StringBuilder error = new StringBuilder();

                if (error.Length > 0)
                    return error.ToString();
                else
                    return null;
            }
        }

        public string this[string columnName]
        {
            get 
            {
                string error = null;

                CommandManager.InvalidateRequerySuggested();
                
                return error;
            }
        }

        #endregion
    }
}
