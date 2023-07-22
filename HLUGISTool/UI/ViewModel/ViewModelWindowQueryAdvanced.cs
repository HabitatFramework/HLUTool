// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2014 Sussex Biodiversity Record Centre
// Copyright © 2019 Greenspace Information for Greater London CIC
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
using System.Data.Odbc;
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
    class ViewModelWindowQueryAdvanced : ViewModelBase, IDataErrorInfo
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

        private Dictionary<Type, Int32> _typeMapSystemToSQL = new Dictionary<Type, int>();

        private string _displayName = "Advanced Query Builder";
        private Cursor _cursorType = Cursors.Arrow;

        private DbBase _db;

        private Dictionary<string, DataTable> _tables;
        private string[] _comparisonOperators;
        private int? _getValueRows = Settings.Default.GetValueRows;
        private Dictionary<string, object> _queryValues;

        private static DataTable _table;
        private static DataColumn _column;
        private static string _comparisonOperator;
        private static object _queryValue;

        private static string _sqlFromTables;
        private static string _sqlWhereClause;

        private string _descriptionFieldName = Settings.Default.LutDescriptionFieldName;
        private int _descriptionFieldOrdinal = Settings.Default.LutDescriptionFieldOrdinal;
        private Regex _queryValueRegex = new Regex(@"\s+:\s+", RegexOptions.IgnoreCase); // @"\A(?<code>[^:\s]+)\s+:\s+(?<desc>[^:]+)\z", RegexOptions.IgnoreCase);

        private long _lastValueCounter = 0;

        #endregion

        #region Constructor

        public ViewModelWindowQueryAdvanced(HluDataSet hluDataset, DbBase hluDatabase)
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

        #endregion

        #region RequestClose

        // declare the delegate since using non-generic pattern
        public delegate void RequestCloseEventHandler(string sqlFromTables, string sqlWhereClause);

        // declare the event
        public event RequestCloseEventHandler RequestClose;

        #endregion

        #region GetValues Command

        /// <summary>
        /// Set GetValues button command.
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
        /// Handles events when the GetValues button is clicked.
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
                ChangeCursor(Cursors.Wait);

                // Create a data reader to retrieve the rows for
                // the required column.
                IDataReader dataReader = null;

                try
                {
                    // Load the data reader to retrieve the rows for
                    // the required column.
                    dataReader = _db.ExecuteReader(String.Format(
                        "SELECT DISTINCT {0} FROM {1} WHERE {0} IS NOT NULL ORDER BY {0}",
                        _db.QuoteIdentifier(Column.ColumnName),
                        _db.QualifyTableName(Table.TableName)),
                        _db.Connection.ConnectionTimeout, CommandType.Text);

                    if (dataReader == null) throw new Exception(String.Format("Error reading values from {0}.{1}", Table.TableName, Column.ColumnName));

                    // Define a new dictionary to hold the column values.
                    Dictionary<string, object> q = new Dictionary<string, object>();

                    // Load the dictionary with the first/next n values
                    // (based on the max value in the user settings.
                    long i = 0;
                    while (i < (_lastValueCounter + _getValueRows) && dataReader.Read())
                    {
                        if (i >= _lastValueCounter)
                        {
                            // Add the key (string) and value (object) to the dictionary
                            // (if the key already exists the value will be updated).
                            if ((Table.TableName.ToLower() == "incid_osmm_updates") &&
                               (Column.ColumnName.ToLower() == "status"))
                            {
                                String status;
                                switch (dataReader.GetValue(0).ToString())
                                {
                                case "-99":
                                    status = "Rejected";
                                    break;
                                case "-2":
                                    status = "Ignored";
                                    break;
                                case "-1":
                                    status = "Applied";
                                    break;
                                case "0":
                                    status = "Pending";
                                    break;
                                default:
                                    status = "Proposed";
                                    break;
                                }
                                q[dataReader.GetValue(0).ToString() + " : " + status] = dataReader.GetValue(0);
                            }
                            else
                                q[dataReader.GetValue(0).ToString()] = dataReader.GetValue(0);
                        }
                        i += 1;
                    }

                    // If the last record has been reached.
                    if (i != (_lastValueCounter + _getValueRows))
                        // Set the last value counter to flag this.
                        _lastValueCounter = -1;
                    else
                        // Move the last value counter on.
                        _lastValueCounter = i;

                    // Set the combobox of values to the new dictionary.
                    _queryValues = q;

                    // Reset the cursor back to normal.
                    ChangeCursor(Cursors.Arrow);
                }
                catch (Exception ex)
                {
                    // Reset the cursor back to normal.
                    ChangeCursor(Cursors.Arrow);

                    MessageBox.Show(ex.Message, "HLU Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    // Close the data reader.
                    if (!dataReader.IsClosed)
                        dataReader.Close();
                }
            }
            else
            {
                _queryValues = null;
                _lastValueCounter = 0;
            }

            OnPropertyChanged("QueryValues");

        }

        /// <summary>
        /// Determine if the GetValues button can be clicked.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool CanGetValues
        {
            get
            {
                // Return false by default (e.g. if
                // there is no table or column selected).
                bool canGet = false;

                if ((Table != null) && (Column != null))
                {
                    // Find the related lookup tables for the selected table and column
                    IEnumerable<DataRelation> parentRelations = Table.ParentRelations.Cast<DataRelation>();
                    IEnumerable<DataRelation> lutRelations = parentRelations.Where(r => r.ChildTable == Table &&
                        r.ParentTable.TableName.StartsWith("lut_", StringComparison.CurrentCultureIgnoreCase) &&
                        r.ChildColumns.Length == 1 && r.ChildColumns.Contains(Column));

                    // If there are no related lookup tables (or there are
                    // more than one relations for the current table and
                    // column) the dropdown list will load automatically.
                    if (lutRelations == null || lutRelations.Count() != 1)
                        // Return true if the last record has not been reached.
                        if (_lastValueCounter != -1)
                            canGet = true;
                }

                // Update the button tool tip.
                OnPropertyChanged("GetValuesToolTip");

                return canGet;
            }
        }

        /// <summary>
        /// Gets or sets the number of rows to get.
        /// </summary>
        /// <value>
        /// The number of rows to get.
        /// </value>
        public int? GetValueRows
        {
            get { return (_getValueRows != null && _getValueRows <= 0) ? 1000 : _getValueRows; }
            set { _getValueRows = value; }
        }

        /// <summary>
        /// Gets the tool tip to display for the GetValues button.
        /// </summary>
        /// <value>
        /// The GetValues button tool tip.
        /// </value>
        public string GetValuesToolTip
        {
            get
            {
                if (Column != null)
                {
                    if (_lastValueCounter == 0)
                        return String.Format("Get first {0} values for column '{1}'", GetValueRows, Column.ColumnName);
                    else
                        return String.Format("Get next {0} values for column '{1}'", GetValueRows, Column.ColumnName);
                }
                else
                    return null;
            }
        }

        #endregion

        #region Clear Command

        /// <summary>
        /// Set the Clear button command.
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
        /// Handles events when the Clear button is clicked.
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
        /// Set the Verify button command.
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
        /// Handles events when the Verify button is clicked.
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void VerifyCommandClick(object param)
        {
            if ((SqlFromTables != null) && (SqlWhereClause != null))
            {
                try
                {
                    // Show the wait cursor whilst verifying the sql.
                    ChangeCursor(Cursors.Wait);

                    // Get a list of all the possible query tables.
                    List<DataTable> tables = new List<DataTable>();
                    if ((ViewModelWindowQueryAdvanced.HluDatasetStatic != null))
                    {
                        tables = ViewModelWindowQueryAdvanced.HluDatasetStatic.incid.ChildRelations
                            .Cast<DataRelation>().Select(r => r.ChildTable).ToList();
                        tables.Add(ViewModelWindowQueryAdvanced.HluDatasetStatic.incid);
                    }

                    // Split the string of query table names created by the
                    // user in the form into an array.
                    string[] fromTables = SqlFromTables.Split(',').Select(s => s.Trim(' ')).Distinct().ToArray();

                    // Select only the database tables that are in the query array.
                    List<DataTable> whereTables = tables.Where(t => fromTables.Contains(t.TableName)).ToList();

                    // Parse the SQL to see if it is valid.
                    if (whereTables.Count() > 0)
                    {
                        // Replace any connection type specific qualifiers and delimiters.
                        string newWhereClause = null;
                        if (SqlWhereClause != null)
                            newWhereClause = ReplaceStringQualifiers(SqlWhereClause);

                        // Validate the SQL by trying to select the top 1 row.
                        string validity = _db.SqlValidate(HluDatasetStatic.incid.PrimaryKey, whereTables, newWhereClause);

                        // Reset the cursor back to normal.
                        ChangeCursor(Cursors.Arrow);

                        // The SQL is valid.
                        if (validity == "1")
                        {
                            // Warn the user that the SQL is invalid.
                            MessageBox.Show(App.GetActiveWindow(), "SQL is valid.", "HLU Query",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        // The SQL is valid but did not return any rows.
                        else if (validity == "0")
                        {
                            // Warn the user that no rows were returned.
                            MessageBox.Show(App.GetActiveWindow(), "SQL is valid but no records were returned.", "HLU Query",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                        // The SQL is not valid.
                        else if (validity != null)
                        {
                            // Warn the user that the SQL is invalid.
                            MessageBox.Show(App.GetActiveWindow(), String.Format("Sql is invalid.\n\n{0}", validity), "HLU Query",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    else
                    {
                        // Reset the cursor back to normal.
                        ChangeCursor(Cursors.Arrow);

                        // Warn the user that the no valid tables were found.
                        MessageBox.Show(App.GetActiveWindow(), "No valid tables were found.", "HLU Query",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    // Reset the cursor back to normal.
                    ChangeCursor(Cursors.Arrow);

                    MessageBox.Show(App.GetActiveWindow(), ex.Message, "HLU Query",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally { }
            }
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
        /// Set the Ok button command.
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
        /// Handles events when the Ok button is clicked.
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
        private bool CanOk { get { return (String.IsNullOrEmpty(Error) &&
            !String.IsNullOrEmpty(SqlFromTables) && !String.IsNullOrEmpty(SqlWhereClause)); } }

        #endregion

        #region Cancel Command

        /// <summary>
        /// Set the Cancel button command.
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
        /// Handles events when the Cancel button is clicked.
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
        /// Set the Load button command.
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
        /// Handles events when the Load button is clicked.
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
        /// Set the Save button command.
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
                    _saveCommand = new RelayCommand(saveAction, param => this.CanSave);
                }

                return _saveCommand;
            }
        }

        /// <summary>
        /// Handles events when the Save button is clicked.
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void SaveCommandClick(object param)
        {
            SaveSQLQuery();
        }

        /// <summary>
        /// Determine if the Save button can be clicked.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CanSave { get { return (!String.IsNullOrEmpty(SqlFromTables) && !String.IsNullOrEmpty(SqlWhereClause)); } }

        #endregion

        #region Table

        /// <summary>
        /// Gets or sets the list of valid data tables that can be queried.
        /// </summary>
        /// <value>
        /// The list of valid tables.
        /// </value>
        public Dictionary<string, DataTable> Tables
        {
            get
            {
                if ((_tables == null) && ((HluDatasetStatic != null)))
                {
                    _tables = HluDatasetStatic.incid.ChildRelations
                        .Cast<DataRelation>().OrderBy(r => r.ChildTable.TableName).ToDictionary(r => r.ChildTable.TableName, r => r.ChildTable);
                    _tables.Add("incid", HluDatasetStatic.incid);
                }
                //return _tables;
                return _tables.OrderBy(x => x.Key).ToDictionary(x => x.Key, x => x.Value);
            }
            set { }
        }

        /// <summary>
        /// Gets or sets the current data table.
        /// </summary>
        /// <value>
        /// The current data table.
        /// </value>
        public DataTable Table
        {
            get { return _table; }
            set
            {
                _table = value;
                _column = null;
                OnPropertyChanged("Table");
                OnPropertyChanged("Columns");
                OnPropertyChanged("ColumnIsEnabled");

                _queryValues = null;
                _lastValueCounter = 0;
                OnPropertyChanged("QueryValues");
                //OnPropertyChanged("QueryValueIsEnabled");
            }
        }

        #endregion

        #region Column

        /// <summary>
        /// Gets or sets the list of valid data columns that can be queried.
        /// </summary>
        /// <value>
        /// The list of valid columns for the current table.
        /// </value>
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

        /// <summary>
        /// Gets or sets the current data column.
        /// </summary>
        /// <value>
        /// The current data column.
        /// </value>
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
                    _lastValueCounter = 0;
                    OnPropertyChanged("QueryValues");
                    //OnPropertyChanged("QueryValueIsEnabled");
                    OnPropertyChanged("CanGetValues");
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether the column combobox is enabled.
        /// </summary>
        /// <value>
        ///   <c>true</c> if the column combobox is enabled; otherwise, <c>false</c>.
        /// </value>
        public bool ColumnIsEnabled
        {
            get { return Table != null; }
        }

        #endregion

        #region Operator

        /// <summary>
        /// Gets the list of valid comparison operators.
        /// </summary>
        /// <value>
        /// The list of valid comparison operators.
        /// </value>
        public string[] ComparisonOperators
        {
            get 
            {
                if (_comparisonOperators == null)
                {
                    _comparisonOperators = new string[] { "=", "<", ">", "<=", ">=", "<>", "(", ")", "AND", "OR",
                        "IS NULL", "IS NOT NULL", "LIKE", "NOT LIKE", "IN ()", "NOT IN ()" };
                }
                return _comparisonOperators;
            }
        }

        /// <summary>
        /// Gets or sets the current comparison operator.
        /// </summary>
        /// <value>
        /// The current comparison operator.
        /// </value>
        public string ComparisonOperator
        {
            get { return _comparisonOperator; }
            set
            {
                _comparisonOperator = value;
                OnPropertyChanged("ComparisonOperator");
            }
        }

        /// <summary>
        /// Gets a value indicating whether the comparison operator combobox is enabled.
        /// </summary>
        /// <value>
        /// <c>true</c> if the comparison operator combobox is enabled; otherwise, <c>false</c>.
        /// </value>
        public bool ComparisonOperatorIsEnabled
        {
            get { return Column != null; }
        }

        #endregion

        #region Value

        /// <summary>
        /// Gets or sets the list of unique data values relating to the
        /// current data table and column.
        /// </summary>
        /// <value>
        /// The list of unique data values for the current table and column.
        /// </value>
        public Dictionary<string, object> QueryValues
        {
            get
            {
                if (_queryValues == null)
                {
                    // Return null if there is no table or column selected.
                    if ((Table == null) || (Column == null))
                    {
                        _lastValueCounter = 0;
                        return _queryValues;
                    }

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
                        // Show the wait cursor whilst loading the values.
                        ChangeCursor(Cursors.Wait);

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

                        //---------------------------------------------------------------------
                        // CHANGED: CR52 Enable support for multiple priority habitat classifications
                        // Enable multiple priority habitat types (from the same or different
                        // classifications) to be assigned
                        //
                        //// If the select table is the 'incid_bap' table then only
                        //// 'PHAP' values from the related lookup table (lut_habitat_type).
                        //var q = lut is HluDataSet.lut_habitat_typeDataTable && Table is HluDataSet.incid_bapDataTable ?
                        //    lut.AsEnumerable().Where(r => Regex.IsMatch(r[lut.PrimaryKey[0].Ordinal].ToString(), @"\APHAP")) :
                        //    lut.AsEnumerable();

                        // If the select table is the 'incid_bap' table then only
                        // bap_priority values from the related lookup table (lut_habitat_type).
                        var q = lut.AsEnumerable();

                        if (lut is HluDataSet.lut_habitat_typeDataTable && Table is HluDataSet.incid_bapDataTable)
                        {
                            q = lut.AsEnumerable().Where(r => r[lut.Columns["bap_priority"]].Equals(true));
                            _queryValues = q.ToDictionary(r => r[lutColumn].ToString() + (descriptionColumn != lutColumn ?
                                " : " + r[descriptionColumn].ToString() : String.Empty), r => r[lutColumn]);
                        }
                        else if (lut is HluDataSet.lut_osmm_habitat_xrefDataTable && Table is HluDataSet.incid_osmm_updatesDataTable)
                        {
                            DataColumn summaryColumn = lut.Columns["ihs_summary"];
                            _queryValues = q.ToDictionary(r => r[lutColumn].ToString() + " : " + r[summaryColumn].ToString(), r => r[lutColumn]);

                            //Dictionary<string, object> qvalues = new Dictionary<string, object>();

                            //foreach (var qv in q.OrderBy(key => key[summaryColumn])) {
                            //    qvalues.Add(qv[lutColumn].ToString() + " : " + qv[summaryColumn].ToString(), qv[lutColumn]);
                            //}
                            //_queryValues = qvalues;
                        }
                        else
                        {
                            _queryValues = q.ToDictionary(r => r[lutColumn].ToString() + (descriptionColumn != lutColumn ?
                                " : " + r[descriptionColumn].ToString() : String.Empty), r => r[lutColumn]);
                        }
                        //---------------------------------------------------------------------

                        // Reset the cursor back to normal.
                        ChangeCursor(Cursors.Arrow);
                    }
                    else
                    {
                        _queryValues = null;
                    }
                }

                OnPropertyChanged("QueryValueIsEnabled");
                return _queryValues;
            }
            set
            {
                _queryValues = value;
            }
        }

        /// <summary>
        /// Gets or sets the current data value.
        /// </summary>
        /// <value>
        /// The current data value.
        /// </value>
        public object QueryValue
        {
            get { return _queryValue != null ? _queryValue : null; }
            set
            {
                _queryValue = value;
                OnPropertyChanged("QueryValue");
            }
        }

        /// <summary>
        /// Gets a value indicating whether the QueryValue combobox is enabled.
        /// </summary>
        /// <value>
        /// <c>true</c> if the QueryValue combobox is enabled; otherwise, <c>false</c>.
        /// </value>
        public bool QueryValueIsEnabled
        {
            get { return _queryValues != null; }
        }

        #endregion

        #region Add Buttons

        #region Add Table

        /// <summary>
        /// Set the AddTable button command.
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
        /// Handles events when the AddTable button is clicked.
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void AddTableCommandClick(object param)
        {
            // Only add the table if it's not already in the list
            if (SqlFromTables == null || SqlFromTables.Split(',').Contains(Table.TableName) == false)
            {
                if (string.IsNullOrEmpty(SqlFromTables) || SqlFromTables.TrimEnd(' ').EndsWith(","))
                    SqlFromTables += Table.TableName;
                else
                    SqlFromTables += String.Concat(", ", Table.TableName);

                OnPropertyChanged("SqlFromTables");
            }
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
        /// Set the AddColumn button command.
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
        /// Handles events when the AddColumn button is clicked.
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void AddColumnCommandClick(object param)
        {
            if (string.IsNullOrEmpty(SqlWhereClause) || SqlWhereClause.EndsWith(" "))
                SqlWhereClause += QuoteIdentifier(Table.TableName) + "." + Column.ColumnName;
            else
                SqlWhereClause += String.Concat(" ", QuoteIdentifier(Table.TableName), ".", Column.ColumnName);

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
        /// Set the AddOperator button command.
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
        /// Handles events when the AddOperator button is clicked.
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
        /// Set the AddValue button command.
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
        /// Handles events when the AddValue button is clicked.
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void AddValueCommandClick(object param)
        {
            if (string.IsNullOrEmpty(SqlWhereClause) || SqlWhereClause.EndsWith(" "))
                SqlWhereClause += QuoteValue(QueryValue);
            else
                SqlWhereClause += String.Concat(" ", QuoteValue(QueryValue));


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

        /// <summary>
        /// Gets or sets the string of data tables to use in the
        /// FROM clause of the SQL.
        /// </summary>
        /// <value>
        /// The string of data tables to use in the FROM clause of the SQL.
        /// </value>
        public string SqlFromTables
        {
            get { return _sqlFromTables; }
            set { _sqlFromTables = value; }
        }

        /// <summary>
        /// Gets or sets the string of conditions to use in the
        /// WHERE clause of the SQL.
        /// </summary>
        /// <value>
        /// The string of conditions to use in the WHERE clause of the SQL.
        /// </value>
        public string SqlWhereClause
        {
            get { return _sqlWhereClause; }
            set { _sqlWhereClause = value; }
        }

        #endregion

        #region Load/Save

        /// <summary>
        /// Loads an existing SQL query.
        /// </summary>
        /// <exception cref="System.Exception"></exception>
        public void LoadSQLQuery()
        {
            // Load the query dialog from file
            string filePath = Settings.Default.SqlPath;
            string fileName = null;

            try
            {
                if (!Directory.Exists(filePath))
                    filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                // Load the file contents into the dialog
                if (LoadSQLQueryFile(ref filePath, ref fileName))
                {
                    if (!File.Exists(fileName))
                    {
                        throw new Exception(String.Format("File {0} was not found!", fileName));
                    }

                    // Read all the lines in the script into an array.
                    string[] lines = File.ReadAllLines(fileName);

                    // Process each line in the script.
                    foreach (string line in lines)
                    {
                        // Remove any leading or trailing spaces from the line.
                        string sqlCmd = line.Trim();

                        // Skip the line if it is empty.
                        if ((sqlCmd.Length == 0) || (string.IsNullOrEmpty(sqlCmd)))
                            continue;

                        // Break the line command into words.
                        string[] words = sqlCmd.Split(' ');

                        // If there are not enough words then skip to the next line.
                        if (words.Length < 1)
                            continue;

                        string removeString;
                        int findIndex;
                        // Split the tables and where lines into parts.
                        switch (words[0].ToLower())
                        {
                            case "tables":
                                removeString = "Tables";
                                findIndex = sqlCmd.IndexOf(removeString);
                                string fromTables = (findIndex < 0)
                                    ? sqlCmd
                                    : sqlCmd.Remove(findIndex, removeString.Length);

                                SqlFromTables = fromTables.Trim().TrimStart('{').TrimEnd('}').Trim();
                                break;
                            case "where":
                                removeString = "Where";
                                findIndex = sqlCmd.IndexOf(removeString);
                                string whereClause = (findIndex < 0)
                                    ? sqlCmd
                                    : sqlCmd.Remove(findIndex, removeString.Length);

                                SqlWhereClause = whereClause.Trim().TrimStart('{').TrimEnd('}').Trim();
                                break;
                            default:
                                break;
                        }
                    }
                    OnPropertyChanged("SqlFromTables");
                    OnPropertyChanged("SqlWhereClause");
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(App.GetActiveWindow(), ex.Message, "HLU Query",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally 
            {

                // If the existing file path does not exist
                // then save the new path in the settings
                if (!Directory.Exists(Settings.Default.SqlPath) && filePath != null)
                {
                    Settings.Default.SqlPath = filePath;
                    Settings.Default.Save();
                }
            }
        }

        /// <summary>
        /// Prompts the user to select an existing SQL query file to load.
        /// </summary>
        /// <param name="queryPath">The path of the SQL query to load.</param>
        /// <param name="queryFile">The file name of the SQL query to load.</param>
        /// <returns></returns>
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

        /// <summary>
        /// Loads an existing SQL query.
        /// </summary>
        /// <exception cref="System.Exception"></exception>
        public void SaveSQLQuery()
        {
            // Save the query dialog to file
            string filePath = Settings.Default.SqlPath;
            string fileName = null;

            try
            {
                if (!Directory.Exists(filePath))
                    filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                // Save the dialog contents to the file
                if (SaveSQLQueryFile(ref filePath, ref fileName))
                {
                    // Create an array of lines.
                    string[] lines = new string[2];

                    // Store the dialog values to the array.
                    lines[0] = String.Format("Tables {0}{1}{2}", "{", SqlFromTables, "}");
                    lines[1] = String.Format("Where {0}{1}{2}", "{", SqlWhereClause, "}");

                    // Write all the lines in the array to the script file.
                    File.WriteAllLines(fileName, lines);

                    if (!File.Exists(fileName))
                    {
                        throw new Exception(String.Format("File {0} was not created!", fileName));
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(App.GetActiveWindow(), ex.Message, "HLU Query",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {

                // If the existing file path does not exist
                // then save the new path in the settings
                if (!Directory.Exists(Settings.Default.SqlPath) && filePath != null)
                {
                    Settings.Default.SqlPath = filePath;
                    Settings.Default.Save();
                }
            }
        }

        /// <summary>
        /// Prompts the user to enter a file name to save the SQL query to.
        /// </summary>
        /// <param name="queryPath">The path to save the SQL query to.</param>
        /// <param name="queryFile">The file name to save the SQL query to.</param>
        /// <returns></returns>
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
                    queryFile = saveFileDlg.FileName;
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

        #endregion

        #region SQLUpdater

        /// <summary>
        /// Replaces any string or date delimiters with connection type specific
        /// versions and qualifies any table names.
        /// </summary>
        /// <param name="words">The words.</param>
        /// <returns></returns>
        internal String ReplaceStringQualifiers(String sqlcmd)
        {
            // Check if a table name (delimited by '[]' characters) is found
            // in the sql command.
            int i1 = 0;
            int i2 = 0;
            String start = String.Empty;
            String end = String.Empty;

            while ((i1 != -1) && (i2 != -1))
            {
                i1 = sqlcmd.IndexOf("[", i2);
                if (i1 != -1)
                {
                    i2 = sqlcmd.IndexOf("]", i1 + 1);
                    if (i2 != -1)
                    {
                        // Strip out the table name.
                        string table = sqlcmd.Substring(i1 + 1, i2 - i1 - 1);

                        // Split the table name from the rest of the sql command.
                        if (i1 == 0)
                            start = String.Empty;
                        else
                            start = sqlcmd.Substring(0, i1);

                        if (i2 == sqlcmd.Length - 1)
                            end = String.Empty;
                        else
                            end = sqlcmd.Substring(i2 + 1);

                        // Replace the table name with a qualified table name.
                        sqlcmd = start + _db.QualifyTableName(table) + end;

                        // Reposition the last index.
                        i2 = sqlcmd.Length - end.Length;
                    }
                }
            }

            // Check if any strings are found (delimited by single quotes)
            // in the sql command.
            i1 = 0;
            i2 = 0;

            while ((i1 != -1) && (i2 != -1))
            {
                i1 = sqlcmd.IndexOf("'", i2);
                if (i1 != -1)
                {
                    i2 = sqlcmd.IndexOf("'", i1 + 1);
                    if (i2 != -1)
                    {
                        // Strip out the text string.
                        string text = sqlcmd.Substring(i1 + 1, i2 - i1 - 1);

                        // Split the text string from the rest of the sql command.
                        if (i1 == 0)
                            start = String.Empty;
                        else
                            start = sqlcmd.Substring(0, i1);

                        if (i2 == sqlcmd.Length - 1)
                            end = String.Empty;
                        else
                            end = sqlcmd.Substring(i2 + 1);

                        // Replace any wild characters found in the text.
                        if (start.TrimEnd().EndsWith(" LIKE"))
                        {
                            text.Replace("_", _db.WildcardSingleMatch);
                            text.Replace("%", _db.WildcardManyMatch);
                        }

                        // Replace the text delimiters with the correct delimiters.
                        sqlcmd = start + _db.QuoteValue(text) + end;

                        // Reposition the last index.
                        i2 = sqlcmd.Length - end.Length;
                    }
                }
            }

            // Check if any dates are found (delimited by '#' characters)
            // in the sql command.
            i1 = 0;
            i2 = 0;

            while ((i1 != -1) && (i2 != -1))
            {
                i1 = sqlcmd.IndexOf("#", i2);
                if (i1 != -1)
                {
                    i2 = sqlcmd.IndexOf("#", i1 + 1);
                    if (i2 != -1)
                    {
                        // Strip out the date string.
                        DateTime dt;
                        DateTime.TryParse(sqlcmd.Substring(i1 + 1, i2 - i1 - 1), out dt);

                        // Split the date string from the rest of the sql command.
                        if (i1 == 0)
                            start = String.Empty;
                        else
                            start = sqlcmd.Substring(0, i1);

                        if (i2 == sqlcmd.Length - 1)
                            end = String.Empty;
                        else
                            end = sqlcmd.Substring(i2 + 1);

                        // Replace the date delimiters with the correct delimiters.
                        sqlcmd = start + _db.QuoteValue(dt) + end;

                        // Reposition the last index.
                        i2 = sqlcmd.Length - end.Length;
                    }
                }
            }
            return sqlcmd;
        }

        #endregion

        #region Quotes & Qualifiers

        /// <summary>
        /// Add quote characters to the specified string (if not
        /// already present).
        /// </summary>
        /// <param name="identifier">The string to add quotes to.</param>
        /// <returns></returns>
        public string QuoteIdentifier(string identifier)
        {
            if (!String.IsNullOrEmpty(identifier))
            {
                if (!identifier.StartsWith(QuotePrefix)) identifier = identifier.Insert(0, QuotePrefix);
                if (!identifier.EndsWith(QuoteSuffix)) identifier += QuoteSuffix;
            }
            return identifier;
        }

        public string QuotePrefix { get { return "["; } }

        public string QuoteSuffix { get { return "]"; } }

        public string StringLiteralDelimiter { get { return "\'"; } }

        public string DateLiteralPrefix { get { return "#"; } }

        public string DateLiteralSuffix { get { return "#"; } }

        public string WildcardSingleMatch { get { return "_"; } }

        public string WildcardManyMatch { get { return "%"; } }

        /// <summary>
        /// Does not escape string delimiter or other special characters.
        /// Does check if value is already quoted.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public string QuoteValue(object value)
        {
            if (value == null) return "NULL";
            Type valueType = value.GetType();
            int colType;

            if (_typeMapSystemToSQL == null || _typeMapSystemToSQL.Count() == 0)
            {
                Dictionary<Type, int> typeMapSystemToSQLAdd = new Dictionary<Type, int>();
                typeMapSystemToSQLAdd.Add(typeof(Object), (int)OdbcType.VarBinary);
                typeMapSystemToSQLAdd.Add(typeof(Boolean), (int)OdbcType.Bit);
                typeMapSystemToSQLAdd.Add(typeof(SByte), (int)OdbcType.Int);
                typeMapSystemToSQLAdd.Add(typeof(Byte), (int)OdbcType.TinyInt);
                typeMapSystemToSQLAdd.Add(typeof(Int16), (int)OdbcType.SmallInt);
                typeMapSystemToSQLAdd.Add(typeof(UInt16), (int)OdbcType.SmallInt);
                typeMapSystemToSQLAdd.Add(typeof(Int32), (int)OdbcType.Int);
                typeMapSystemToSQLAdd.Add(typeof(UInt32), (int)OdbcType.Int);
                typeMapSystemToSQLAdd.Add(typeof(Int64), (int)OdbcType.BigInt);
                typeMapSystemToSQLAdd.Add(typeof(UInt64), (int)OdbcType.BigInt);
                typeMapSystemToSQLAdd.Add(typeof(Single), (int)OdbcType.Real);
                typeMapSystemToSQLAdd.Add(typeof(Double), (int)OdbcType.Double);
                typeMapSystemToSQLAdd.Add(typeof(Decimal), (int)OdbcType.Decimal);
                typeMapSystemToSQLAdd.Add(typeof(DateTime), (int)OdbcType.DateTime);
                typeMapSystemToSQLAdd.Add(typeof(TimeSpan), (int)OdbcType.DateTime);
                typeMapSystemToSQLAdd.Add(typeof(Byte[]), (int)OdbcType.VarBinary);
                typeMapSystemToSQLAdd.Add(typeof(Guid), (int)OdbcType.UniqueIdentifier);
                if (Settings.Default.DbIsUnicode)
                {
                    typeMapSystemToSQLAdd.Add(typeof(Char), (int)OdbcType.NChar);
                    typeMapSystemToSQLAdd.Add(typeof(String), (int)OdbcType.NText);
                    typeMapSystemToSQLAdd.Add(typeof(Char[]), (int)OdbcType.NText);
                }
                else
                {
                    typeMapSystemToSQLAdd.Add(typeof(Char), (int)OdbcType.Char);
                    typeMapSystemToSQLAdd.Add(typeof(String), (int)OdbcType.Text);
                    typeMapSystemToSQLAdd.Add(typeof(Char[]), (int)OdbcType.Text);
                }

                foreach (KeyValuePair<Type, int> kv in typeMapSystemToSQLAdd)
                {
                    if (!_typeMapSystemToSQL.ContainsKey(kv.Key))
                        _typeMapSystemToSQL.Add(kv.Key, kv.Value);
                }

                if (Settings.Default.DbIsUnicode)
                {
                    ReplaceType(typeof(Char), (int)OdbcType.NChar, _typeMapSystemToSQL);
                    ReplaceType(typeof(String), (int)OdbcType.NText, _typeMapSystemToSQL);
                    ReplaceType(typeof(Char[]), (int)OdbcType.NText, _typeMapSystemToSQL);
                }
                else
                {
                    ReplaceType(typeof(Char), (int)OdbcType.Char, _typeMapSystemToSQL);
                    ReplaceType(typeof(String), (int)OdbcType.Text, _typeMapSystemToSQL);
                    ReplaceType(typeof(Char[]), (int)OdbcType.Text, _typeMapSystemToSQL);
                }
            }

            if (_typeMapSystemToSQL.TryGetValue(valueType, out colType))
            {
                string s = valueType == typeof(DateTime) ? ((DateTime)value).ToString("s").Replace("T", " ") : value.ToString();
                switch ((OdbcType)colType)
                {
                    case OdbcType.Char:
                    case OdbcType.NChar:
                    case OdbcType.NText:
                    case OdbcType.NVarChar:
                    case OdbcType.Text:
                    case OdbcType.VarChar:
                        if (s.Length == 0) return StringLiteralDelimiter + StringLiteralDelimiter;
                        if (!s.StartsWith(StringLiteralDelimiter)) s = StringLiteralDelimiter + s;
                        if (!s.EndsWith(StringLiteralDelimiter)) s += StringLiteralDelimiter;
                        return s;
                    case OdbcType.Date:
                    case OdbcType.DateTime:
                    case OdbcType.SmallDateTime:
                    case OdbcType.Time:
                    case OdbcType.Timestamp:
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

        /// <summary>
        /// Replaces the specified database data type.
        /// </summary>
        /// <param name="sysType">Type of the system.</param>
        /// <param name="dbTypeNew">The database type new.</param>
        /// <param name="typeDictionary">The type dictionary to replace the values in.</param>
        protected void ReplaceType(Type sysType, int dbTypeNew, Dictionary<Type, int> typeDictionary)
        {
            int dbTypeOld;
            if (typeDictionary.TryGetValue(sysType, out dbTypeOld) && (dbTypeOld != dbTypeNew))
            {
                typeDictionary.Remove(sysType);
                typeDictionary.Add(sysType, dbTypeNew);
            }
        }

        #endregion

        #region Cursor

        /// <summary>
        /// Gets the cursor type to use when the cursor is over the window.
        /// </summary>
        /// <value>
        /// The window cursor type.
        /// </value>
        public Cursor WindowCursor { get { return _cursorType; } }

        public void ChangeCursor(Cursor cursorType)
        {
            _cursorType = cursorType;
            OnPropertyChanged("WindowCursor");
            if (cursorType == Cursors.Wait)
                DispatcherHelper.DoEvents();
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
