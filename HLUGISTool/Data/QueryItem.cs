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
using System.Text.RegularExpressions;
using System.Windows.Input;
using HLU.Data.Model;
using HLU.Properties;
using HLU.UI.ViewModel;

namespace HLU.Data
{
    public class QueryItem : ViewModelBase, IDataErrorInfo
    {
        #region Fields

        private string _displayName;
        private SqlFilterCondition _sqlCond;
        private bool _isComplete;
        private DataTable _table;
        private DataColumn _column;
        private string _queryValueText;
        private Dictionary<string, DataTable> _tables;
        private string[] _booleanOperators;
        private string[] _comparisonOperators;
        private Cursor _cursorType = Cursors.Arrow;
        private string _descriptionFieldName = Settings.Default.LutDescriptionFieldName;
        private int _descriptionFieldOrdinal = Settings.Default.LutDescriptionFieldOrdinal;
        private Regex _queryValueRegex = new Regex(@"\s+:\s+", RegexOptions.IgnoreCase); // @"\A(?<code>[^:\s]+)\s+:\s+(?<desc>[^:]+)\z", RegexOptions.IgnoreCase);
        
        #endregion

        #region Constructor

        public QueryItem()
        {
            _sqlCond = new SqlFilterCondition();
            _sqlCond.OpenParentheses = "(";
            _sqlCond.CloseParentheses = ")";
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
            get { return DisplayName; }
        }

        #endregion

        #region Properties

        #region Dropdown Lists

        public string[] BooleanOperators
        {
            get
            {
                if (_booleanOperators == null)
                    _booleanOperators = new string[] { "AND", "AND NOT", "OR", "OR NOT" };
                return _booleanOperators;
            }
        }

        public Dictionary<string, DataTable> Tables
        {
            get
            {
                if ((_tables == null) && ((ViewModelQueryBuilder.HluDatasetStatic != null) &&
                    (ViewModelQueryBuilder.HluDatasetStatic.incid != null)))
                {
                    _tables = ViewModelQueryBuilder.HluDatasetStatic.incid.ChildRelations
                        .Cast<DataRelation>().ToDictionary(r => r.ChildTable.TableName, r => r.ChildTable);
                    _tables.Add("incid", ViewModelQueryBuilder.HluDatasetStatic.incid);
                }
                return _tables;
            }
            set { }
        }

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

        public string[] ComparisonOperators
        {
            get 
            {
                if (_comparisonOperators == null)
                {
                    _comparisonOperators = new string[] { "=", "<", ">", "<=", ">=", "<>", 
                        "contains", "begins with", "ends with", "IS NULL", "IS NOT NULL", "LIKE", "IN ()", "NOT IN ()" };
                }
                return _comparisonOperators;
            }
        }

        public Dictionary<string, object> Values
        {
            get
            {
                if ((Table == null) || (Column == null) || (!String.IsNullOrEmpty(_sqlCond.Operator) &&
                    _sqlCond.Operator.ToUpper().EndsWith("NULL"))) return null;

                _cursorType = Cursors.Wait;
                OnPropertyChanged("CursorType");

                IEnumerable<DataRelation> parentRelations = Table.ParentRelations.Cast<DataRelation>();
                IEnumerable<DataRelation> lutRelations = parentRelations.Where(r => r.ChildTable == Table &&
                    r.ParentTable.TableName.StartsWith("lut_", StringComparison.CurrentCultureIgnoreCase) && 
                    r.ChildColumns.Length == 1 && r.ChildColumns.Contains(Column));
                DataRelation lutRelation;
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

                    var q = lut is HluDataSet.lut_habitat_typeDataTable && Table is HluDataSet.incid_bapDataTable ?
                        lut.AsEnumerable().Where(r => Regex.IsMatch(r[lut.PrimaryKey[0].Ordinal].ToString(), @"\APHAP")) :
                        lut.AsEnumerable();

                    return q.ToDictionary(r => r[lutColumn].ToString() + (descriptionColumn != lutColumn ? 
                        " : " + r[descriptionColumn].ToString() : String.Empty), r => r[lutColumn]);
                }
                return null;
            }
        }
        
        #endregion

        public Cursor CursorType
        {
            get { return _cursorType; }
        }

        #region SQLFilterCondition Values

        public SqlFilterCondition WhereCondition { get { return _sqlCond; } }

        public bool IsComplete { get { return _isComplete; } }

        public string BooleanOperator
        {
            get { return _sqlCond.BooleanOperator; }
            set { _sqlCond.BooleanOperator = value; }
        }

        public string OpenParentheses
        {
            get { return _sqlCond.OpenParentheses; }
            set { _sqlCond.OpenParentheses = value; }
        }

        public DataTable Table
        {
            get { return _table; }
            set
            {
                _sqlCond.Table = value;
                if ((_sqlCond.Table != null) && (ViewModelQueryBuilder.HluDatasetStatic != null) && 
                    ViewModelQueryBuilder.HluDatasetStatic.Tables.Contains(_sqlCond.Table.TableName))
                {
                    _table = _sqlCond.Table;
                    OnPropertyChanged("Table");
                    OnPropertyChanged("Columns");
                    OnPropertyChanged("ColumnIsEnabled");
                }
            }
        }

        public DataColumn Column
        {
            get { return _column; }
            set
            { 
                _sqlCond.Column = value;
                if ((Table != null) && (_sqlCond.Column != null) && 
                    Table.Columns.Contains(_sqlCond.Column.ColumnName))
                {
                    _column = _sqlCond.Column;
                    OnPropertyChanged("Column");
                    OnPropertyChanged("Values");
                    OnPropertyChanged("ComparisonOperatorIsEnabled");
                }
            }
        }

        public bool ColumnIsEnabled
        {
            get { return Table != null; }
        }

        public string ComparisonOperator
        {
            get { return _sqlCond.Operator; }
            set
            {
                _sqlCond.Operator = value;
                OnPropertyChanged("ComparisonOperator");
                OnPropertyChanged("Values");
                OnPropertyChanged("QueryValueIsEnabled");
            }
        }

        public bool ComparisonOperatorIsEnabled
        {
            get { return Column != null; }
        }

        public string QueryValue
        {
            get { return _sqlCond.Value != null ? _sqlCond.Value.ToString() : null; }
            set
            {
                if (String.IsNullOrEmpty(_queryValueText) || (_queryValueRegex.Split(_queryValueText).Length > 1))
                    _sqlCond.Value = value;
                _cursorType = Cursors.Arrow;
                OnPropertyChanged("CursorType");
                OnPropertyChanged("QueryValue");
                OnPropertyChanged("QueryValueText");
            }
        }

        public string QueryValueText
        {
            get { return _sqlCond.Value != null ? _sqlCond.Value.ToString() : null; }
            set
            {
                _cursorType = Cursors.Arrow;
                OnPropertyChanged("CursorType");

                if ((_queryValueText == null) || !_queryValueText.Equals(value))
                {
                    if (!String.IsNullOrEmpty(value))
                        _queryValueText = _queryValueRegex.Split(value)[0];
                    else
                        _queryValueText = value;

                    _sqlCond.Value = _queryValueText;
                    OnPropertyChanged("QueryValueText");
                }
            }
        }

        public bool QueryValueIsEnabled
        {
            get
            {
                return Column != null && !String.IsNullOrEmpty(ComparisonOperator) &&
                    !_sqlCond.Operator.ToUpper().EndsWith("NULL");
            }
        }

        public string CloseParentheses
        {
            get { return _sqlCond.CloseParentheses; }
            set { _sqlCond.CloseParentheses = value; }
        }

        #endregion

        #endregion

        #region IDataErrorInfo Members

        public string Error
        {
            get
            {
                if ((_sqlCond.Table == null) || (_sqlCond.Column == null) ||
                    (_sqlCond.ColumnSystemType == null) || String.IsNullOrEmpty(_sqlCond.Operator) ||
                    (_sqlCond.Operator.ToUpper().EndsWith("NULL") != ((_sqlCond.Value == null) ||
                    _sqlCond.Value.ToString().Equals(String.Empty))))
                {
                    _isComplete = false;
                    return "Error: A table, column, operator and (except for the IS NULL operators) a value are required";
                }
                else
                {
                    _isComplete = true;
                    return String.Empty;
                }
            }
        }

        public string this[string columnName]
        {
            get
            {
                string error = null;

                switch (columnName)
                {
                    case "Table":
                        //error = "Please select a table name";
                        break;
                    case "Column":
                        //error = "Please select a column name";
                        break;
                    case "ComparisonOperator":
                        //error = "Please select an operator";
                        break;
                    case "QueryValue":
                        //error = "Please select or enter a query value";
                        break;
                }

                return error;
            }
        }

        #endregion
    }
}
