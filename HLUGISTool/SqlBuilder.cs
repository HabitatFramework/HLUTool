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
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HLU.Properties;

namespace HLU
{
    public class SqlFilterCondition : ICloneable
    {
        private string _booleanOperator;
        private string _openParentheses;
        private DataTable _table;
        private DataColumn _column;
        private Type _columnSystemType;
        private string _operator;
        private object _value;
        private string _closeParentheses;

        /// <summary>
        /// Defaults to "AND".
        /// </summary>
        public string BooleanOperator
        {
            get { return _booleanOperator; }
            set { _booleanOperator = value; }
        }

        /// <summary>
        /// Defaults to String.Empty.
        /// </summary>
        public string OpenParentheses
        {
            get { return _openParentheses; }
            set { _openParentheses = value; }
        }

        public DataTable Table
        {
            get { return _table; }
            set { _table = value;}
        }

        public DataColumn Column
        {
            get { return _column; }
            set { _column = value; }
        }

        /// <summary>
        /// Defaults to "=".
        /// </summary>
        public string Operator
        {
            get { return _operator; }
            set { _operator = value; }
        }

        /// <summary>
        /// Defaults to this.Column.DataType but can be set to typeof(DataColumn) to model table relations.
        /// </summary>
        public Type ColumnSystemType
        {
            get { return _columnSystemType != null ? _columnSystemType : Column.DataType; }
            set { _columnSystemType = value; }
        }

        public object Value
        {
            get { return _value; }
            set { _value = value; }
        }

        /// <summary>
        /// Defaults to String.Empty.
        /// </summary>
        public string CloseParentheses
        {
            get { return _closeParentheses; }
            set { _closeParentheses = value; }
        }

        public SqlFilterCondition()
        {
            SetDefaults();
        }

        public SqlFilterCondition(DataTable table, DataColumn column, object value)
        {
            _table = table;
            _column = column;
            _value = value;
            SetDefaults();
        }

        public SqlFilterCondition(string booleanOp, DataTable table, DataColumn column, object value)
        {
            _booleanOperator = booleanOp;
            _table = table;
            _column = column;
            _value = value;
            SetDefaults();
        }

        public SqlFilterCondition(string booleanOp, DataTable table, DataColumn column, Type systemDataType, string openParentheses, string closeParentheses, object value)
        {
            _booleanOperator = booleanOp;
            _table = table;
            _column = column;
            _columnSystemType = systemDataType;
            _openParentheses = openParentheses;
            _closeParentheses = closeParentheses;
            _value = value;
            SetDefaults();
        }

        private void SetDefaults()
        {
            if (String.IsNullOrEmpty(_booleanOperator))
                _booleanOperator = "AND";
            if (String.IsNullOrEmpty(_openParentheses))
                _openParentheses = String.Empty;
            _operator = "=";
            if (String.IsNullOrEmpty(_closeParentheses))
                _closeParentheses = String.Empty;
        }

        public SqlFilterCondition Clone()
        {
            return (SqlFilterCondition)((ICloneable)this).Clone();
        }

        object ICloneable.Clone()
        {
            return this.MemberwiseClone();
        }
    }

    public abstract class SqlBuilder
    {
        #region Abstract

        public abstract string QuotePrefix { get; }

        public abstract string QuoteSuffix { get; }

        public abstract string StringLiteralDelimiter { get; }

        public abstract string DateLiteralPrefix { get; }

        public abstract string DateLiteralSuffix { get; }

        public abstract string WildcardSingleMatch { get; }

        public abstract string WildcardManyMatch { get; }

        public abstract string QuoteValue(object value);

        public abstract DataTable SqlSelect(bool selectDistinct, DataColumn[] targetColumns, List<SqlFilterCondition> whereConds);

        public abstract DataTable SqlSelect(bool selectDistinct, DataTable[] targetTables, List<SqlFilterCondition> whereConds);

        #endregion

        #region Protected

        protected Dictionary<Type, Int32> _typeMapSystemToSQL;
        
        protected Dictionary<Int32, Type> _typeMapSQLToSystem;
        
        protected Type DbToSystemType(int dbTypeCode)
        {
            Type sysType;
            if (_typeMapSQLToSystem.TryGetValue(dbTypeCode, out sysType))
                return sysType;
            else
                return (Type)Type.Missing;
        }

        protected int SystemToDbType(Type sysType)
        {
            int typeCode;
            if (_typeMapSystemToSQL.TryGetValue(sysType, out typeCode))
                return typeCode;
            else
                return -1;
        }

        #endregion

        #region Public Methods
        
        public abstract string QuoteIdentifier(string identifier);

        public virtual bool QualifyColumnNames(DataColumn[] targetColumns)
        {
            if ((targetColumns == null) || (targetColumns.Length == 0)) return false;

            return (from c in targetColumns
                    group c by c.Table.TableName into g
                    select g).Count() > 1;
        }

        public virtual string ColumnAlias(DataColumn c)
        {
            if (c == null)
                return String.Empty;
            else
                return ColumnAlias(c.Table.TableName, c.ColumnName);
        }

        public virtual string ColumnAlias(string tableName, string columnName)
        {
            if (String.IsNullOrEmpty(columnName))
                return String.Empty;
            else if (String.IsNullOrEmpty(tableName))
                return columnName;
            else
                return tableName + Resources.ColumnTableNameSeparator + columnName;
        }

        public abstract string TargetList(DataColumn[] targetColumns, bool quoteIdentifiers, 
            bool checkQualify, ref bool qualifyColumns, out DataTable resultTable);

        public string TargetList(DataTable[] targetTables, bool quoteIdentifiers, 
            ref bool qualifyColumns, out DataTable resultTable)
        {
            if ((targetTables == null) || (targetTables.Length == 0) || (targetTables[0].Columns.Count == 0))
            {
                resultTable = new DataTable();
                return String.Empty;
            }

            resultTable = null;

            try
            {
                DataColumn[] targetList = targetTables.SelectMany(t => t.Columns.Cast<DataColumn>()).ToArray();
                qualifyColumns = targetTables.Length > 1;
                return TargetList(targetList, quoteIdentifiers, false, ref qualifyColumns, out resultTable);
            }
            catch { resultTable = new DataTable(); }
            return String.Empty;
        }

        public string FromList(bool includeFrom, bool quoteIdentifiers,
            DataTable[] targetTables, ref List<SqlFilterCondition> whereClause, out bool additionalTables)
        {
            DataColumn[] targetColumns = targetTables.SelectMany(t => t.Columns.Cast<DataColumn>()).ToArray();
            return FromList(includeFrom, targetColumns, quoteIdentifiers, ref whereClause, out additionalTables);
        }

        public string FromList(bool includeFrom, DataColumn[] targetColumns, 
            bool quoteIdentifiers, ref List<SqlFilterCondition> whereClause, out bool additionalTables)
        {
            DataTable[] colTables = targetColumns.Select(c => c.Table).Distinct().ToArray();
            var whereTables = whereClause.Select(con => con.Table).Distinct().Where(t => !colTables.Contains(t));

            int numTables = colTables.Length;
            colTables = colTables.Concat(whereTables).ToArray();
            additionalTables = colTables.Length > numTables;

            whereClause = JoinClause(colTables).Concat(whereClause).ToList();
            
            return FromList(includeFrom, quoteIdentifiers, colTables.Select(t => t.TableName).ToArray());
        }

        /// <summary>
        /// Create a string of database tables that the data will be selected from.
        /// </summary>
        /// <param name="includeFrom">If set to <c>true</c> include 'FROM' in the returned string.</param>
        /// <param name="quoteIdentifiers">If set to <c>true</c> wrap identifiers in quotes.</param>
        /// <param name="targetColumns">The target columns to be selected.</param>
        /// <param name="fromTables">A list of the tables to select from.</param>
        /// <param name="whereClause">The where clause statements generated to satisfy the table joins.</param>
        /// <param name="additionalTables">Set to <c>true</c> if tables, in addition to those relating
        /// to the target columns, are required.</param>
        /// <returns>A string of database tables to select from.</returns>
        public string FromList(bool includeFrom, bool quoteIdentifiers, DataColumn[] targetColumns,
            List<DataTable> fromTables, ref List<SqlFilterCondition> whereClause, out bool additionalTables)
        {
            DataTable[] colTables = targetColumns.Select(c => c.Table).Distinct().ToArray();
            var whereTables = fromTables.Distinct().Where(t => !colTables.Contains(t));

            int numTables = colTables.Length;
            colTables = colTables.Concat(whereTables).ToArray();
            additionalTables = colTables.Length > numTables;

            whereClause = JoinClause(colTables).ToList();

            return FromList(includeFrom, quoteIdentifiers, colTables.Select(t => t.TableName).ToArray());
        }

        public string FromList(bool includeFrom, bool quoteIdentifiers, string[] tableNames)
        {
            if ((tableNames == null) || (tableNames.Length == 0)) return String.Empty;
            StringBuilder sbFromList = new StringBuilder();
            if (quoteIdentifiers)
            {
                foreach (string tableName in tableNames)
                    sbFromList.Append(",").Append(QuoteIdentifier(tableName));
                return (includeFrom ? " FROM " : "") + sbFromList.Remove(0, 1).ToString();
            }
            else
            {
                return (includeFrom ? " FROM " : "") + String.Join(",", tableNames);
            }
        }

        public virtual List<List<SqlFilterCondition>> JoinWhereClauseLists(List<List<SqlFilterCondition>> inWhereClause)
        {
            List<List<SqlFilterCondition>> outWhereClause = new List<List<SqlFilterCondition>>();
            foreach (List<SqlFilterCondition> oneWhereClause in inWhereClause)
            {
                if ((outWhereClause.Count == 0) || (oneWhereClause[0].BooleanOperator.ToUpper() == "OR"))
                    outWhereClause.Add(oneWhereClause);
                else
                    outWhereClause[outWhereClause.Count - 1].AddRange(oneWhereClause);
            }
            return outWhereClause;
        }

        public virtual string WhereClause(bool includeWhere, bool quoteIdentifiers,
            bool qualifyColumns, List<List<SqlFilterCondition>> whereConds)
        {
            return WhereClause(includeWhere, quoteIdentifiers, qualifyColumns, 
                whereConds.SelectMany(cond => cond).ToList());
        }

        public virtual string WhereClause(bool includeWhere, bool quoteIdentifiers, 
            bool qualifyColumns, List<SqlFilterCondition> whereConds)
        {
            if ((whereConds != null) && (whereConds.Count > 0))
            {
                StringBuilder sbWhereClause = new StringBuilder(includeWhere ? " WHERE " : " ");
                for (int i = 0; i < whereConds.Count; i++)
                {
                    SqlFilterCondition sqlCond = whereConds[i];

                    if (i != 0)
                    {
                        if (!String.IsNullOrEmpty(sqlCond.BooleanOperator))
                            sbWhereClause.Append(String.Format(" {0} ", sqlCond.BooleanOperator));
                        else
                            sbWhereClause.Append(" AND ");
                    }
                    
                    sbWhereClause.Append(sqlCond.OpenParentheses);

                    if (quoteIdentifiers)
                    {
                        if (qualifyColumns && !String.IsNullOrEmpty(sqlCond.Table.TableName))
                            sbWhereClause.Append(String.Format("{0}.{1}", QuoteIdentifier(sqlCond.Table.TableName), 
                                QuoteIdentifier(sqlCond.Column.ColumnName)));
                        else
                            sbWhereClause.Append(QuoteIdentifier(sqlCond.Column.ColumnName));
                    }
                    else
                    {
                        if (qualifyColumns && !String.IsNullOrEmpty(sqlCond.Table.TableName))
                            sbWhereClause.Append(String.Format("{0}.{1}", sqlCond.Table.TableName, sqlCond.Column.ColumnName));
                        else
                            sbWhereClause.Append(sqlCond.Column.ColumnName);
                    }

                    DataColumn c;
                    if (!String.IsNullOrEmpty(sqlCond.Operator))
                    {
                        if ((sqlCond.ColumnSystemType == typeof(DataColumn)) && 
                            ((c = sqlCond.Value as DataColumn) != null)) // table relation
                        {
                            if (quoteIdentifiers)
                            {
                                if (qualifyColumns)
                                    sbWhereClause.Append(String.Format(" {0} {1}", sqlCond.Operator,
                                        QuoteIdentifier(c.Table.TableName) + "." + QuoteIdentifier(c.ColumnName)));
                                else
                                    sbWhereClause.Append(String.Format(" {0} {1}", 
                                        sqlCond.Operator, QuoteIdentifier(c.ColumnName)));
                            }
                            else
                            {
                                if (qualifyColumns)
                                    sbWhereClause.Append(String.Format(" {0} {1}",
                                        sqlCond.Operator, c.Table.TableName + "." + c.ColumnName));
                                else
                                    sbWhereClause.Append(String.Format(" {0} {1}", sqlCond.Operator, c.ColumnName));
                            }
                        }
                        else if (sqlCond.Operator.ToUpper().EndsWith("NULL"))
                        {
                            sbWhereClause.Append(String.Format(" {0} ", sqlCond.Operator));
                        }
                        else if (sqlCond.ColumnSystemType == typeof(System.String))
                        {
                            switch (sqlCond.Operator.ToUpper())
                            {
                                case "IN ()":
                                case "NOT IN ()":
                                    sbWhereClause.Append(String.Format(" {0}", 
                                        sqlCond.Operator.Remove(sqlCond.Operator.Length - 1, 1)));
                                    Regex r = new Regex(QuotePrefix + @"[^" + QuotePrefix + "]*" + QuoteSuffix + "|[^,]+",
                                    RegexOptions.IgnorePatternWhitespace);
                                    sbWhereClause.Append(r.Matches(sqlCond.Value.ToString()).Cast<Match>()
                                        .Aggregate(new StringBuilder(), (sb, m) => sb.Append(String.Format("{0},", 
                                            QuoteValue(m.Value)))));
                                    sbWhereClause.Remove(sbWhereClause.Length - 1, 1);
                                    sbWhereClause.Append(")");
                                    break;
                                case "BEGINS WITH":
                                    sbWhereClause.Append(" LIKE " + QuoteValue(String.Format("{0}{1}",
                                        sqlCond.Value, WildcardManyMatch)));
                                    break;
                                case "ENDS WITH":
                                    sbWhereClause.Append(" LIKE " + QuoteValue(String.Format("{1}{0}",
                                        sqlCond.Value, WildcardManyMatch)));
                                    break;
                                case "CONTAINS":
                                    sbWhereClause.Append(" LIKE " + QuoteValue(String.Format("{1}{0}{1}",
                                        sqlCond.Value, WildcardManyMatch)));
                                    break;
                                default:
                                    sbWhereClause.Append(String.Format(" {0} {1}", sqlCond.Operator,
                                        QuoteValue(GetUnderlyingType(sqlCond))));
                                    break;
                            }
                        }
                        else
                        {
                            switch (sqlCond.Operator.ToUpper())
                            {
                                case "IN ()":
                                    sbWhereClause.Append(" IN (").Append(sqlCond.Value).Append(") ");
                                    break;
                                case "NOT IN ()":
                                    sbWhereClause.Append(" NOT IN (").Append(sqlCond.Value).Append(") ");
                                    break;
                                default:
                                    sbWhereClause.Append(" ").Append(sqlCond.Operator).Append(" ");
                                    sbWhereClause.Append(QuoteValue(GetUnderlyingType(sqlCond)));
                                    break;
                            }
                        }
                    }
                    sbWhereClause.Append(sqlCond.CloseParentheses);
                }
                return sbWhereClause.ToString();
            }
            else
            {
                return string.Empty;
            }
        }
        
        #endregion

        #region Private Methods

        private List<SqlFilterCondition> JoinClause(DataTable[] queryTables)
        {
            if ((queryTables == null) || (queryTables.Length == 0)) return null;

            List<SqlFilterCondition> joinClause = new List<SqlFilterCondition>();

            for (int i = 0; i < queryTables.Length; i++)
            {
                DataTable p = queryTables[i];
                for (int j = 0; j < queryTables.Length; j++)
                {
                    if (j == i) continue;
                    DataTable c = queryTables[j];
                    var children = p.ChildRelations.Cast<DataRelation>().Where(r => r.ChildTable == c);
                    if (children.Count() > 0)
                    {
                        DataRelation r = children.ElementAt(0);
                        for (int k = 0; k < r.ParentColumns.Length; k++)
                        {
                            SqlFilterCondition joinWhere = new SqlFilterCondition();
                            joinWhere.BooleanOperator = String.Empty;
                            if (k == 0)
                            {
                                joinWhere.OpenParentheses = "((";
                                joinWhere.CloseParentheses = "))";
                            }
                            else
                            {
                                joinWhere.OpenParentheses = "(";
                                joinWhere.CloseParentheses = ")";
                            }
                            joinWhere.Table = r.ParentTable;
                            joinWhere.Column = r.ParentColumns[k];
                            joinWhere.ColumnSystemType = typeof(DataColumn);
                            joinWhere.Operator = "=";
                            joinWhere.Value = r.ChildColumns[k];
                            joinClause.Add(joinWhere);
                        }
                    }
                }
            }

            return joinClause;
        }

        private object GetUnderlyingType(SqlFilterCondition sqlCond)
        {
            try
            {
                DataRow row = sqlCond.Table.NewRow();
                row[sqlCond.Column] = sqlCond.Value;
                return row[sqlCond.Column];
            }
            catch { return sqlCond.Value; }
        }
        
        #endregion
    }
}
