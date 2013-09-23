using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HLU.Data.Model;

namespace HLU.UI.ViewModel
{
    static class ViewModelWindowMainHelpers
    {
        public static List<List<SqlFilterCondition>> GisSelectionToWhereClause<T>(
            DataRow[] selectedRows, int[] keyColumOrdinals, int blockSize, T targetTable)
            where T : DataTable
        {
            DataTable selectionTable = selectedRows[0].Table;

            List<List<SqlFilterCondition>> whereClause = new List<List<SqlFilterCondition>>();

            int i = 0;
            while (i < selectedRows.Length)
            {
                List<SqlFilterCondition> whereClauseBlock = new List<SqlFilterCondition>();
                int j = i;

                while (j < selectedRows.Length)
                {
                    DataRow r = selectedRows[j];

                    for (int k = 0; k < keyColumOrdinals.Length; k++)
                    {
                        SqlFilterCondition cond = new SqlFilterCondition();

                        if (k == 0)
                        {
                            cond.BooleanOperator = "OR";
                            cond.OpenParentheses = "(";
                        }
                        else
                        {
                            cond.BooleanOperator = "AND";
                            cond.OpenParentheses = String.Empty;
                        }
                        cond.Column = selectionTable.Columns[keyColumOrdinals[k]];
                        cond.Table = targetTable;
                        cond.ColumnSystemType = selectionTable.Columns[k].DataType;
                        cond.Operator = "=";
                        cond.Value = r[keyColumOrdinals[k]];
                        if (k == keyColumOrdinals.Length - 1)
                            cond.CloseParentheses = ")";
                        else
                            cond.CloseParentheses = String.Empty;

                        whereClauseBlock.Add(cond);
                    }

                    j++;
                    if (whereClauseBlock.Count >= blockSize)
                        break;
                }

                if (whereClauseBlock.Count > 0)
                    whereClause.Add(whereClauseBlock);

                i = j;
            }

            return whereClause;
        }

        public static List<List<SqlFilterCondition>> IncidSelectionToWhereClause<T>(int incidPageSize,
            int incidOrdinal, T incidTable, IEnumerable<string> incidList) where T : DataTable
        {
            if ((incidList == null) || (incidList.Count() == 0)) return null;

            return (from b in incidList.Select((i, index) => new
            {
                Block = index / incidPageSize,
                Condition = new SqlFilterCondition("OR", incidTable, incidTable.Columns[incidOrdinal], i)
            })
                    group b by b.Block into g
                    select g.Select(b => b.Condition).ToList()).ToList();
        }

        public static string IhsSummary(string[] ihsCodes)
        {
            StringBuilder buildSummary = ihsCodes.Where(c => !String.IsNullOrEmpty(c))
                .Aggregate(new StringBuilder(), (sb, c) => sb.Append(String.Format(".{0}", c)));
            if (buildSummary.Length > 1)
                return buildSummary.Remove(0, 1).ToString();
            else
                return null;
        }

        public static bool RowIsDirty<R>(R row)
            where R : DataRow
        {
            return row != null && row.RowState != DataRowState.Unchanged && 
                row.RowState != DataRowState.Detached;
        }

        public static string GetOperationsCode(HluDataSet hluDS, ViewModelWindowMain.Operations modifyOperation)
        {
            if ((hluDS == null) || (hluDS.lut_operation == null)) return null;

            string operationName = Enum.GetName(typeof(ViewModelWindowMain.Operations), modifyOperation);
            string descriptionPattern = Regex.Matches(operationName, "[A-Z][^A-Z]*").Cast<Match>()
                .Aggregate(new StringBuilder(), (sb, m) => sb.Append(@"\s*" + operationName.Substring(m.Index, m.Length)))
                .Append(@"\s*").ToString();

            var o = hluDS.lut_operation
                .Where(r => Regex.IsMatch(r.description, descriptionPattern, RegexOptions.IgnoreCase));
            if (o.Count() == 1)
                return o.First().code;
            else
                return null;
        }
    }
}
