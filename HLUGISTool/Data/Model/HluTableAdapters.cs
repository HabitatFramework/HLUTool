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
using System.Data.Common;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using HLU.Data.Connection;

namespace HLU.Data.Model.HluDataSetTableAdapters
{
    public partial class HluTableAdapter<T, R> : Component
        where T : DataTable, new()
        where R : DataRow
    {
        #region Fields

        private IDbDataAdapter _adapter;

        private IDbCommand[] _commandCollection;

        private bool _clearBeforeFill;

        private DbBase _db;

        private T _hluTable;

        private int _columnCount;

        private string _originalSelectCommand;

        private Dictionary<string, Type> _columnsDic;
        
        private Dictionary<string, int> _paramsDelOrig;
        
        private Dictionary<string, int> _paramsUpdCurr;
        
        private Dictionary<string, int> _paramsUpdOrig;

        #endregion

        #region Constructor

        internal HluTableAdapter(DbBase db)
        {
            try
            {
                if (db == null) throw new ArgumentNullException("db");

                if (typeof(T).GetProperty("Item").PropertyType != typeof(R))
                    throw new ArgumentException("Type parameter R must be the row type of T.", "R");

                var columns = typeof(T).GetProperties().Where(pi => pi.PropertyType == typeof(DataColumn));
                _columnsDic = columns.ToDictionary(pi => pi.Name, pi => pi.PropertyType);

                _db = db;
                this.ClearBeforeFill = true;
            }
            catch { throw; }
        }

        #endregion

        protected internal IDbDataAdapter Adapter
        {
            get
            {
                if ((this._adapter == null)) this.InitAdapter();
                return this._adapter;
            }
        }

        internal IDbConnection Connection { get { return _db.Connection; } }

        internal IDbTransaction Transaction { get { return _db.Transaction; } }

        protected IDbCommand[] CommandCollection
        {
            get
            {
                if ((this._commandCollection == null)) this.InitCommandCollection();
                return this._commandCollection;
            }
        }

        public bool ClearBeforeFill
        {
            get { return this._clearBeforeFill; }
            set { this._clearBeforeFill = value; }
        }

        private void InitAdapter()
        {
            if (_hluTable == null)
            {
                _hluTable = new T();
                _columnCount = _hluTable.Columns.Count;
            }

            IDbDataAdapter adapter = _db.CreateAdapter(_hluTable);

            if (adapter == null)
                throw new Exception(String.Format("Table '{0}' has no primary key.",
                    _db.QuoteIdentifier(_hluTable.TableName)));
            else
                _adapter = adapter;

            if (this.Adapter != null)
            {
                _originalSelectCommand = adapter.SelectCommand.CommandText;

                _paramsDelOrig = (from DbParameter p in _adapter.DeleteCommand.Parameters
                                  where p.Direction == ParameterDirection.Input &&
                                  p.SourceVersion == DataRowVersion.Original && !p.SourceColumnNullMapping
                                  select new
                                  {
                                      key = p.SourceColumn,
                                      value = _adapter.DeleteCommand.Parameters.IndexOf(p)
                                  }
                                  ).ToDictionary(kv => kv.key, kv => kv.value);

                _paramsUpdCurr = new Dictionary<string, int>();
                _paramsUpdOrig = new Dictionary<string, int>();

                for (int i = 0; i < _adapter.UpdateCommand.Parameters.Count; i++)
                {
                    DbParameter p = (DbParameter)_adapter.UpdateCommand.Parameters[i];
                    if (p.Direction != ParameterDirection.Input) continue;
                    switch (p.SourceVersion)
                    {
                        case DataRowVersion.Current:
                            if (!String.IsNullOrEmpty(p.SourceColumn)) 
                                _paramsUpdCurr.Add(p.SourceColumn, i);
                            break;
                        case DataRowVersion.Original:
                            if (!p.SourceColumnNullMapping) 
                                _paramsUpdOrig.Add(p.SourceColumn, i);
                            break;
                    }
                }
            }
        }

        private void InitCommandCollection()
        {
            this._commandCollection = new IDbCommand[1];
            if ((this._adapter == null) || (this._adapter.SelectCommand == null)) InitAdapter();
            this._commandCollection[0] = _db.CreateCommand();
            this._commandCollection[0].Connection = this.Connection;
            this._commandCollection[0].CommandText = this._adapter.SelectCommand.CommandText;
            this._commandCollection[0].CommandType = CommandType.Text;
        }

        [DataObjectMethodAttribute(DataObjectMethodType.Fill, true)]
        public virtual int Fill(T dataTable)
        {
            if (dataTable == null) dataTable = new T();
            this.Adapter.SelectCommand = this.CommandCollection[0];
            if ((this.ClearBeforeFill == true)) dataTable.Clear();
            int returnValue = _db.FillTable(this.Adapter.SelectCommand.CommandText, ref dataTable);
            return returnValue;
        }

        public virtual int Fill(T dataTable, string whereClause)
        {
            if (!String.IsNullOrEmpty(whereClause))
            {
                if (dataTable == null) dataTable = new T();
                this.CommandCollection[0].CommandText = _originalSelectCommand +
                    (!whereClause.TrimStart().ToUpper().StartsWith("WHERE") ? " WHERE " : "") + whereClause;
                return Fill(dataTable);
            }
            return -1;
        }

        public virtual int Fill(T dataTable, List<SqlFilterCondition> whereClause)
        {
            if ((whereClause == null) || (whereClause.Count == 0))
                return Fill(dataTable);

            try
            {
                if (dataTable == null)
                    dataTable = new T();
                else if ((this.ClearBeforeFill == true))
                    dataTable.Clear();

                this.CommandCollection[0].CommandText = _originalSelectCommand +
                    _db.WhereClause(true, true, true, whereClause);
                
                return Fill(dataTable);
            }
            catch { return -1; }
        }

        public virtual int Fill(T dataTable, List<List<SqlFilterCondition>> whereClause)
        {
            if ((whereClause == null) || (whereClause.Count == 0))
                return Fill(dataTable);

            try
            {
                if (dataTable == null)
                    dataTable = new T();
                else if ((this.ClearBeforeFill == true))
                    dataTable.Clear();

                bool backupClearBeforeFill = this.ClearBeforeFill;
                this.ClearBeforeFill = false;

                foreach (List<SqlFilterCondition> oneWhereClause in whereClause)
                {
                    this.CommandCollection[0].CommandText = _originalSelectCommand +
                        _db.WhereClause(true, true, true, oneWhereClause);
                    Fill(dataTable);
                }

                this.ClearBeforeFill = backupClearBeforeFill;

                return dataTable.Rows.Count;
            }
            catch { return -1; }
        }

        [DataObjectMethodAttribute(DataObjectMethodType.Select, true)]
        public virtual T GetData()
        {
            this.Adapter.SelectCommand = this.CommandCollection[0];
            _db.FillTable(ref _hluTable);
            return _hluTable;
        }

        public virtual int Update(T dataTable)
        {
            return _db.Update<T>(dataTable);
        }

        public virtual int Update(HluDataSet dataSet)
        {
            return this.Adapter.Update(dataSet);
        }

        public virtual int Update(R dataRow)
        {
            return _db.Update<T, R>(new R[] { dataRow });
        }

        public virtual int Update(R[] dataRows)
        {
            return _db.Update<T, R>(dataRows);
        }

        [DataObjectMethodAttribute(DataObjectMethodType.Delete, true)]
        public virtual int Delete(R originalRow)
        {
            ConnectionState previousConnectionState = this.Adapter.UpdateCommand.Connection.State;

            try
            {
                // on any error the entire operation should fail, so we don't check whether parameters are found
                foreach (KeyValuePair<string, Type> kv in _columnsDic)
                {
                    DataColumn col = 
                        (DataColumn)originalRow.Table.GetType().GetProperty(kv.Key).GetValue(originalRow.Table, null);

                    ((IDataParameter)this.Adapter.DeleteCommand.Parameters[
                        _paramsDelOrig.Single(p => p.Key == col.ColumnName).Value]).Value =
                        (originalRow.IsNull(col)) ? (object)DBNull.Value : originalRow[col];
                }

                if ((this.Adapter.DeleteCommand.Connection.State & ConnectionState.Open) != ConnectionState.Open)
                    this.Adapter.DeleteCommand.Connection.Open();

                if (_db.Transaction != null) this.Adapter.DeleteCommand.Transaction = _db.Transaction;

                int returnValue = this.Adapter.DeleteCommand.ExecuteNonQuery();

                return returnValue;
            }
            finally
            {
                if ((previousConnectionState == ConnectionState.Closed))
                    this.Adapter.DeleteCommand.Connection.Close();
            }
        }

        [DataObjectMethodAttribute(DataObjectMethodType.Insert, true)]
        public virtual int Insert(R row)
        {
            ConnectionState previousConnectionState = this.Adapter.InsertCommand.Connection.State;

            try
            {
                for (int i = 0; i < _columnCount; i++)
                    ((IDataParameter)this.Adapter.InsertCommand.Parameters[i]).Value = 
                        row.IsNull(row.Table.Columns[i]) ? DBNull.Value : row[i];

                if ((this.Adapter.InsertCommand.Connection.State & ConnectionState.Open) != ConnectionState.Open)
                    this.Adapter.InsertCommand.Connection.Open();

                if (_db.Transaction != null) this.Adapter.InsertCommand.Transaction = _db.Transaction;
                
                int returnValue = this.Adapter.InsertCommand.ExecuteNonQuery();
                
                return returnValue;
            }
            finally
            {
                if ((previousConnectionState == ConnectionState.Closed))
                    this.Adapter.InsertCommand.Connection.Close();
            }
        }

        [DataObjectMethodAttribute(DataObjectMethodType.Update, true)]
        public virtual int Update(R newRow, R originalRow)
        {
            ConnectionState previousConnectionState = this.Adapter.UpdateCommand.Connection.State;
           
            try
            {
                // on any error the entire operation should fail, so we don't check whether parameters are found
                foreach (KeyValuePair<string, Type> kv in _columnsDic)
                {
                    DataColumn col =
                        (DataColumn)originalRow.Table.GetType().GetProperty(kv.Key).GetValue(originalRow.Table, null);

                    ((IDataParameter)this.Adapter.UpdateCommand.Parameters[
                        _paramsUpdOrig.Single(p => p.Key == col.ColumnName).Value]).Value =
                        (originalRow.IsNull(col)) ? (object)DBNull.Value : originalRow[col];
                }

                if ((this.Adapter.UpdateCommand.Connection.State & ConnectionState.Open) != ConnectionState.Open)
                    this.Adapter.UpdateCommand.Connection.Open();

                if (_db.Transaction != null) this.Adapter.UpdateCommand.Transaction = _db.Transaction;

                int returnValue = this.Adapter.UpdateCommand.ExecuteNonQuery();
                
                return returnValue;
            }
            finally
            {
                if ((previousConnectionState == ConnectionState.Closed))
                    this.Adapter.UpdateCommand.Connection.Close();
            }
        }
    }

    /// <summary>
    /// TableAdapterManager is used to coordinate TableAdapters in the dataset to enable Hierarchical Update scenarios
    ///</summary>
    [DesignerCategoryAttribute("code")]
    [ToolboxItem(true)]
    public partial class TableAdapterManager : Component
    {
        #region Fields

        #region TableAdapters

        private HluTableAdapter<HluDataSet.exportsDataTable, HluDataSet.exportsRow> _exportsTableAdapter;

        private HluTableAdapter<HluDataSet.exports_fieldsDataTable, HluDataSet.exports_fieldsRow> _exportsFieldsTableAdapter;

        private HluTableAdapter<HluDataSet.historyDataTable, HluDataSet.historyRow> _historyTableAdapter;

        private HluTableAdapter<HluDataSet.incidDataTable, HluDataSet.incidRow> _incidTableAdapter;

        private HluTableAdapter<HluDataSet.incid_bapDataTable, HluDataSet.incid_bapRow> _incid_bapTableAdapter;

        private HluTableAdapter<HluDataSet.incid_ihs_complexDataTable, HluDataSet.incid_ihs_complexRow> _incid_ihs_complexTableAdapter;

        private HluTableAdapter<HluDataSet.incid_ihs_formationDataTable, HluDataSet.incid_ihs_formationRow> _incid_ihs_formationTableAdapter;

        private HluTableAdapter<HluDataSet.incid_ihs_managementDataTable, HluDataSet.incid_ihs_managementRow> _incid_ihs_managementTableAdapter;

        private HluTableAdapter<HluDataSet.incid_ihs_matrixDataTable, HluDataSet.incid_ihs_matrixRow> _incid_ihs_matrixTableAdapter;

        private HluTableAdapter<HluDataSet.incid_mm_polygonsDataTable, HluDataSet.incid_mm_polygonsRow> _incid_mm_polygonsTableAdapter;

        private HluTableAdapter<HluDataSet.incid_sourcesDataTable, HluDataSet.incid_sourcesRow> _incid_sourcesTableAdapter;

        private HluTableAdapter<HluDataSet.lut_bap_quality_determinationDataTable, HluDataSet.lut_bap_quality_determinationRow> _lut_bap_quality_determinationTableAdapter;

        private HluTableAdapter<HluDataSet.lut_bap_quality_interpretationDataTable, HluDataSet.lut_bap_quality_interpretationRow> _lut_bap_quality_interpretationTableAdapter;

        private HluTableAdapter<HluDataSet.lut_boundary_mapDataTable, HluDataSet.lut_boundary_mapRow> _lut_boundary_mapTableAdapter;

        private HluTableAdapter<HluDataSet.lut_habitat_classDataTable, HluDataSet.lut_habitat_classRow> _lut_habitat_classTableAdapter;

        private HluTableAdapter<HluDataSet.lut_habitat_typeDataTable, HluDataSet.lut_habitat_typeRow> _lut_habitat_typeTableAdapter;

        private HluTableAdapter<HluDataSet.lut_habitat_type_ihs_habitatDataTable, HluDataSet.lut_habitat_type_ihs_habitatRow> _lut_habitat_type_ihs_habitatTableAdapter;

        private HluTableAdapter<HluDataSet.lut_ihs_complexDataTable, HluDataSet.lut_ihs_complexRow> _lut_ihs_complexTableAdapter;

        private HluTableAdapter<HluDataSet.lut_ihs_formationDataTable, HluDataSet.lut_ihs_formationRow> _lut_ihs_formationTableAdapter;

        private HluTableAdapter<HluDataSet.lut_ihs_habitatDataTable, HluDataSet.lut_ihs_habitatRow> _lut_ihs_habitatTableAdapter;

        private HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_complexDataTable, HluDataSet.lut_ihs_habitat_ihs_complexRow> _lut_ihs_habitat_ihs_complexTableAdapter;

        private HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_formationDataTable, HluDataSet.lut_ihs_habitat_ihs_formationRow> _lut_ihs_habitat_ihs_formationTableAdapter;

        private HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_managementDataTable, HluDataSet.lut_ihs_habitat_ihs_managementRow> _lut_ihs_habitat_ihs_managementTableAdapter;

        private HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_matrixDataTable, HluDataSet.lut_ihs_habitat_ihs_matrixRow> _lut_ihs_habitat_ihs_matrixTableAdapter;

        private HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_nvcDataTable, HluDataSet.lut_ihs_habitat_ihs_nvcRow> _lut_ihs_habitat_ihs_nvcTableAdapter;

        private HluTableAdapter<HluDataSet.lut_ihs_managementDataTable, HluDataSet.lut_ihs_managementRow> _lut_ihs_ManagementTableAdapter;

        private HluTableAdapter<HluDataSet.lut_ihs_matrixDataTable, HluDataSet.lut_ihs_matrixRow> _lut_ihs_matrixTableAdapter;

        private HluTableAdapter<HluDataSet.lut_importanceDataTable, HluDataSet.lut_importanceRow> _lut_importanceTableAdapter;

        private HluTableAdapter<HluDataSet.lut_last_incidDataTable, HluDataSet.lut_last_incidRow> _lut_last_incidTableAdapter;

        private HluTableAdapter<HluDataSet.lut_legacy_habitatDataTable, HluDataSet.lut_legacy_habitatRow> _lut_legacy_habitatTableAdapter;

        private HluTableAdapter<HluDataSet.lut_operationDataTable, HluDataSet.lut_operationRow> _lut_operationTableAdapter;

        private HluTableAdapter<HluDataSet.lut_processDataTable, HluDataSet.lut_processRow> _lut_processTableAdapter;

        private HluTableAdapter<HluDataSet.lut_reasonDataTable, HluDataSet.lut_reasonRow> _lut_reasonTableAdapter;

        private HluTableAdapter<HluDataSet.lut_site_idDataTable, HluDataSet.lut_site_idRow> _lut_site_idTableAdapter;

        private HluTableAdapter<HluDataSet.lut_sourcesDataTable, HluDataSet.lut_sourcesRow> _lut_sourcesTableAdapter;

        private HluTableAdapter<HluDataSet.lut_userDataTable, HluDataSet.lut_userRow> _lut_userTableAdapter;

        private HluTableAdapter<HluDataSet.lut_versionDataTable, HluDataSet.lut_versionRow> _lut_versionTableAdapter;

        #endregion

        private UpdateOrderOption _updateOrder;

        private bool _backupDataSetBeforeUpdate;

        private DbBase _db;

        private Dictionary<Type, PropertyInfo> _tableAdapterMatches;

        private string _sameConnErrorMsg =
            "All TableAdapters managed by a TableAdapterManager must use the same connection string.";

        public static Type[] DataTableTypes = new Type[] { typeof(HluDataSet.incidDataTable),
                typeof(HluDataSet.incid_bapDataTable), typeof(HluDataSet.incid_ihs_complexDataTable), 
                typeof(HluDataSet.incid_ihs_formationDataTable), typeof(HluDataSet.incid_ihs_managementDataTable), 
                typeof(HluDataSet.incid_ihs_matrixDataTable), typeof(HluDataSet.incid_sourcesDataTable) };

        public static Type[] LookupTableTypes = new Type[] { typeof(HluDataSet.exportsDataTable), 
                typeof(HluDataSet.exports_fieldsDataTable),  
                typeof(HluDataSet.lut_bap_quality_determinationDataTable), 
                typeof(HluDataSet.lut_bap_quality_interpretationDataTable), 
                typeof(HluDataSet.lut_boundary_mapDataTable), typeof(HluDataSet.lut_habitat_classDataTable), 
                typeof(HluDataSet.lut_habitat_typeDataTable), typeof(HluDataSet.lut_habitat_type_ihs_habitatDataTable), 
                typeof(HluDataSet.lut_ihs_complexDataTable), typeof(HluDataSet.lut_ihs_formationDataTable), 
                typeof(HluDataSet.lut_ihs_habitatDataTable), typeof(HluDataSet.lut_ihs_habitat_ihs_complexDataTable), 
                typeof(HluDataSet.lut_ihs_habitat_ihs_formationDataTable), 
                typeof(HluDataSet.lut_ihs_habitat_ihs_managementDataTable), 
                typeof(HluDataSet.lut_ihs_habitat_ihs_matrixDataTable), typeof(HluDataSet.lut_ihs_habitat_ihs_nvcDataTable), 
                typeof(HluDataSet.lut_ihs_managementDataTable), typeof(HluDataSet.lut_ihs_matrixDataTable), 
                typeof(HluDataSet.lut_importanceDataTable),
                typeof(HluDataSet.lut_last_incidDataTable), typeof(HluDataSet.lut_operationDataTable), 
                typeof(HluDataSet.lut_processDataTable), typeof(HluDataSet.lut_reasonDataTable), 
                typeof(HluDataSet.lut_site_idDataTable), typeof(HluDataSet.lut_sourcesDataTable), 
                typeof(HluDataSet.lut_userDataTable), typeof(HluDataSet.lut_versionDataTable) };
        #endregion

        #region Properties

        #region Table Adapters

        public HluTableAdapter<HluDataSet.exportsDataTable, HluDataSet.exportsRow> exportsTableAdapter
        {
            get { return this._exportsTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._exportsTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.exports_fieldsDataTable, HluDataSet.exports_fieldsRow> exportsFieldsTableAdapter
        {
            get { return this._exportsFieldsTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._exportsFieldsTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_versionDataTable, HluDataSet.lut_versionRow> lut_versionTableAdapter
        {
            get { return this._lut_versionTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_versionTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.historyDataTable, HluDataSet.historyRow> historyTableAdapter
        {
            get { return this._historyTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._historyTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.incidDataTable, HluDataSet.incidRow> incidTableAdapter
        {
            get { return this._incidTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._incidTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.incid_bapDataTable, HluDataSet.incid_bapRow> incid_bapTableAdapter
        {
            get { return this._incid_bapTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._incid_bapTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.incid_ihs_complexDataTable, HluDataSet.incid_ihs_complexRow> incid_ihs_complexTableAdapter
        {
            get { return this._incid_ihs_complexTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._incid_ihs_complexTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.incid_ihs_formationDataTable, HluDataSet.incid_ihs_formationRow> incid_ihs_formationTableAdapter
        {
            get { return this._incid_ihs_formationTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._incid_ihs_formationTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.incid_ihs_managementDataTable, HluDataSet.incid_ihs_managementRow> incid_ihs_managementTableAdapter
        {
            get { return this._incid_ihs_managementTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._incid_ihs_managementTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.incid_ihs_matrixDataTable, HluDataSet.incid_ihs_matrixRow> incid_ihs_matrixTableAdapter
        {
            get { return this._incid_ihs_matrixTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._incid_ihs_matrixTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.incid_mm_polygonsDataTable, HluDataSet.incid_mm_polygonsRow> incid_mm_polygonsTableAdapter
        {
            get { return this._incid_mm_polygonsTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._incid_mm_polygonsTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.incid_sourcesDataTable, HluDataSet.incid_sourcesRow> incid_sourcesTableAdapter
        {
            get { return this._incid_sourcesTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._incid_sourcesTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_bap_quality_determinationDataTable, HluDataSet.lut_bap_quality_determinationRow> lut_bap_quality_determinationTableAdapter
        {
            get { return this._lut_bap_quality_determinationTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_bap_quality_determinationTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_bap_quality_interpretationDataTable, HluDataSet.lut_bap_quality_interpretationRow> lut_bap_quality_interpretationTableAdapter
        {
            get { return this._lut_bap_quality_interpretationTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_bap_quality_interpretationTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_boundary_mapDataTable, HluDataSet.lut_boundary_mapRow> lut_boundary_mapTableAdapter
        {
            get { return this._lut_boundary_mapTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_boundary_mapTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_habitat_classDataTable, HluDataSet.lut_habitat_classRow> lut_habitat_classTableAdapter
        {
            get { return this._lut_habitat_classTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_habitat_classTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_habitat_typeDataTable, HluDataSet.lut_habitat_typeRow> lut_habitat_typeTableAdapter
        {
            get { return this._lut_habitat_typeTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_habitat_typeTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_habitat_type_ihs_habitatDataTable, HluDataSet.lut_habitat_type_ihs_habitatRow> lut_habitat_type_ihs_habitatTableAdapter
        {
            get { return this._lut_habitat_type_ihs_habitatTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_habitat_type_ihs_habitatTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_ihs_complexDataTable, HluDataSet.lut_ihs_complexRow> lut_ihs_complexTableAdapter
        {
            get { return this._lut_ihs_complexTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_ihs_complexTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_ihs_formationDataTable, HluDataSet.lut_ihs_formationRow> lut_ihs_formationTableAdapter
        {
            get { return this._lut_ihs_formationTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_ihs_formationTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_ihs_habitatDataTable, HluDataSet.lut_ihs_habitatRow> lut_ihs_habitatTableAdapter
        {
            get { return this._lut_ihs_habitatTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_ihs_habitatTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_complexDataTable, HluDataSet.lut_ihs_habitat_ihs_complexRow> lut_ihs_habitat_ihs_complexTableAdapter
        {
            get { return this._lut_ihs_habitat_ihs_complexTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_ihs_habitat_ihs_complexTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_formationDataTable, HluDataSet.lut_ihs_habitat_ihs_formationRow> lut_ihs_habitat_ihs_formationTableAdapter
        {
            get { return this._lut_ihs_habitat_ihs_formationTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_ihs_habitat_ihs_formationTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_managementDataTable, HluDataSet.lut_ihs_habitat_ihs_managementRow> lut_ihs_habitat_ihs_managementTableAdapter
        {
            get { return this._lut_ihs_habitat_ihs_managementTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_ihs_habitat_ihs_managementTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_matrixDataTable, HluDataSet.lut_ihs_habitat_ihs_matrixRow> lut_ihs_habitat_ihs_matrixTableAdapter
        {
            get { return this._lut_ihs_habitat_ihs_matrixTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_ihs_habitat_ihs_matrixTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_nvcDataTable, HluDataSet.lut_ihs_habitat_ihs_nvcRow> lut_ihs_habitat_ihs_nvcTableAdapter
        {
            get { return this._lut_ihs_habitat_ihs_nvcTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_ihs_habitat_ihs_nvcTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_ihs_managementDataTable, HluDataSet.lut_ihs_managementRow> lut_ihs_managementTableAdapter
        {
            get { return this._lut_ihs_ManagementTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_ihs_ManagementTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_ihs_matrixDataTable, HluDataSet.lut_ihs_matrixRow> lut_ihs_matrixTableAdapter
        {
            get { return this._lut_ihs_matrixTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_ihs_matrixTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_importanceDataTable, HluDataSet.lut_importanceRow> lut_importanceTableAdapter
        {
            get { return this._lut_importanceTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_importanceTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_last_incidDataTable, HluDataSet.lut_last_incidRow> lut_last_incidTableAdapter
        {
            get { return this._lut_last_incidTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_last_incidTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_legacy_habitatDataTable, HluDataSet.lut_legacy_habitatRow> lut_legacy_habitatTableAdapter
        {
            get { return this._lut_legacy_habitatTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_legacy_habitatTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_operationDataTable, HluDataSet.lut_operationRow> lut_operationTableAdapter
        {
            get { return this._lut_operationTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_operationTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_processDataTable, HluDataSet.lut_processRow> lut_processTableAdapter
        {
            get { return this._lut_processTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_processTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_reasonDataTable, HluDataSet.lut_reasonRow> lut_reasonTableAdapter
        {
            get { return this._lut_reasonTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_reasonTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_site_idDataTable, HluDataSet.lut_site_idRow> lut_site_idTableAdapter
        {
            get { return this._lut_site_idTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_site_idTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_sourcesDataTable, HluDataSet.lut_sourcesRow> lut_sourcesTableAdapter
        {
            get { return this._lut_sourcesTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_sourcesTableAdapter = value;
            }
        }

        public HluTableAdapter<HluDataSet.lut_userDataTable, HluDataSet.lut_userRow> lut_userTableAdapter
        {
            get { return this._lut_userTableAdapter; }
            set
            {
                if (!this.MatchTableAdapterConnection(value.Connection))
                    throw new ArgumentException(_sameConnErrorMsg);
                this._lut_userTableAdapter = value;
            }
        }

        #endregion

        public UpdateOrderOption UpdateOrder
        {
            get { return this._updateOrder; }
            set { this._updateOrder = value; }
        }

        public bool BackupDataSetBeforeUpdate
        {
            get { return this._backupDataSetBeforeUpdate; }
            set { this._backupDataSetBeforeUpdate = value; }
        }

        [Browsable(false)]
        public IDbConnection Connection { get { return _db.Connection; } }

        [Browsable(false)]
        public int TableAdapterInstanceCount
        {
            get
            {
                int count = 0;

                if (this._exportsTableAdapter != null) count++;

                if (this._exportsFieldsTableAdapter != null) count++;

                if (this._historyTableAdapter != null) count++;

                if (this._incidTableAdapter != null) count++;

                if (this._incid_bapTableAdapter != null) count++;

                if (this._incid_ihs_complexTableAdapter != null) count++;

                if (this._incid_ihs_formationTableAdapter != null) count++;

                if (this._incid_ihs_managementTableAdapter != null) count++;

                if (this._incid_ihs_matrixTableAdapter != null) count++;

                if (this._incid_mm_polygonsTableAdapter != null) count++;

                if (this._incid_sourcesTableAdapter != null) count++;

                if (this._lut_bap_quality_determinationTableAdapter != null) count++;

                if (this._lut_bap_quality_interpretationTableAdapter != null) count++;

                if (this._lut_boundary_mapTableAdapter != null) count++;

                if (this._lut_habitat_classTableAdapter != null) count++;

                if (this._lut_habitat_typeTableAdapter != null) count++;

                if (this._lut_habitat_type_ihs_habitatTableAdapter != null) count++;

                if (this._lut_ihs_complexTableAdapter != null) count++;

                if (this._lut_ihs_formationTableAdapter != null) count++;

                if (this._lut_ihs_habitatTableAdapter != null) count++;

                if (this._lut_ihs_habitat_ihs_complexTableAdapter != null) count++;

                if (this._lut_ihs_habitat_ihs_formationTableAdapter != null) count++;

                if (this._lut_ihs_habitat_ihs_managementTableAdapter != null) count++;

                if (this._lut_ihs_habitat_ihs_matrixTableAdapter != null) count++;

                if (this._lut_ihs_habitat_ihs_nvcTableAdapter != null) count++;

                if (this._lut_ihs_ManagementTableAdapter != null) count++;

                if (this._lut_ihs_matrixTableAdapter != null) count++;

                if (this._lut_importanceTableAdapter != null) count++;

                if (this._lut_last_incidTableAdapter != null) count++;

                if (this._lut_operationTableAdapter != null) count++;

                if (this._lut_processTableAdapter != null) count++;

                if (this._lut_reasonTableAdapter != null) count++;

                if (this._lut_site_idTableAdapter != null) count++;

                if (this._lut_sourcesTableAdapter != null) count++;

                if (this._lut_userTableAdapter != null) count++;

                if (this._lut_versionTableAdapter != null) count++;

                return count;
            }
        }

        #endregion

        #region Constructor

        internal TableAdapterManager(DbBase db, Scope createAdapters)
        {
            if ((db == null) || (db.Connection == null)) throw new ArgumentNullException("db");

            string errorMessage;
            if (!IsHluDataSet(db, out errorMessage)) throw new ArgumentException("db", errorMessage);

            _db = db;

            _tableAdapterMatches = (from pt in typeof(HluDataSet).GetProperties()
                                    from pa in this.GetType().GetProperties().Where(pi => pi.PropertyType.GetGenericArguments().Count() > 0)
                                    where pa.PropertyType.GetGenericArguments().Contains(pt.PropertyType)
                                    select new 
                                    { 
                                        key = pt.PropertyType, 
                                        value = pa
                                    }
                                    ).ToDictionary(kv => kv.key, kv => kv.value);

            try
            {
                switch (createAdapters)
                {
                    case Scope.DataNoMMPolygonsHistory:
                        CreateAdaptersData();
                        break;
                    case Scope.Data:
                        CreateAdaptersData();
                        CreateAdapterMMPolygons();
                        break;
                    case Scope.Lookup:
                        CreateAdaptersLut();
                        break;
                    case Scope.AllButMMPolygonsHistory:
                        CreateAdaptersData();
                        CreateAdaptersLut();
                        break;
                    case Scope.All:
                        CreateAdaptersData();
                        CreateAdapterMMPolygons();
                        CreateAdaptersLut();
                        break;
                }
            }
            catch { throw; }
        }

        internal static bool IsHluDataSet(DbBase db, out string errorMessage)
        {
            errorMessage = null;

            try
            {
                DataTable schemaTable = db.GetSchema("Columns", db.RestrictionNameSchema, db.DefaultSchema,
                    (DbConnection)db.Connection, (DbTransaction)db.Transaction);
                
                if (schemaTable != null)
                {
                    var cols = from t in TableAdapterManager.DataTableTypes
                               from c in t.GetProperties().Where(pi => pi.PropertyType == typeof(DataColumn))
                               select new
                               {
                                   Table = t.Name.Remove(t.Name.Length - 9),
                                   Column = c.Name.Remove(c.Name.Length - 6)
                               };

                    string[][] missingSchemaElems = (from c in cols
                                                     let schema = from r in schemaTable.AsEnumerable()
                                                                  select new
                                                                  {
                                                                      Table = r.Field<string>("TABLE_NAME"),
                                                                      Column = r.Field<string>("COLUMN_NAME")
                                                                  }
                                                     where schema.Count(s => s.Column == c.Column && s.Table == c.Table) == 0
                                                     select new string[] { db.QuoteIdentifier(c.Table), db.QuoteIdentifier(c.Column) }
                                                     ).ToArray();

                    if (missingSchemaElems.Length > 0)
                    {
                        StringBuilder messageText = new StringBuilder();
                        int i = 0;
                        while (i < missingSchemaElems.Length)
                        {
                            string table = missingSchemaElems[i][0];
                            messageText.Append("\n\nTable: ").Append(table);
                            StringBuilder columnList = new StringBuilder();
                            while ((i < missingSchemaElems.Length) && (missingSchemaElems[i][0] == table))
                            {
                                columnList.Append(", ").Append(missingSchemaElems[i++][1]);
                            }
                            if (columnList.Length > 0)
                                messageText.Append(columnList.Remove(0, 1).Insert(0, "\nColumns:"));
                        }
                        errorMessage = String.Format("Connection does not point to a valid HLU database." +
                            "\nBad schema objects: {0}", messageText);
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    errorMessage = "Failed to get schema information from database.";
                }
            }
            catch (Exception ex) { errorMessage = ex.Message; }
            
            return false;
        }

        private void CreateAdaptersData()
        {
            try
            {
                _historyTableAdapter = new HluTableAdapter<HluDataSet.historyDataTable, HluDataSet.historyRow>(_db);
                _incidTableAdapter = new HluTableAdapter<HluDataSet.incidDataTable, HluDataSet.incidRow>(_db);
                _incid_bapTableAdapter = new HluTableAdapter<HluDataSet.incid_bapDataTable, HluDataSet.incid_bapRow>(_db);
                _incid_ihs_complexTableAdapter = new HluTableAdapter<HluDataSet.incid_ihs_complexDataTable, HluDataSet.incid_ihs_complexRow>(_db);
                _incid_ihs_formationTableAdapter = new HluTableAdapter<HluDataSet.incid_ihs_formationDataTable, HluDataSet.incid_ihs_formationRow>(_db);
                _incid_ihs_managementTableAdapter = new HluTableAdapter<HluDataSet.incid_ihs_managementDataTable, HluDataSet.incid_ihs_managementRow>(_db);
                _incid_ihs_matrixTableAdapter = new HluTableAdapter<HluDataSet.incid_ihs_matrixDataTable, HluDataSet.incid_ihs_matrixRow>(_db);
                _incid_sourcesTableAdapter = new HluTableAdapter<HluDataSet.incid_sourcesDataTable, HluDataSet.incid_sourcesRow>(_db);
            }
            catch { throw; }
        }

        private void CreateAdapterMMPolygons()
        {
            _incid_mm_polygonsTableAdapter = new HluTableAdapter<HluDataSet.incid_mm_polygonsDataTable, HluDataSet.incid_mm_polygonsRow>(_db);
        }

        private void CreateAdaptersLut()
        {
            try
            {
                _exportsTableAdapter = new HluTableAdapter<HluDataSet.exportsDataTable, HluDataSet.exportsRow>(_db);
                _exportsFieldsTableAdapter = new HluTableAdapter<HluDataSet.exports_fieldsDataTable, HluDataSet.exports_fieldsRow>(_db);
                _lut_bap_quality_determinationTableAdapter = new HluTableAdapter<HluDataSet.lut_bap_quality_determinationDataTable, HluDataSet.lut_bap_quality_determinationRow>(_db);
                _lut_bap_quality_interpretationTableAdapter = new HluTableAdapter<HluDataSet.lut_bap_quality_interpretationDataTable, HluDataSet.lut_bap_quality_interpretationRow>(_db);
                _lut_boundary_mapTableAdapter = new HluTableAdapter<HluDataSet.lut_boundary_mapDataTable, HluDataSet.lut_boundary_mapRow>(_db);
                _lut_habitat_classTableAdapter = new HluTableAdapter<HluDataSet.lut_habitat_classDataTable, HluDataSet.lut_habitat_classRow>(_db);
                _lut_habitat_typeTableAdapter = new HluTableAdapter<HluDataSet.lut_habitat_typeDataTable, HluDataSet.lut_habitat_typeRow>(_db);
                _lut_habitat_type_ihs_habitatTableAdapter = new HluTableAdapter<HluDataSet.lut_habitat_type_ihs_habitatDataTable, HluDataSet.lut_habitat_type_ihs_habitatRow>(_db);
                _lut_ihs_complexTableAdapter = new HluTableAdapter<HluDataSet.lut_ihs_complexDataTable, HluDataSet.lut_ihs_complexRow>(_db);
                _lut_ihs_formationTableAdapter = new HluTableAdapter<HluDataSet.lut_ihs_formationDataTable, HluDataSet.lut_ihs_formationRow>(_db);
                _lut_ihs_habitatTableAdapter = new HluTableAdapter<HluDataSet.lut_ihs_habitatDataTable, HluDataSet.lut_ihs_habitatRow>(_db);
                _lut_ihs_habitat_ihs_complexTableAdapter = new HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_complexDataTable, HluDataSet.lut_ihs_habitat_ihs_complexRow>(_db);
                _lut_ihs_habitat_ihs_formationTableAdapter = new HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_formationDataTable, HluDataSet.lut_ihs_habitat_ihs_formationRow>(_db);
                _lut_ihs_habitat_ihs_managementTableAdapter = new HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_managementDataTable, HluDataSet.lut_ihs_habitat_ihs_managementRow>(_db);
                _lut_ihs_habitat_ihs_matrixTableAdapter = new HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_matrixDataTable, HluDataSet.lut_ihs_habitat_ihs_matrixRow>(_db);
                _lut_ihs_habitat_ihs_nvcTableAdapter = new HluTableAdapter<HluDataSet.lut_ihs_habitat_ihs_nvcDataTable, HluDataSet.lut_ihs_habitat_ihs_nvcRow>(_db);
                _lut_ihs_ManagementTableAdapter = new HluTableAdapter<HluDataSet.lut_ihs_managementDataTable, HluDataSet.lut_ihs_managementRow>(_db);
                _lut_ihs_matrixTableAdapter = new HluTableAdapter<HluDataSet.lut_ihs_matrixDataTable, HluDataSet.lut_ihs_matrixRow>(_db);
                _lut_importanceTableAdapter = new HluTableAdapter<HluDataSet.lut_importanceDataTable, HluDataSet.lut_importanceRow>(_db);
                _lut_last_incidTableAdapter = new HluTableAdapter<HluDataSet.lut_last_incidDataTable, HluDataSet.lut_last_incidRow>(_db);
                _lut_operationTableAdapter = new HluTableAdapter<HluDataSet.lut_operationDataTable, HluDataSet.lut_operationRow>(_db);
                _lut_processTableAdapter = new HluTableAdapter<HluDataSet.lut_processDataTable, HluDataSet.lut_processRow>(_db);
                _lut_reasonTableAdapter = new HluTableAdapter<HluDataSet.lut_reasonDataTable, HluDataSet.lut_reasonRow>(_db);
                _lut_site_idTableAdapter = new HluTableAdapter<HluDataSet.lut_site_idDataTable, HluDataSet.lut_site_idRow>(_db);
                _lut_sourcesTableAdapter = new HluTableAdapter<HluDataSet.lut_sourcesDataTable, HluDataSet.lut_sourcesRow>(_db);
                _lut_userTableAdapter = new HluTableAdapter<HluDataSet.lut_userDataTable, HluDataSet.lut_userRow>(_db);
                _lut_versionTableAdapter = new HluTableAdapter<HluDataSet.lut_versionDataTable, HluDataSet.lut_versionRow>(_db);
            }
            catch { throw; }
        }

        #endregion

        #region Fill

        public void Fill(HluDataSet hluDS, Scope fillTables, bool clearBeforeFill)
        {
            if (hluDS == null) throw new ArgumentException("hluDS");

            try
            {
                switch (fillTables)
                {
                    case Scope.DataNoMMPolygonsHistory:
                        Fill(hluDS, DataTableTypes, clearBeforeFill);
                        break;
                    case Scope.Data:
                        Fill(hluDS, DataTableTypes, clearBeforeFill);
                        Fill(hluDS, typeof(HluDataSet.incid_mm_polygonsDataTable), clearBeforeFill);
                        Fill(hluDS, typeof(HluDataSet.historyDataTable), clearBeforeFill);
                        break;
                    case Scope.Lookup:
                        Fill(hluDS, LookupTableTypes, clearBeforeFill);
                        break;
                    case Scope.AllButMMPolygonsHistory:
                        Fill(hluDS, DataTableTypes, clearBeforeFill);
                        Fill(hluDS, LookupTableTypes, clearBeforeFill);
                        break;
                    case Scope.All:
                        Fill(hluDS, LookupTableTypes, clearBeforeFill);
                        Fill(hluDS, DataTableTypes, clearBeforeFill);
                        Fill(hluDS, typeof(HluDataSet.incid_mm_polygonsDataTable), clearBeforeFill);
                        Fill(hluDS, typeof(HluDataSet.historyDataTable), clearBeforeFill);
                        break;
                }
            }
            catch { throw; }
        }

        public void Fill(HluDataSet hluDS, Type tableType, bool clearBeforeFill)
        {
            Fill(hluDS, new Type[] { tableType }, clearBeforeFill);
        }

        public void Fill(HluDataSet hluDS, Type tableType, string whereClause, bool clearBeforeFill)
        {
            Fill(hluDS, new Type[] { tableType }, new string[] { whereClause }, clearBeforeFill);
        }

        public void Fill(HluDataSet hluDS, Type[] tableTypes,  bool clearBeforeFill)
        {
            try
            {
                foreach (Type tableType in tableTypes)
                {
                    PropertyInfo adapterPropertyInfo;
                    PropertyInfo tablePropertyInfo;
                    object adapterProperty;
                    
                    if (FillShared(hluDS, tableType, clearBeforeFill, out adapterPropertyInfo,
                        out tablePropertyInfo, out adapterProperty))
                    {
                        MethodInfo fillMethodInfo = 
                            adapterPropertyInfo.PropertyType.GetMethod("Fill", new Type[] { tableType });

                        fillMethodInfo.Invoke(adapterProperty, new object[] { 
                            hluDS.GetType().InvokeMember(tablePropertyInfo.Name, 
                            BindingFlags.GetProperty, null, hluDS, null) });
                    }
                    else
                    {
                        throw new ArgumentException("table", 
                            String.Format("Table '{0}' is not a member of HluDataSet.", tableType.Name));
                    }
                }
            }
            catch { throw; }
        }

        public void Fill(HluDataSet hluDS, Type[] tableTypes, string[] whereClauses, bool clearBeforeFill)
        {
            try
            {
                for (int i = 0; i < tableTypes.Length; i++)
                {
                    Type tableType = tableTypes[i];

                    PropertyInfo adapterPropertyInfo;
                    PropertyInfo tablePropertyInfo;
                    object adapterProperty;

                    if (FillShared(hluDS, tableType, clearBeforeFill, out adapterPropertyInfo,
                        out tablePropertyInfo, out adapterProperty))
                    {
                        MethodInfo fillMethodInfo = adapterPropertyInfo.PropertyType.GetMethod("Fill",
                            new Type[] { tableType, typeof(string) });

                        fillMethodInfo.Invoke(adapterProperty, new object[] { 
                            hluDS.GetType().InvokeMember(tablePropertyInfo.Name, 
                            BindingFlags.GetProperty, null, hluDS, null), whereClauses[i] });
                    }
                    else
                    {
                        throw new ArgumentException("table",
                            String.Format("Table '{0}' is not a member of HluDataSet.", tableType.Name));
                    }
                }
            }
            catch { throw; }
        }

        private bool FillShared(HluDataSet hluDS, Type tableType, bool clearBeforeFill,
            out PropertyInfo adapterPropertyInfo, out PropertyInfo tablePropertyInfo, out object adapterProperty)
        {
            tablePropertyInfo = null;
            adapterProperty = null;

            if (_tableAdapterMatches.TryGetValue(tableType, out adapterPropertyInfo))
            {
                tablePropertyInfo = typeof(HluDataSet).GetProperties().Single(pi => pi.PropertyType == tableType);

                adapterProperty = this.GetType().InvokeMember(adapterPropertyInfo.Name,
                     BindingFlags.GetProperty, null, this, null);

                PropertyInfo clearBeforeFillPropertyInfo = adapterProperty.GetType().GetProperty("ClearBeforeFill");

                bool previousClearBeforeFill = (bool)clearBeforeFillPropertyInfo.GetValue(adapterProperty, null);

                adapterProperty.GetType().GetProperty("ClearBeforeFill")
                    .SetValue(adapterProperty, clearBeforeFill, null);
                return true;
            }
            else
            {
                return false;
            }
        }

        private void FillOneTable<T, R>(HluTableAdapter<T, R> adapter, T table, bool clearBeforeFill)
            where T : DataTable, new()
            where R : DataRow
        {
            if ((adapter == null) || (table == null)) return;

            bool previousClearBeforeFill = adapter.ClearBeforeFill;

            try
            {
                adapter.ClearBeforeFill = clearBeforeFill;
                adapter.Fill(table);
            }
            catch
            {
                adapter.ClearBeforeFill = previousClearBeforeFill;
                throw;
            }
        }

        #endregion

        #region Update

        /// <summary>
        /// Prepares one adapter for update operation.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="R"></typeparam>
        /// <param name="tableAdapter"></param>
        /// <param name="adaptersWithAcceptChangesDuringUpdate"></param>
        private void PrepareUpdate<T, R>(HluTableAdapter<T, R> tableAdapter,
            ref List<DataAdapter> adaptersWithAcceptChangesDuringUpdate)
            where T : DataTable, new()
            where R : DataRow
        {
            if ((tableAdapter != null))
            {
                DataAdapter dataAdapter = tableAdapter as DataAdapter;
                if (dataAdapter.AcceptChangesDuringUpdate)
                {
                    dataAdapter.AcceptChangesDuringUpdate = false;
                    adaptersWithAcceptChangesDuringUpdate.Add(dataAdapter);
                }
            }
        }

        /// <summary>
        /// Perfom Update for one adapter.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="adapter"></param>
        /// <param name="table"></param>
        /// <param name="allChangedRows"></param>
        /// <param name="allAddedRows"></param>
        /// <param name="result"></param>
        private void UpdateUpdatedRows<T, R>(HluTableAdapter<T, R> adapter, T table,
            List<DataRow> allChangedRows, List<DataRow> allAddedRows, ref int result)
            where T : DataTable, new()
            where R : DataRow
        {
            if ((adapter != null))
            {
                R[] updatedRows = (R[])table.Select(null, null, DataViewRowState.ModifiedCurrent);
                updatedRows = this.GetRealUpdatedRows(updatedRows, allAddedRows);
                int affected;
                if ((updatedRows != null) && (0 < updatedRows.Length))
                {
                    if ((affected = adapter.Update(updatedRows)) != -1)
                        result += affected;
                    allChangedRows.AddRange(updatedRows);
                }
            }
        }

        /// <summary>
        /// Perfom Insert for one adapter.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="adapter"></param>
        /// <param name="table"></param>
        /// <param name="allAddedRows"></param>
        /// <param name="result"></param>
        private void UpdateInsertedRows<T, R>(HluTableAdapter<T, R> adapter, T table,
            List<DataRow> allAddedRows, ref int result)
            where T : DataTable, new()
            where R : DataRow
        {
            if (adapter != null)
            {
                R[] addedRows = (R[])table.Select(null, null, DataViewRowState.Added);
                int affected;
                if ((addedRows != null) && (0 < addedRows.Length))
                {
                    if ((affected = adapter.Update(addedRows)) != -1)
                        result += affected;
                    allAddedRows.AddRange(addedRows);
                }
            }
        }

        /// <summary>
        /// Perfom Delete for one adapter.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="adapter"></param>
        /// <param name="table"></param>
        /// <param name="allChangedRows"></param>
        /// <param name="result"></param>
        private void UpdateDeletedRows<T, R>(HluTableAdapter<T, R> adapter, T table,
            List<DataRow> allChangedRows, ref int result)
            where T : DataTable, new()
            where R : DataRow
        {
            if (adapter != null)
            {
                R[] deletedRows = (R[])table.Select(null, null, DataViewRowState.Deleted);
                int affected;
                if ((deletedRows != null) && (0 < deletedRows.Length))
                {
                    if ((affected = adapter.Update(deletedRows)) != -1)
                        result += affected;
                    allChangedRows.AddRange(deletedRows);
                }
            }
        }

        /// <summary>
        ///Update rows in top-down order.
        ///</summary>
        private int UpdateUpdatedRows(HluDataSet dataSet, List<DataRow> allChangedRows, List<DataRow> allAddedRows)
        {
            int result = 0;

            UpdateUpdatedRows(_lut_habitat_classTableAdapter, dataSet.lut_habitat_class,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_ihs_habitatTableAdapter, dataSet.lut_ihs_habitat,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_userTableAdapter, dataSet.lut_user,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_boundary_mapTableAdapter, dataSet.lut_boundary_map,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_habitat_typeTableAdapter, dataSet.lut_habitat_type,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_habitat_type_ihs_habitatTableAdapter, dataSet.lut_habitat_type_ihs_habitat,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_reasonTableAdapter, dataSet.lut_reason,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_ihs_complexTableAdapter, dataSet.lut_ihs_complex,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_sourcesTableAdapter, dataSet.lut_sources,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_ihs_matrixTableAdapter, dataSet.lut_ihs_matrix,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_exportsTableAdapter, dataSet.exports,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_operationTableAdapter, dataSet.lut_operation,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_importanceTableAdapter, dataSet.lut_importance,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_ihs_ManagementTableAdapter, dataSet.lut_ihs_management,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_ihs_formationTableAdapter, dataSet.lut_ihs_formation,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_processTableAdapter, dataSet.lut_process,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_incid_mm_polygonsTableAdapter, dataSet.incid_mm_polygons,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_incidTableAdapter, dataSet.incid,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_bap_quality_determinationTableAdapter, dataSet.lut_bap_quality_determination,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_bap_quality_interpretationTableAdapter, dataSet.lut_bap_quality_interpretation,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_incid_bapTableAdapter, dataSet.incid_bap,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_incid_ihs_complexTableAdapter, dataSet.incid_ihs_complex,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_last_incidTableAdapter, dataSet.lut_last_incid,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_versionTableAdapter, dataSet.lut_version,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_historyTableAdapter, dataSet.history,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_site_idTableAdapter, dataSet.lut_site_id,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_ihs_habitat_ihs_complexTableAdapter, dataSet.lut_ihs_habitat_ihs_complex,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_ihs_habitat_ihs_formationTableAdapter, dataSet.lut_ihs_habitat_ihs_formation,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_incid_sourcesTableAdapter, dataSet.incid_sources,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_incid_ihs_matrixTableAdapter, dataSet.incid_ihs_matrix,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_ihs_habitat_ihs_managementTableAdapter, dataSet.lut_ihs_habitat_ihs_management,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_incid_ihs_managementTableAdapter, dataSet.incid_ihs_management,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_incid_ihs_formationTableAdapter, dataSet.incid_ihs_formation,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_ihs_habitat_ihs_matrixTableAdapter, dataSet.lut_ihs_habitat_ihs_matrix,
                allChangedRows, allAddedRows, ref result);

            UpdateUpdatedRows(_lut_ihs_habitat_ihs_nvcTableAdapter, dataSet.lut_ihs_habitat_ihs_nvc,
                allChangedRows, allAddedRows, ref result);

            return result;
        }

        /// <summary>
        ///Insert rows in top-down order.
        ///</summary>
        private int UpdateInsertedRows(HluDataSet dataSet, List<DataRow> allAddedRows)
        {
            int result = 0;

            UpdateInsertedRows(_lut_habitat_classTableAdapter, dataSet.lut_habitat_class,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_ihs_habitatTableAdapter, dataSet.lut_ihs_habitat,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_userTableAdapter, dataSet.lut_user,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_boundary_mapTableAdapter, dataSet.lut_boundary_map,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_habitat_typeTableAdapter, dataSet.lut_habitat_type,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_habitat_type_ihs_habitatTableAdapter, dataSet.lut_habitat_type_ihs_habitat,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_reasonTableAdapter, dataSet.lut_reason,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_ihs_complexTableAdapter, dataSet.lut_ihs_complex,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_sourcesTableAdapter, dataSet.lut_sources,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_ihs_matrixTableAdapter, dataSet.lut_ihs_matrix,
                allAddedRows, ref result);

            UpdateInsertedRows(_exportsTableAdapter, dataSet.exports,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_operationTableAdapter, dataSet.lut_operation,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_importanceTableAdapter, dataSet.lut_importance,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_ihs_ManagementTableAdapter, dataSet.lut_ihs_management,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_ihs_formationTableAdapter, dataSet.lut_ihs_formation,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_processTableAdapter, dataSet.lut_process,
                allAddedRows, ref result);

            UpdateInsertedRows(_incid_mm_polygonsTableAdapter, dataSet.incid_mm_polygons,
                allAddedRows, ref result);

            UpdateInsertedRows(_incidTableAdapter, dataSet.incid,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_bap_quality_determinationTableAdapter, dataSet.lut_bap_quality_determination,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_bap_quality_interpretationTableAdapter, dataSet.lut_bap_quality_interpretation,
                allAddedRows, ref result);

            UpdateInsertedRows(_incid_bapTableAdapter, dataSet.incid_bap,
                allAddedRows, ref result);

            UpdateInsertedRows(_incid_ihs_complexTableAdapter, dataSet.incid_ihs_complex,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_last_incidTableAdapter, dataSet.lut_last_incid,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_versionTableAdapter, dataSet.lut_version,
                allAddedRows, ref result);

            UpdateInsertedRows(_historyTableAdapter, dataSet.history,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_site_idTableAdapter, dataSet.lut_site_id,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_ihs_habitat_ihs_complexTableAdapter, dataSet.lut_ihs_habitat_ihs_complex,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_ihs_habitat_ihs_formationTableAdapter, dataSet.lut_ihs_habitat_ihs_formation,
                allAddedRows, ref result);

            UpdateInsertedRows(_incid_sourcesTableAdapter, dataSet.incid_sources,
                allAddedRows, ref result);

            UpdateInsertedRows(_incid_ihs_matrixTableAdapter, dataSet.incid_ihs_matrix,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_ihs_habitat_ihs_managementTableAdapter, dataSet.lut_ihs_habitat_ihs_management,
                allAddedRows, ref result);

            UpdateInsertedRows(_incid_ihs_managementTableAdapter, dataSet.incid_ihs_management,
                allAddedRows, ref result);

            UpdateInsertedRows(_incid_ihs_formationTableAdapter, dataSet.incid_ihs_formation,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_ihs_habitat_ihs_matrixTableAdapter, dataSet.lut_ihs_habitat_ihs_matrix,
                allAddedRows, ref result);

            UpdateInsertedRows(_lut_ihs_habitat_ihs_nvcTableAdapter, dataSet.lut_ihs_habitat_ihs_nvc,
                allAddedRows, ref result);

            return result;
        }

        /// <summary>
        ///Delete rows in bottom-up order.
        ///</summary>
        private int UpdateDeletedRows(HluDataSet dataSet, List<DataRow> allChangedRows)
        {
            int result = 0;

            UpdateDeletedRows(_lut_ihs_habitat_ihs_nvcTableAdapter, dataSet.lut_ihs_habitat_ihs_nvc,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_ihs_habitat_ihs_matrixTableAdapter, dataSet.lut_ihs_habitat_ihs_matrix,
                allChangedRows, ref result);

            UpdateDeletedRows(_incid_ihs_formationTableAdapter, dataSet.incid_ihs_formation,
                allChangedRows, ref result);

            UpdateDeletedRows(_incid_ihs_managementTableAdapter, dataSet.incid_ihs_management,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_ihs_habitat_ihs_managementTableAdapter, dataSet.lut_ihs_habitat_ihs_management,
                allChangedRows, ref result);

            UpdateDeletedRows(_incid_ihs_matrixTableAdapter, dataSet.incid_ihs_matrix,
                allChangedRows, ref result);

            UpdateDeletedRows(_incid_sourcesTableAdapter, dataSet.incid_sources,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_ihs_habitat_ihs_formationTableAdapter, dataSet.lut_ihs_habitat_ihs_formation,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_ihs_habitat_ihs_complexTableAdapter, dataSet.lut_ihs_habitat_ihs_complex,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_site_idTableAdapter, dataSet.lut_site_id,
                allChangedRows, ref result);

            UpdateDeletedRows(_historyTableAdapter, dataSet.history,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_versionTableAdapter, dataSet.lut_version,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_last_incidTableAdapter, dataSet.lut_last_incid,
                allChangedRows, ref result);

            UpdateDeletedRows(_incid_ihs_complexTableAdapter, dataSet.incid_ihs_complex,
                allChangedRows, ref result);

            UpdateDeletedRows(_incid_bapTableAdapter, dataSet.incid_bap,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_bap_quality_interpretationTableAdapter, dataSet.lut_bap_quality_interpretation,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_bap_quality_determinationTableAdapter, dataSet.lut_bap_quality_determination,
                allChangedRows, ref result);

            UpdateDeletedRows(_incidTableAdapter, dataSet.incid,
                allChangedRows, ref result);

            UpdateDeletedRows(_incid_mm_polygonsTableAdapter, dataSet.incid_mm_polygons,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_processTableAdapter, dataSet.lut_process,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_ihs_formationTableAdapter, dataSet.lut_ihs_formation,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_ihs_ManagementTableAdapter, dataSet.lut_ihs_management,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_importanceTableAdapter, dataSet.lut_importance,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_operationTableAdapter, dataSet.lut_operation,
                allChangedRows, ref result);

            UpdateDeletedRows(_exportsTableAdapter, dataSet.exports,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_ihs_matrixTableAdapter, dataSet.lut_ihs_matrix,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_sourcesTableAdapter, dataSet.lut_sources,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_ihs_complexTableAdapter, dataSet.lut_ihs_complex,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_reasonTableAdapter, dataSet.lut_reason,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_habitat_type_ihs_habitatTableAdapter, dataSet.lut_habitat_type_ihs_habitat,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_habitat_typeTableAdapter, dataSet.lut_habitat_type,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_boundary_mapTableAdapter, dataSet.lut_boundary_map,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_userTableAdapter, dataSet.lut_user,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_ihs_habitatTableAdapter, dataSet.lut_ihs_habitat,
                allChangedRows, ref result);

            UpdateDeletedRows(_lut_habitat_classTableAdapter, dataSet.lut_habitat_class,
                allChangedRows, ref result);

            return result;
        }

        /// <summary>
        ///Remove inserted rows that become updated rows after calling TableAdapter.Update(inserted rows) first
        ///</summary>
        private R[] GetRealUpdatedRows<R>(R[] updatedRows, List<DataRow> allAddedRows)
            where R : DataRow
        {
            if ((updatedRows == null) || (updatedRows.Length < 1)) return updatedRows;

            if ((allAddedRows == null) || (allAddedRows.Count < 1)) return updatedRows;

            List<R> realUpdatedRows = new List<R>();
            for (int i = 0; (i < updatedRows.Length); i = (i + 1))
            {
                R row = updatedRows[i];
                if ((allAddedRows.Contains(row) == false)) realUpdatedRows.Add(row);
            }
            return realUpdatedRows.ToArray();
        }

        /// <summary>
        ///Update all changes to the dataset.
        ///</summary>
        public virtual int UpdateAll(HluDataSet dataSet)
        {
            if (dataSet == null) throw new ArgumentNullException("dataSet");

            if (!dataSet.HasChanges()) return 0;

            IDbConnection workConnection = this.Connection;
            if ((workConnection == null))
                throw new ApplicationException("TableAdapterManager contains no connection information." +
                    "Set each TableAdapterManager TableAdapter property to a valid TableAdapter instance.");

            bool workConnOpened = false;
            if (((workConnection.State & ConnectionState.Broken) == ConnectionState.Broken))
            {
                workConnection.Close();
            }

            if ((workConnection.State == ConnectionState.Closed))
            {
                workConnection.Open();
                workConnOpened = true;
            }

            if (!_db.BeginTransaction(true, IsolationLevel.ReadCommitted))
                throw new ApplicationException("The transaction cannot begin. The current data connection does not " +
                    "support transactions or the current state is not allowing the transaction to begin.");

            List<DataRow> allChangedRows = new List<DataRow>();
            List<DataRow> allAddedRows = new List<DataRow>();
            List<DataAdapter> adaptersWithAcceptChangesDuringUpdate = new List<DataAdapter>();
            int result = 0;
            DataSet backupDataSet = null;
            if (this.BackupDataSetBeforeUpdate)
            {
                backupDataSet = new DataSet();
                backupDataSet.Merge(dataSet);
            }

            try
            {
                // ---- Prepare for update -----------
                //
                PrepareUpdate(this._exportsTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._historyTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._incidTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._incid_bapTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._incid_ihs_complexTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._incid_ihs_formationTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._incid_ihs_managementTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._incid_ihs_matrixTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._incid_mm_polygonsTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._incid_sourcesTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_bap_quality_determinationTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_bap_quality_interpretationTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_boundary_mapTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_habitat_classTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_habitat_typeTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_habitat_type_ihs_habitatTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_ihs_complexTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_ihs_formationTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_ihs_habitatTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_ihs_habitat_ihs_complexTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_ihs_habitat_ihs_formationTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_ihs_habitat_ihs_managementTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_ihs_habitat_ihs_matrixTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_ihs_habitat_ihs_nvcTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_ihs_ManagementTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_ihs_matrixTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_importanceTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_last_incidTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_operationTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_processTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_reasonTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_site_idTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_sourcesTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_userTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);
                PrepareUpdate(this._lut_versionTableAdapter, ref adaptersWithAcceptChangesDuringUpdate);

                // 
                //---- Perform updates -----------
                //
                if ((this.UpdateOrder == UpdateOrderOption.UpdateInsertDelete))
                {
                    result = (result + this.UpdateUpdatedRows(dataSet, allChangedRows, allAddedRows));
                    result = (result + this.UpdateInsertedRows(dataSet, allAddedRows));
                }
                else
                {
                    result = (result + this.UpdateInsertedRows(dataSet, allAddedRows));
                    result = (result + this.UpdateUpdatedRows(dataSet, allChangedRows, allAddedRows));
                }
                result = (result + this.UpdateDeletedRows(dataSet, allChangedRows));

                // 
                //---- Commit updates -----------
                //
                _db.CommitTransaction();

                if ((0 < allAddedRows.Count))
                {
                    DataRow[] rows = new System.Data.DataRow[allAddedRows.Count];
                    allAddedRows.CopyTo(rows);
                    for (int i = 0; i < rows.Length; i++)
                    {
                        DataRow row = rows[i];
                        row.AcceptChanges();
                    }
                }

                if ((0 < allChangedRows.Count))
                {
                    DataRow[] rows = new System.Data.DataRow[allChangedRows.Count];
                    allChangedRows.CopyTo(rows);
                    for (int i = 0; i < rows.Length; i++)
                    {
                        DataRow row = rows[i];
                        row.AcceptChanges();
                    }
                }
            }
            catch (Exception ex)
            {
                _db.RollbackTransaction();
                // ---- Restore the dataset -----------
                if (this.BackupDataSetBeforeUpdate)
                {
                    System.Diagnostics.Debug.Assert((backupDataSet != null));
                    dataSet.Clear();
                    dataSet.Merge(backupDataSet);
                }
                else
                {
                    if ((0 < allAddedRows.Count))
                    {
                        DataRow[] rows = new System.Data.DataRow[allAddedRows.Count];
                        allAddedRows.CopyTo(rows);
                        for (int i = 0; i < rows.Length; i++)
                        {
                            DataRow row = rows[i];
                            row.AcceptChanges();
                            row.SetAdded();
                        }
                    }
                }
                throw ex;
            }
            finally
            {
                if (workConnOpened) workConnection.Close();

                if ((0 < adaptersWithAcceptChangesDuringUpdate.Count))
                {
                    DataAdapter[] adapters = new DataAdapter[adaptersWithAcceptChangesDuringUpdate.Count];
                    adaptersWithAcceptChangesDuringUpdate.CopyTo(adapters);
                    for (int i = 0; i < adapters.Length; i++)
                    {
                        DataAdapter adapter = adapters[i];
                        adapter.AcceptChangesDuringUpdate = true;
                    }
                }
            }
            return result;
        }

        #endregion

        protected virtual void SortSelfReferenceRows<R>(R[] rows, DataRelation relation, bool childFirst)
            where R : DataRow
        {
            Array.Sort<R>(rows, new SelfReferenceComparer<R>(relation, childFirst));
        }

        private bool MatchTableAdapterConnection(DbBase inputDb)
        {
            if (inputDb == null)
                return false;
            else
                return inputDb.Equals(_db);
        }

        protected virtual bool MatchTableAdapterConnection(IDbConnection inputConnection)
        {
            if (inputConnection == null) // this.Connection is never null
                return false;
            else if (String.Equals(this.Connection.ConnectionString, inputConnection.ConnectionString,
                StringComparison.Ordinal)) return true;
            else
                return false;
        }

        public enum Scope
        {
            None = 0,
            DataNoMMPolygonsHistory = 1,
            Data = 2,
            Lookup = 3,
            AllButMMPolygonsHistory = 4,
            All = 5
        }

        /// <summary>
        ///Update Order Option
        ///</summary>
        public enum UpdateOrderOption
        {
            InsertUpdateDelete = 0,
            UpdateInsertDelete = 1
        }

        /// <summary>
        ///Used to sort self-referenced table's rows
        ///</summary>
        private class SelfReferenceComparer<R> : object, IComparer<R>
            where R : DataRow
        {
            private DataRelation _relation;

            private int _childFirst;

            internal SelfReferenceComparer(DataRelation relation, bool childFirst)
            {
                this._relation = relation;
                if (childFirst)
                {
                    this._childFirst = -1;
                }
                else
                {
                    this._childFirst = 1;
                }
            }

            private bool IsChildAndParent(R child, R parent)
            {
                Debug.Assert((child != null));
                Debug.Assert((parent != null));
                DataRow newParent = child.GetParentRow(this._relation, DataRowVersion.Default);
                for (
                ; ((newParent != null)
                            && ((object.ReferenceEquals(newParent, child) == false)
                            && (object.ReferenceEquals(newParent, parent) == false)));
                )
                {
                    newParent = newParent.GetParentRow(this._relation, DataRowVersion.Default);
                }
                if ((newParent == null))
                {
                    for (newParent = child.GetParentRow(this._relation, DataRowVersion.Original); ((newParent != null)
                                && ((object.ReferenceEquals(newParent, child) == false)
                                && (object.ReferenceEquals(newParent, parent) == false)));
                    )
                    {
                        newParent = newParent.GetParentRow(this._relation, DataRowVersion.Original);
                    }
                }
                if (object.ReferenceEquals(newParent, parent))
                {
                    return true;
                }
                return false;
            }

            public int Compare(R row1, R row2)
            {
                if (object.ReferenceEquals(row1, row2))
                {
                    return 0;
                }
                if ((row1 == null))
                {
                    return -1;
                }
                if ((row2 == null))
                {
                    return 1;
                }

                // Is row1 the child or grandchild of row2
                if (this.IsChildAndParent(row1, row2))
                {
                    return this._childFirst;
                }

                // Is row2 the child or grandchild of row1
                if (this.IsChildAndParent(row2, row1))
                {
                    return (-1 * this._childFirst);
                }
                return 0;
            }
        }
    }
}
