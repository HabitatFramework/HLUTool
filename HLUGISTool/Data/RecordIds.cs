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
using System.Data;
using System.Linq;
using HLU.Data.Connection;
using HLU.Data.Model;
using HLU.Data.Model.HluDataSetTableAdapters;
using HLU.UI.ViewModel;

namespace HLU.Data
{
    class RecordIds
    {
        #region Fields

        private string _siteID;
        DbBase _db;
        HluDataSet _hluDataset;
        ViewModelWindowMain.GeometryTypes _gisLayerType;
        TableAdapterManager _hluTableAdapterMgr;
        private string _habitatVersion;
        private int _incidCurrentNumber = -1;
        //TODO: Replace with Condition and Secondary
        //int _nextIncidIhsMatrixId = -1;
        //int _nextIncidIhsFormationId = -1;
        //int _nextIncidIhsManagementId = -1;
        //int _nextIncidIhsComplexId = -1;
        int _nextIncidBapId = -1;
        int _nextIncidSourcesId = -1;

        #endregion

        #region ctor

        public RecordIds(DbBase db, HluDataSet hluDataset, 
            TableAdapterManager hluTableAdapterMgr, ViewModelWindowMain.GeometryTypes gisLayerType)
        {
            if (db == null) throw new ArgumentException("db");
            if (hluDataset == null) throw new ArgumentException("hluDataset");
            if (hluTableAdapterMgr == null) throw new ArgumentException("hluTableAdapterMgr");

            _db = db;
            _hluDataset = hluDataset;
            _hluTableAdapterMgr = hluTableAdapterMgr;
            _gisLayerType = gisLayerType;
            if (_hluDataset.lut_last_incid.IsInitialized && _hluDataset.lut_last_incid.Count == 0)
            {
                if (_hluTableAdapterMgr.lut_last_incidTableAdapter == null)
                    _hluTableAdapterMgr.lut_last_incidTableAdapter =
                        new HluTableAdapter<HluDataSet.lut_last_incidDataTable, HluDataSet.lut_last_incidRow>(_db);
                _hluTableAdapterMgr.Fill(_hluDataset,
                    new Type[] { typeof(HluDataSet.lut_last_incidDataTable) }, false);
            }
            _incidCurrentNumber = CurrentMaxIncidNumber(false);
            InitializeIncidChildRecordIds();
        }

        public void InitializeIncidChildRecordIds()
        {
            //TODO: Replace with Condition and Secondary
            object retVal;
            //object retVal = _db.ExecuteScalar(String.Format("SELECT MAX({0}) + 1 FROM {1}",
            //    _db.QuoteIdentifier(_hluDataset.incid_ihs_matrix.matrix_idColumn.ColumnName),
            //    _db.QualifyTableName(_hluDataset.incid_ihs_matrix.TableName)),
            //    _db.Connection.ConnectionTimeout, CommandType.Text);
            //_nextIncidIhsMatrixId = retVal != DBNull.Value && retVal != null ? (int)retVal : 1;

            //retVal = _db.ExecuteScalar(String.Format("SELECT MAX({0}) + 1 FROM {1}",
            //    _db.QuoteIdentifier(_hluDataset.incid_ihs_formation.formation_idColumn.ColumnName),
            //    _db.QualifyTableName(_hluDataset.incid_ihs_formation.TableName)),
            //    _db.Connection.ConnectionTimeout, CommandType.Text);
            //_nextIncidIhsFormationId = retVal != DBNull.Value && retVal != null ? (int)retVal : 1;

            //retVal=_db.ExecuteScalar(String.Format("SELECT MAX({0}) + 1 FROM {1}",
            //    _db.QuoteIdentifier(_hluDataset.incid_ihs_management.management_idColumn.ColumnName),
            //    _db.QualifyTableName(_hluDataset.incid_ihs_management.TableName)),
            //    _db.Connection.ConnectionTimeout, CommandType.Text);
            //_nextIncidIhsManagementId = retVal != DBNull.Value && retVal != null ? (int)retVal : 1;

            //retVal = _db.ExecuteScalar(String.Format("SELECT MAX({0}) + 1 FROM {1}",
            //    _db.QuoteIdentifier(_hluDataset.incid_ihs_complex.complex_idColumn.ColumnName),
            //    _db.QualifyTableName(_hluDataset.incid_ihs_complex.TableName)),
            //    _db.Connection.ConnectionTimeout, CommandType.Text);
            //_nextIncidIhsComplexId = retVal != DBNull.Value && retVal != null ? (int)retVal : 1;

            retVal = _db.ExecuteScalar(String.Format("SELECT MAX({0}) + 1 FROM {1}",
                _db.QuoteIdentifier(_hluDataset.incid_bap.bap_idColumn.ColumnName),
                _db.QualifyTableName(_hluDataset.incid_bap.TableName)),
                _db.Connection.ConnectionTimeout, CommandType.Text);
            _nextIncidBapId = retVal != DBNull.Value && retVal != null ? (int)retVal : 1;

            retVal = _db.ExecuteScalar(String.Format("SELECT MAX({0}) + 1 FROM {1}",
                _db.QuoteIdentifier(_hluDataset.incid_sources.incid_source_idColumn.ColumnName),
                _db.QualifyTableName(_hluDataset.incid_sources.TableName)),
                _db.Connection.ConnectionTimeout, CommandType.Text);
            _nextIncidSourcesId = retVal != DBNull.Value && retVal != null ? (int)retVal : 1;
        }

        #endregion

        #region Public Properties

        public string HabitatVersion
        {
            get
            {
                // Get the habitat_version from the lut_version table.
                if (String.IsNullOrEmpty(_habitatVersion))
                    if (_hluDataset.lut_version.Count > 0)
                        _habitatVersion = _hluDataset.lut_version.ElementAt(_hluDataset.lut_version.Count - 1).habitat_version;
                    else
                        _habitatVersion = "0";

                return _habitatVersion;
            }
        }

        public string SiteID
        {
            get
            {
                if (String.IsNullOrEmpty(_siteID))
                {
                    if (_hluDataset.lut_site_id.Count > 0)
                    {
                        switch (_gisLayerType)
                        {
                            case ViewModelWindowMain.GeometryTypes.Point:
                                _siteID = _hluDataset.lut_site_id.ElementAt(_hluDataset.lut_site_id.Count - 1).site_id_point;
                                break;
                            case ViewModelWindowMain.GeometryTypes.Line:
                                _siteID = _hluDataset.lut_site_id.ElementAt(_hluDataset.lut_site_id.Count - 1).site_id_line;
                                break;
                            case ViewModelWindowMain.GeometryTypes.Polygon:
                                _siteID = _hluDataset.lut_site_id.ElementAt(_hluDataset.lut_site_id.Count - 1).site_id_polygon;
                                break;
                        }
                    }
                    else
                    {
                        _siteID = "0000";
                    }
                }
                return _siteID;
            }
        }

        /// <summary>
        /// Gets the next available INCID, checking lut_last_incid and incid tables in DB.
        /// Increments the number and saves the new value back to lut_last_incid in DB.
        /// </summary>
        public string NextIncid
        {
            get
            {
                _incidCurrentNumber = CurrentMaxIncidNumber(true);
                return IncidString(_incidCurrentNumber);
            }
        }

        public string CurrentIncid
        {
            get { return SiteID + ":" + _incidCurrentNumber.ToString("D7"); }
        }

        //TODO: Replace with Condition and Secondary
        //public int NextIncidIhsMatrixId
        //{
        //    get
        //    {
        //        _nextIncidIhsMatrixId = NextID(_nextIncidIhsMatrixId, _hluDataset.incid_ihs_matrix,
        //            _hluDataset.incid_ihs_matrix.matrix_idColumn.Ordinal);
        //        return _nextIncidIhsMatrixId;
        //    }
        //}

        //public int NextIncidIhsFormationId
        //{
        //    get
        //    {
        //        _nextIncidIhsFormationId = NextID(_nextIncidIhsFormationId, _hluDataset.incid_ihs_formation,
        //            _hluDataset.incid_ihs_formation.formation_idColumn.Ordinal);
        //        return _nextIncidIhsFormationId;
        //    }
        //}

        //public int NextIncidIhsManagementId
        //{
        //    get
        //    {
        //        _nextIncidIhsManagementId = NextID(_nextIncidIhsManagementId, _hluDataset.incid_ihs_management,
        //            _hluDataset.incid_ihs_management.management_idColumn.Ordinal);
        //        return _nextIncidIhsManagementId;
        //    }
        //}

        //public int NextIncidIhsComplexId
        //{
        //    get
        //    {
        //        _nextIncidIhsComplexId = NextID(_nextIncidIhsComplexId, _hluDataset.incid_ihs_complex,
        //            _hluDataset.incid_ihs_complex.complex_idColumn.Ordinal);
        //        return _nextIncidIhsComplexId;
        //    }
        //}

        public int CurrentIncidBapId
        {
            get
            {
                return NextID(_nextIncidBapId, _hluDataset.incid_bap, 
                    _hluDataset.incid_bap.bap_idColumn.Ordinal) - 1;
            }
        }

        public int NextIncidBapId
        {
            get
            {
                _nextIncidBapId = NextID(_nextIncidBapId, _hluDataset.incid_bap, 
                    _hluDataset.incid_bap.bap_idColumn.Ordinal);
                return _nextIncidBapId;
            }
        }

        public int NextIncidSourcesId
        {
            get
            {
                _nextIncidSourcesId = NextID(_nextIncidSourcesId, _hluDataset.incid_sources, 
                    _hluDataset.incid_sources.incid_source_idColumn.Ordinal);
                return _nextIncidSourcesId;
            }
        }

        public int MaxIncidNumber
        {
            get { return (int)Math.Pow((double)10, (double)(IncidString(1).Length - SiteID.Length - 1)) - 1; }
        }

        #endregion

        #region Public methods

        public static int IncidNumber(string incidString)
        {
            try
            {
                int i;
                if (Int32.TryParse(incidString.Split(':')[1], out i))
                    return i;
                else
                    return -1;
            }
            catch { return -1; }
        }

        public string IncidString(int incidNumber)
        {
            return SiteID + ":" + incidNumber.ToString("D7");
        }

        public string MaxToidFragmentId(string toid)
        {
            if (!String.IsNullOrEmpty(toid))
            {
                try
                {
                    object retVal = _db.ExecuteScalar(String.Format("SELECT MAX({0}) FROM {1} WHERE {2} = {3}",
                        _db.QuoteIdentifier(_hluDataset.incid_mm_polygons.toid_fragment_idColumn.ColumnName),
                        _db.QualifyTableName(_hluDataset.incid_mm_polygons.TableName),
                        _db.QuoteIdentifier(_hluDataset.incid_mm_polygons.toidColumn.ColumnName),
                        _db.QuoteValue(toid)), _db.Connection.ConnectionTimeout, CommandType.Text);
                    return retVal.ToString() ?? "00000";
                }
                catch { return null; }
            }
            else
            {
                return null;
            }
        }

        #endregion

        #region Private

        private int CurrentMaxIncidNumber(bool increment)
        {
            try
            {
                int maxIncidNumber = 0;

                // check in-memory incid table
                if (_hluDataset.incid.Count > 0)
                    _hluDataset.incid.Max(r => IncidNumber(r.incid));

                // check lut_last_incid in DB
                _hluTableAdapterMgr.Fill(_hluDataset, typeof(HluDataSet.lut_last_incidDataTable), true);
                HluDataSet.lut_last_incidRow lastIncidRow = null;
                if (_hluDataset.lut_last_incid.Count > 0)
                {
                    lastIncidRow =
                        _hluDataset.lut_last_incid.ElementAt(_hluDataset.lut_last_incid.Count - 1);
                    if (lastIncidRow.last_incid > maxIncidNumber)
                        maxIncidNumber = lastIncidRow.last_incid;
                }

                // check DB incid table
                string sql = String.Format("SELECT MAX({0}) FROM {1}",
                    _db.QuoteIdentifier(_hluDataset.incid.incidColumn.ColumnName),
                    _db.QualifyTableName(_hluDataset.incid.TableName));
                object result = _db.ExecuteScalar(sql, _db.Connection.ConnectionTimeout, CommandType.Text);
                int dbMax;
                if ((result != DBNull.Value) && (result != null) && 
                    ((dbMax = IncidNumber(result.ToString())) > maxIncidNumber)) maxIncidNumber = dbMax;

                if (increment)
                {
                    // increment the greatest value found and save to lut_last_incid
                    maxIncidNumber++;
                    if (lastIncidRow != null)
                        lastIncidRow.last_incid = maxIncidNumber;
                    else
                        lastIncidRow = _hluDataset.lut_last_incid.Addlut_last_incidRow(maxIncidNumber);
                    _hluTableAdapterMgr.lut_last_incidTableAdapter.Update(lastIncidRow);
                }

                return maxIncidNumber;
            }
            catch { return -1; }
        }

        public int NextID<T>(int nextID, T table, int idColumnOrdinal)
            where T : DataTable
        {
            if ((nextID == -1) && (table.Rows.Count > 0) && (table != null))
                nextID = table.AsEnumerable().Max(r => r.Field<int>(idColumnOrdinal)) + 1;

            string sql = String.Format("SELECT MAX({0}) + 1 FROM {1}",
                _db.QuoteIdentifier(table.Columns[idColumnOrdinal].ColumnName), _db.QualifyTableName(table.TableName));
            object result = _db.ExecuteScalar(sql, _db.Connection.ConnectionTimeout, CommandType.Text);
            int dbMax;
            if ((result != DBNull.Value) && (result != null) && Int32.TryParse(result.ToString(), out dbMax) &&
                (nextID < dbMax)) nextID = dbMax;
            else
                nextID += 1;

            return nextID;
        }

        #endregion
    }
}
