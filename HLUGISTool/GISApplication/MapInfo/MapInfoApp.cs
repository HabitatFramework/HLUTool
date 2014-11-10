// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// Copyright © 2013-2014 Thames Valley Environmental Records Centre
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
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using HLU.Data;
using HLU.Data.Model;
using HLU.Properties;
using HLU.UI.ViewModel;
using MapinfoWrapper.Mapinfo;
using MapinfoWrapper.Mapinfo.Internals;
using Microsoft.Win32;
using System.Globalization;

namespace HLU.GISApplication.MapInfo
{
    class MapInfoApp : GISApp
    {
        #region Private Fields

        //private COMMapinfo _mapInfoComObj;
        
        private DMapInfo _mapInfoApp;

        private Process _mapInfoProcess;

        private Process[] _mapInfoProcsPreStart;

        private IntPtr _mapInfoParentWindow;
        
        private string _mapPath;
        
        private string _hluLayer;

        private string _hluLayerOld;

        private string _selName;

        /// <summary>
        /// The list of valid HLU map layers in the workspace.
        /// </summary>
        private List<GISLayer> _hluLayerList;

        /// <summary>
        /// The current valid HLU map layer in the workspace.
        /// </summary>
        private GISLayer _hluCurrentLayer;

        HluGISLayer.incid_mm_polygonsDataTable _hluLayerStructure;

        /// <summary>
        /// Position is equivalent to column ordinal in _hluLayerStructure; 
        /// values at position to ordinal of corresponding field in MapInfo layer (which is one based!)
        /// </summary>
        private int[] _hluFieldMap;

        /// <summary>
        /// Array of field names of HLU MapInfo layer. 
        /// _hluFieldNames[i] = ColumnInfo(_hluLayer, QuoteValue(String.Format("Col{0}", i + 1)), ColumnInfoConstant)
        /// </summary>
        private string[] _hluFieldNames;

        private string _hluColumnList;

        private int _hluMapWindowID = -1;

        /// <summary>
        /// Area unit of measurement.
        /// </summary>
        private string _unitArea = "sq m";

        /// <summary>
        /// Distance unit of measurement.
        /// </summary>
        private string _unitDistance = "m";

        /// <summary>
        /// Maximum (nominal) allowable length of a SQL query.
        /// </summary>
        private int _maxSqlLength = Settings.Default.MaxSqlLengthMapInfo;

        private string _historyGeom1ColumnName = ViewModelWindowMain.HistoryGeometry1ColumnName;

        private string _historyGeom2ColumnName = ViewModelWindowMain.HistoryGeometry2ColumnName;

        private Regex _whereNullRegex = new Regex(@"\s+((?<isnot>IS\s+NOT)|(?<is>IS))\s+(?<null>NULL)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        #endregion

        public MapInfoApp(string mapPath)
        {
            _mapPath = mapPath;
            PopulateTypeMaps();
        }

        /// <summary>
        /// Units in which history reports polygon areas. Defaults to square meters.
        /// </summary>
        public override AreaUnits AreaUnit
        {
            set
            {
                switch (value)
                {
                    case AreaUnits.Acres:
                        _unitArea = "acre";
                        break;
                    case AreaUnits.Hectares:
                        _unitArea = "hectare";
                        break;
                    case AreaUnits.Perches:
                        _unitArea = "perch";
                        break;
                    case AreaUnits.Roods:
                        _unitArea = "rood";
                        break;
                    case AreaUnits.SquareCentimeters:
                        _unitArea = "sq cm";
                        break;
                    case AreaUnits.SquareChains:
                        _unitArea = "sq ch";
                        break;
                    case AreaUnits.SquareFeet:
                        _unitArea = "sq ft";
                        break;
                    case AreaUnits.SquareInches:
                        _unitArea = "sq in";
                        break;
                    case AreaUnits.SquareKilometers:
                        _unitArea = "sq km";
                        break;
                    case AreaUnits.SquareLinks:
                        _unitArea = "sq li";
                        break;
                    case AreaUnits.SquareMeters:
                        _unitArea = "sq m";
                        break;
                    case AreaUnits.SquareMiles:
                        _unitArea = "sq mi";
                        break;
                    case AreaUnits.SquareMillimeters:
                        _unitArea = "sq mm";
                        break;
                    case AreaUnits.SquareRods:
                        _unitArea = "sq rd";
                        break;
                    case AreaUnits.SquareSurveyFeet:
                        _unitArea = "sq  survey ft";
                        break;
                    case AreaUnits.SquareYards:
                        _unitArea = "sq yd";
                        break;
                }
            }
        }

        /// <summary>
        /// Units in which history reports polyline lengths and polygon perimeters. Defaults to meters.
        /// </summary>
        public override DistanceUnits DistanceUnit
        {
            set
            {
                switch (value)
                {
                    case DistanceUnits.Centimeters:
                        _unitDistance = "cm";
                        break;
                    case DistanceUnits.Chains:
                        _unitDistance = "ch";
                        break;
                    case DistanceUnits.Feet:
                        _unitDistance = "ft";
                        break;
                    case DistanceUnits.Inches:
                        _unitDistance = "in";
                        break;
                    case DistanceUnits.Kilometers:
                        _unitDistance = "km";
                        break;
                    case DistanceUnits.Links:
                        _unitDistance = "li";
                        break;
                    case DistanceUnits.Meters:
                        _unitDistance = "m";
                        break;
                    case DistanceUnits.Miles:
                        _unitDistance = "mi";
                        break;
                    case DistanceUnits.Millimeters:
                        _unitDistance = "mm";
                        break;
                    case DistanceUnits.NauticalMiles:
                        _unitDistance = "nmi";
                        break;
                    case DistanceUnits.SurveyFeet:
                        _unitDistance = "survey ft";
                        break;
                    case DistanceUnits.Yards:
                        _unitDistance = "yd";
                        break;
                }
            }
        }

        /// <summary>
        /// Maximum (nominal) allowable length of a SQL query.
        /// </summary>
        public override int MaxSqlLength
        {
            get { return _maxSqlLength; }
        }

        public override object ApplicationObject
        {
            get { return _mapInfoApp; }
        }

        public override GISApplications ApplicationType
        {
            get { return GISApplications.MapInfo; }
        }

        public override bool IsRunning
        {
            get
            {
                try
                {
                    if (!WinAPI.IsWindow(_mapInfoParentWindow))
                    {
                        _mapInfoApp = null;
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                catch
                {
                    _mapInfoApp = null;
                    return false;
                }
            }
        }

        public override bool IsEditing
        {
            get
            {
                int editableLayerNo = Int32.Parse(_mapInfoApp.Eval(String.Format("MapperInfo({0}, {1})",
                    _hluMapWindowID, (int)MapInfoConstants.MapperInfo.MAPPER_INFO_EDIT_LAYER)));

                if (editableLayerNo > 0)
                {
                    string editableLayerName = _mapInfoApp.Eval(String.Format("LayerInfo({0}, {1}, {2})",
                        _hluMapWindowID, editableLayerNo, (int)MapInfoConstants.LayerInfo.LAYER_INFO_NAME));

                    if (editableLayerName == _hluLayer)
                    {
                        return _mapInfoApp.Eval(String.Format("LayerInfo({0}, {1}, {2})",
                            _hluMapWindowID, editableLayerNo, (int)MapInfoConstants.LayerInfo.LAYER_INFO_EDITABLE)) == "T";
                    }
                }

                return false;
            }
        }

        /// <summary>
        /// The name of the HLU layer/table.
        /// </summary>
        public override string HluLayerName
        {
            get { return _hluLayer; }
        }

        /// <summary>
        /// The name of the incid field in the HLU table structure.
        /// </summary>
        public override string IncidFieldName
        {
            get { return GetFieldName(_hluLayerStructure.incidColumn.Ordinal); }
        }

        /// <summary>
        /// The number of valid hlu layers.
        /// </summary>
        public override int HluLayerCount
        {
            get { return _hluLayerList.Count(); }
        }

        /// <summary>
        /// The list of valid hlu layers.
        /// </summary>
        public override List<GISLayer> ValidHluLayers
        {
            get { return _hluLayerList; }
        }

        /// <summary>
        /// The properties of the current hlu layer.
        /// </summary>
        public override GISLayer CurrentHluLayer
        {
            get { return _hluCurrentLayer; }
        }

        /// <summary>
        /// Gets the currently selected map features and writes their UIDs into a DataTable.
        /// </summary>
        /// <param name="resultTable">DataTable to be filled with attributes of the selected features. 
        /// Column names and data types (per type map) match field names of GIS layer.</param>
        public override void ReadMapSelection(ref DataTable resultTable)
        {
            try { ReadSelectedRows(true, true, false, false, _hluLayer, ref resultTable); }
            catch { }
        }

        /// <summary>
        /// Flashes a single list of selected features whose UIDs are passed in.
        /// </summary>
        /// <param name="whereClause">Where clause identifying the features to be flashed.</param>
        //---------------------------------------------------------------------
        // CHANGED: CR23 (Merged features)
        // Flash all the features relating to the selected incid at once.  In this case
        // their criteria will fit within a single item in the outer list so they will
        // all flash at once.
        public override void FlashSelectedFeature(List<SqlFilterCondition> whereClause)
        {
            if (String.IsNullOrEmpty(_selName)) return;

            string prevSelName = String.Copy(_selName);

            try
            {
                DataTable selTable = _hluLayerStructure.Clone();
                selTable.TableName = _selName;

                for (int i = 0; i < whereClause.Count; i++)
                {
                    SqlFilterCondition cond = whereClause[i];
                    cond.Table = selTable;
                    whereClause[i] = cond;
                }

                if (SqlSelect(false, false, selTable.Columns.Cast<DataColumn>().ToArray(), 
                    _selName, false, false, whereClause, null) == null) return;

                string flashLayer = _mapInfoApp.Eval(String.Format("SelectionInfo({0})",
                    (int)MapInfoConstants.SelectionInfo.SEL_INFO_SELNAME));

                _mapInfoApp.Do(String.Format("Add Map Window {0} Layer {1}", _hluMapWindowID, flashLayer));
                _mapInfoApp.Do(String.Format("Set Map Window {0} Layer {1} Display Global " + 
                    "Global Pen(1, 2, {2}) Global Brush(2, {2}, {3})", _hluMapWindowID, flashLayer, 
                    (int)MapInfoConstants.Colors.RED, (int)MapInfoConstants.Colors.WHITE));

                for (int i = 0; i < 2; i++)
                {
                    _mapInfoApp.Do(String.Format("Set Map Window {0} Layer {1} Display Off", _hluMapWindowID, flashLayer));
                    Thread.Sleep(300);
                    _mapInfoApp.Do(String.Format("Set Map Window {0} Layer {1} Display Global", _hluMapWindowID, flashLayer));
                }

                _mapInfoApp.Do(String.Format("Set Map Redraw Off")); 
                _mapInfoApp.Do(String.Format("Remove Map Window {0} Layer {1}", _hluMapWindowID, flashLayer));
                _mapInfoApp.Do(String.Format("Close Table {0}", flashLayer));
                _mapInfoApp.Do(String.Format("Set Map Redraw On"));
            }
            catch { }
            finally
            {
                // Reset the original selection table name
                _selName = String.Copy(prevSelName);

                // Re-select all the features in the original selection
                DataTable resetTable = _hluLayerStructure.Clone();
                resetTable.TableName = _selName;
                List<SqlFilterCondition> reSelectionWhereClause;
                reSelectionWhereClause = new List<SqlFilterCondition>();
                SqlFilterCondition cond = whereClause[0];
                cond.Operator = "IS NOT NULL";
                cond.BooleanOperator = "AND";
                cond.OpenParentheses = String.Empty;
                cond.CloseParentheses = String.Empty;
                cond.Table = resetTable;
                reSelectionWhereClause.Add(cond);
                SqlSelect(false, false, resetTable.Columns.Cast<DataColumn>().ToArray(),
                    _selName, false, false, reSelectionWhereClause, null);
            }
        }
        //---------------------------------------------------------------------

        /// <summary>
        /// Flashes multiple lists of selected features whose UIDs are passed in.
        /// </summary>
        /// <param name="whereClauses">Where clauses identifying the features to be flashed.</param>
        //---------------------------------------------------------------------
        // CHANGED: CR23 (Merged features)
        // Flash all the features relating to the selected incid at once.  In this case
        // they will flash in groups if there are too many criteria to fit within a single
        // item in the outer list.
        public override void FlashSelectedFeatures(List<List<SqlFilterCondition>> whereClauses)
        {
            if (String.IsNullOrEmpty(_selName)) return;

            string prevSelName = String.Copy(_selName);

            try
            {
                DataTable selTable = _hluLayerStructure.Clone();
                selTable.TableName = _selName;

                foreach (List<SqlFilterCondition> whereClause in whereClauses)
                {
                    for (int i = 0; i < whereClause.Count; i++)
                    {
                        SqlFilterCondition cond = whereClause[i];
                        cond.Table = selTable;
                        whereClause[i] = cond;
                    }

                    if (SqlSelect(false, false, selTable.Columns.Cast<DataColumn>().ToArray(),
                        _selName, false, false, whereClause, null) == null) return;

                    string flashLayer = _mapInfoApp.Eval(String.Format("SelectionInfo({0})",
                        (int)MapInfoConstants.SelectionInfo.SEL_INFO_SELNAME));

                    _mapInfoApp.Do(String.Format("Add Map Window {0} Layer {1}", _hluMapWindowID, flashLayer));
                    _mapInfoApp.Do(String.Format("Set Map Window {0} Layer {1} Display Global " +
                        "Global Pen(1, 2, {2}) Global Brush(2, {2}, {3})", _hluMapWindowID, flashLayer,
                        (int)MapInfoConstants.Colors.RED, (int)MapInfoConstants.Colors.WHITE));

                    //---------------------------------------------------------------------
                    // FIX: 018 Bring ArcGIS and MapInfo into line by flashing all features twice
                    for (int i = 0; i < 2; i++)
                    //---------------------------------------------------------------------
                    {
                        _mapInfoApp.Do(String.Format("Set Map Window {0} Layer {1} Display Off", _hluMapWindowID, flashLayer));
                        Thread.Sleep(300);
                        _mapInfoApp.Do(String.Format("Set Map Window {0} Layer {1} Display Global", _hluMapWindowID, flashLayer));
                    }

                    _mapInfoApp.Do(String.Format("Set Map Redraw Off"));
                    _mapInfoApp.Do(String.Format("Remove Map Window {0} Layer {1}", _hluMapWindowID, flashLayer));
                    _mapInfoApp.Do(String.Format("Close Table {0}", flashLayer));
                    _mapInfoApp.Do(String.Format("Set Map Redraw On"));
                }
            }
            catch { }
            finally
            {
                // Reset the original selection table name
                _selName = String.Copy(prevSelName);

                // Re-select all the features in the original selection
                DataTable resetTable = _hluLayerStructure.Clone();
                resetTable.TableName = _selName;
                List<SqlFilterCondition> reSelectionWhereClause;
                reSelectionWhereClause = new List<SqlFilterCondition>();
                SqlFilterCondition cond = whereClauses[0][0];
                cond.Operator = "IS NOT NULL";
                cond.BooleanOperator = "AND";
                cond.OpenParentheses = String.Empty;
                cond.CloseParentheses = String.Empty;
                cond.Table = resetTable;
                reSelectionWhereClause.Add(cond);
                SqlSelect(false, false, resetTable.Columns.Cast<DataColumn>().ToArray(),
                    _selName, false, false, reSelectionWhereClause, null);
            }
        }
        //---------------------------------------------------------------------

        /// <summary>
        /// Following a feature split performed by the user with regular GIS tools, updates the toid_fragment_id of 
        /// the newly created feature(s), incrementing from the highest toid_fragment_id in use for that toid, which is
        /// passed in as parameter lastToidFragmentID. toid_fragment_id is an integer stored as varchar(5) with 
        /// leading zeros (.ToString("D5")).
        /// </summary>
        /// <param name="currentToidFragmentID">Current toid_fragment_id for the selected toid.</param>
        /// <param name="lastToidFragmentID">Highest toid_fragment_id for the selected toid.</param>
        /// <param name="selectionWhereClause">Where clause underlying the current selection set.</param>
        /// <param name="historyColumns">All columns of the GIS layer structure except the geometry property columns.</param>
        /// <returns>DataTable with the columns in the historyColumns parameter, containing the attributes of all 
        /// the features resulting from the split. 
        /// The first row corresponds to the original feature (the one that was split).</returns>
        public override DataTable SplitFeature(string currentToidFragmentID, string lastToidFragmentID,
            List<SqlFilterCondition> selectionWhereClause, DataColumn[] historyColumns)
        {
            bool rollbackChanges = false;

            try
            {
                if (String.IsNullOrEmpty(_selName)) _selName = _mapInfoApp.Eval(String.Format(
                    "SelectionInfo({0})", (int)MapInfoConstants.SelectionInfo.SEL_INFO_TABLENAME));
                if (!TableExists(_selName)) return null;

                string origSelName = _selName;

                DataTable selTable = CreateHistoryTable(!String.IsNullOrEmpty(_selName) &&
                    _mapInfoApp.Eval(String.Format("SelectionInfo({0})",
                    (int)MapInfoConstants.SelectionInfo.SEL_INFO_TABLENAME)) ==
                    _hluLayer ? _selName : _hluLayer, false, historyColumns);

                ChangeWhereClauseTable(new KeyValuePair<DataTable, DataTable>[] { new KeyValuePair<DataTable, DataTable>(_hluLayerStructure, selTable) },
                    ref selectionWhereClause);

                // get history data for selected features
                DataTable historyTable = SqlSelect(false, false, selTable.Columns.Cast<DataColumn>().ToArray(),
                    _selName, true, false, selectionWhereClause, null);
                if ((historyTable == null) || (historyTable.Rows.Count < 2)) return null;

                // Get the number of selected rows from MapInfo
                int selNum = Int32.Parse(_mapInfoApp.Eval(String.Format("SelectionInfo({0})",
                    (int)MapInfoConstants.SelectionInfo.SEL_INFO_NROWS)));

                // Set the new ToidFragmentID
                string numFormat = String.Format("D{0}", lastToidFragmentID.Length);
                int newToidFragmentIDnum = Int32.Parse(lastToidFragmentID);
                string toidFragFieldName = GetFieldName(_hluLayerStructure.toid_fragment_idColumn.Ordinal);

                rollbackChanges = true;

                //---------------------------------------------------------------------
                // FIXED: KI106 (Shape area and length values)
                // Includes updates for the geom1 and geom2 columns as the features
                // have changed in size
                string lastColName = GetFieldName(_hluLayerStructure.ihs_summaryColumn.Ordinal);
                int ixGeom1 = Int32.Parse(_mapInfoApp.Eval(String.Format("ColumnInfo({0}, {1}, {2})", _selName,
                    QuoteValue(lastColName), (int)MapInfoConstants.ColumnInfo.COL_INFO_NUM))) + 1;
                int ixGeom2 = ixGeom1 + 1;
                string geomCol1Name = _mapInfoApp.Eval(String.Format("ColumnInfo({0}, {1}, {2})", _selName,
                    QuoteValue(String.Format("Col{0}", ixGeom1)), (int)MapInfoConstants.ColumnInfo.COL_INFO_NAME));
                string geomCol2Name = _mapInfoApp.Eval(String.Format("ColumnInfo({0}, {1}, {2})", _selName,
                    QuoteValue(String.Format("Col{0}", ixGeom2)), (int)MapInfoConstants.ColumnInfo.COL_INFO_NAME));

                double geom1;
                double geom2;
                string fetchCommand = String.Format("Fetch Next From {0}", _selName);
                _mapInfoApp.Do(String.Format("Fetch First From {0}", _selName));

                // For each of the selected rows (there will be more than one if the feature
                // was split into more than two parts)
                for (int i = 1; i <= selNum; i++)
                {
                    // Get the geometry values from the layer
                    GetGeometryInfo(_selName, out geom1, out geom2);

                    // Update the geometry values
                    _mapInfoApp.Do(String.Format("Update {0} Set {1} = {3}, {2} = {4} Where RowID = {5}",
                        _selName, QuoteIdentifier(geomCol1Name), QuoteIdentifier(geomCol2Name), geom1, geom2, i));

                    // Update the toid_fragment_id for all but the first feature (which can keep the existing
                    // value)
                    if (i > 1)
                    {
                        _mapInfoApp.Do(String.Format("Update {0} Set {1} = {2} Where RowID = {3}", _selName,
                            toidFragFieldName, QuoteValue((++newToidFragmentIDnum).ToString(numFormat)), i));
                    }
                    
                    // Fetch the next row
                    _mapInfoApp.Do(fetchCommand);
                }
                //---------------------------------------------------------------------

                rollbackChanges = false;

                // remove cutter object from cosmetic layer
                _mapInfoApp.Do(String.Format("Delete From LayerInfo({0}, 0, 1)", _hluMapWindowID));

                // remove cutter object from HLU layer
                DataTable cleanupTable = selTable.Clone();
                cleanupTable.TableName = _hluLayer;
                DataTable retTable = SqlSelect(true, false, cleanupTable.Columns.Cast<DataColumn>().ToArray(),
                    _hluLayer, false, false, new List<SqlFilterCondition>(new SqlFilterCondition[] 
                    { new SqlFilterCondition( cleanupTable, cleanupTable.Columns[_hluLayerStructure.incidColumn.ColumnName], 
                        String.Empty) }), null);
                if ((retTable != null) && TableExists(_selName) && (_selName != origSelName) && 
                    (Int32.Parse(_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})", _selName,
                    (int)MapInfoConstants.TableInfo.TAB_INFO_NROWS))) > 0))
                {
                    _mapInfoApp.Do(String.Format("Delete From {0}", _selName));
                    _mapInfoApp.Do(String.Format("Close Table {0}", _selName));
                }

                CommitChanges();

                // re-select split polygons (all rows in orignal selection)
                List<SqlFilterCondition> reSelectionWhereClause;
                //---------------------------------------------------------------------
                // FIX: 029 Ignore case when comparing column names
                // Ignore case when comparing column names so that GIS layer
                // names may be mixed/upper case.
                var q = selectionWhereClause.Where(c => c.Column.ColumnName.ToLower() == _hluLayerStructure.toidColumn.ColumnName.ToLower());
                if (q.Count() == 1)
                {
                    selTable = CreateHistoryTable(_hluLayer, false, historyColumns);
                    SqlFilterCondition cond = new SqlFilterCondition("AND", selTable,
                        _hluLayerStructure.toidColumn, q.ElementAt(0).Value);
                    cond.Table = selTable;
                    reSelectionWhereClause = new List<SqlFilterCondition>();
                    reSelectionWhereClause.Add(cond);
                    //---------------------------------------------------------------------
                    // FIXED: KI110 (Physical split)
                    // The filter should search for the currentToidFragmentID (pre-split)
                    // and the newTOIDFragmentID that was created (post-split).
                    cond = new SqlFilterCondition("AND", selTable,
                        selTable.Columns[_hluLayerStructure.toid_fragment_idColumn.ColumnName], typeof(DataColumn), "(", String.Empty, currentToidFragmentID);
                    cond.Operator = "=";
                    reSelectionWhereClause.Add(cond);
                    cond = new SqlFilterCondition("OR", selTable,
                        selTable.Columns[_hluLayerStructure.toid_fragment_idColumn.ColumnName], typeof(DataColumn), "(", String.Empty, lastToidFragmentID);
                    cond.Operator = ">";
                    reSelectionWhereClause.Add(cond);
                    cond = new SqlFilterCondition("AND", selTable,
                        selTable.Columns[_hluLayerStructure.toid_fragment_idColumn.ColumnName], typeof(DataColumn), String.Empty, "))", newToidFragmentIDnum.ToString(numFormat));
                    cond.Operator = "<=";
                    reSelectionWhereClause.Add(cond);
                    //---------------------------------------------------------------------
                    SqlSelect(true, true, selTable.Columns.Cast<DataColumn>().ToArray(),
                        _hluLayer, false, false, reSelectionWhereClause, null);
                }
                else if (TableExists(origSelName))
                {
                    reSelectionWhereClause = new List<SqlFilterCondition>();
                    SqlFilterCondition cond = selectionWhereClause[0];
                    cond.Operator = "IS NOT NULL";
                    cond.BooleanOperator = "AND";
                    cond.OpenParentheses = String.Empty;
                    cond.CloseParentheses = String.Empty;
                    selTable.TableName = origSelName;
                    reSelectionWhereClause.Add(cond);
                    retTable = SqlSelect(true, true, selTable.Columns.Cast<DataColumn>().ToArray(),
                        origSelName, false, false, reSelectionWhereClause, null);
                }

                if (TableExists(origSelName))
                    _mapInfoApp.Do(String.Format("Close Table {0}", origSelName));

                return historyTable;
            }
            catch
            {
                if (rollbackChanges)
                    _mapInfoApp.RunCommand(String.Format("Rollback Table {0}", _hluLayer));
                return null;
            }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR10 (Attribute updates for incid subsets)
        // The old incid number is passed together with the new incid
        // number so that only features belonging to the old incid are
        // updated.
        public override DataTable SplitFeaturesLogically(string oldIncid, string newIncid, DataColumn[] historyColumns)
        {
            try
            {
                if (!TableExists(_selName)) return null;

                DataTable historyTable = CreateHistoryTable(_selName, true, historyColumns);

                int ixGeom1 = historyColumns.Length;
                int ixGeom2 = historyColumns.Length + 1;

                //---------------------------------------------------------------------
                // FIXED: KI107 (GIS layer column names)
                // Ignore case when comparing column names so that GIS layer names may be mixed/upper case
                string[] historyColumnNames = historyTable.Columns.Cast<DataColumn>()
                    .Select(c => _hluFieldNames.ToList().Contains(c.ColumnName, StringComparer.OrdinalIgnoreCase) ?
                    GetFieldName(_hluLayerStructure.Columns[c.ColumnName].Ordinal).ToLower() : c.ColumnName.ToLower()).ToArray();

                var q = historyColumns.Where(c => !_hluFieldNames.ToList().Contains(c.ColumnName, StringComparer.OrdinalIgnoreCase));
                string newToidFragmentColumnName = q.Count() == 1 ? 
                    q.ElementAt(0).ColumnName.Replace(GISApp.HistoryAdditionalFieldsDelimiter, String.Empty) : null;
                //---------------------------------------------------------------------

                int numRows = Int32.Parse(_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})",
                    _selName, (int)MapInfoConstants.TableInfo.TAB_INFO_NROWS)));

                string numFormat = String.Format("D{0}", _hluLayerStructure.toid_fragment_idColumn.MaxLength);

                string fetchCommand = String.Format("Fetch Rec {0} From {1}", "{0}", _selName);
                string incidCommand = String.Format("{0}.{1}", _selName,
                    GetFieldName(_hluLayerStructure.incidColumn.Ordinal));
                string fragCommand = String.Format("{0}.{1}", _selName,
                    GetFieldName(_hluLayerStructure.toid_fragment_idColumn.Ordinal));
                string toidFragmentCommand = String.Format("{0}.{1}", _selName,
                    GetFieldName(_hluLayerStructure.toid_fragment_idColumn.Ordinal));
                string readCommandTemplate = String.Format("{0}.{1}", _selName, "{0}");
                string updateCommandTemplate = String.Format("Update {0} Set {1} = {2} Where RowID = {3}",
                    _selName, GetFieldName(_hluLayerStructure.incidColumn.Ordinal), QuoteValue(newIncid), "{0}");
                
                for (int i = 1; i <= numRows; i++)
                {
                    _mapInfoApp.RunCommand(String.Format(fetchCommand, i));
                    //---------------------------------------------------------------------
                    // CHANGED: CR10 (Attribute updates for incid subsets)
                    // Only collect the history details and update the incid number if
                    // the each feature belongs to the old incid.
                    if (_mapInfoApp.Eval(incidCommand) == oldIncid)
                    {
                        CollectHistory(ixGeom1, ixGeom2, readCommandTemplate, historyColumnNames, ref historyTable);
                        if (!String.IsNullOrEmpty(newToidFragmentColumnName))
                            historyTable.Rows[historyTable.Rows.Count - 1][newToidFragmentColumnName] = _mapInfoApp.Eval(fragCommand);
                        _mapInfoApp.Do(String.Format(updateCommandTemplate, i));
                    }
                    //---------------------------------------------------------------------
                }

                if ((historyTable != null) && !CommitChanges())
                    historyTable = null;

                return historyTable;
            }
            catch
            {
                _mapInfoApp.Do(String.Format("Rollback Table {0}", _hluLayer));
                return null;
            }
        }
        //---------------------------------------------------------------------

        private void CollectHistory(int ixGeom1, int ixGeom2, string readCommandTemplate, 
            string[] historyColumnNames, ref DataTable historyTable)
        {
            double geom1;
            double geom2;
            object[] itemArray = new object[historyTable.Columns.Count];
            for (int i = 0; i < ixGeom1; i++)
            {
                //---------------------------------------------------------------------
                // FIXED: KI107 (GIS layer column names)
                // Ignore case when comparing column names so that GIS layer names may be mixed/upper case
                if (_hluFieldNames.ToList().Contains(historyColumnNames[i], StringComparer.OrdinalIgnoreCase))
                    itemArray[i] = _mapInfoApp.Eval(String.Format(readCommandTemplate, historyColumnNames[i]));
                //---------------------------------------------------------------------
            }
            GetGeometryInfo(_selName, out geom1, out geom2);
            itemArray[ixGeom1] = geom1;
            itemArray[ixGeom2] = geom2;
            historyTable.Rows.Add(itemArray);
        }

        private void ChangeWhereClauseTable(KeyValuePair<DataTable, DataTable>[] tables, 
            ref List<SqlFilterCondition> whereClause)
        {
            DataTable prevTable = null;
            for (int i = 0; i < whereClause.Count; i++)
            {
                SqlFilterCondition cond = whereClause[i];
                if (cond.Table != prevTable)
                {
                    var q = tables.Where(t => t.Key.TableName == cond.Table.TableName);
                    if (q.Count() == 1) prevTable = q.ElementAt(0).Value;
                }
                if (prevTable != null)
                {
                    cond.Table = prevTable;
                    whereClause[i] = cond;
                }
            }
        }

        private DataTable CreateHistoryTable(string tableName, bool addGeomInfoColumns, DataColumn[] historyColumns)
        {
            DataTable historyTable = new DataTable(tableName);
            foreach (DataColumn c in historyColumns)
                historyTable.Columns.Add(new DataColumn(
                    c.ColumnName.Replace(GISApp.HistoryAdditionalFieldsDelimiter, String.Empty), c.DataType));

            if (addGeomInfoColumns)
                return AddGeometryInfoColumns(historyTable);
            else
                return historyTable;
        }

        private DataTable AddGeometryInfoColumns(DataTable resultTable)
        {
            int ordinal = resultTable.Columns.IndexOf(_historyGeom1ColumnName);
            if (ordinal == -1)
                resultTable.Columns.Add(new DataColumn(_historyGeom1ColumnName, typeof(double)));
            else if (resultTable.Columns[ordinal].DataType != typeof(double))
                return null;

            ordinal = resultTable.Columns.IndexOf(_historyGeom2ColumnName);
            if (ordinal == -1)
                resultTable.Columns.Add(new DataColumn(_historyGeom2ColumnName, typeof(double)));
            else if (resultTable.Columns[ordinal].DataType != typeof(double))
                return null;

            return resultTable;
        }

        private void GetGeometryInfo(string layer, out double geom1, out double geom2)
        {
            MapInfoConstants.ObjectInfoTypes objType = (MapInfoConstants.ObjectInfoTypes)Int32.Parse(
                _mapInfoApp.Eval(String.Format("ObjectInfo({0}.obj, {1})", layer, 
                (int)MapInfoConstants.ObjectInfo.OBJ_INFO_TYPE)));

            geom1 = -1;
            geom2 = -1;
            
            switch (objType)
            {
                case MapInfoConstants.ObjectInfoTypes.OBJ_TYPE_POINT:
                    geom1 = Convert.ToDouble(_mapInfoApp.Eval(String.Format("ObjectGeography({0}.obj, {1}",
                        layer, (int)MapInfoConstants.ObjectGeography.OBJ_GEO_POINTX)));
                    geom2 = Convert.ToDouble(_mapInfoApp.Eval(String.Format("ObjectGeography({0}.obj, {1}",
                        layer, (int)MapInfoConstants.ObjectGeography.OBJ_GEO_POINTY)));
                    break;
                case MapInfoConstants.ObjectInfoTypes.OBJ_TYPE_PLINE:
                    geom1 = Convert.ToDouble(_mapInfoApp.Eval(
                        String.Format("ObjectLen({0}.obj, \"{1}\"", layer, _unitDistance)));
                    break;
                case MapInfoConstants.ObjectInfoTypes.OBJ_TYPE_REGION:
                    geom1 = Convert.ToDouble(_mapInfoApp.Eval(
                        String.Format("Perimeter({0}.obj, \"{1}\")", layer, _unitDistance)));
                    geom2 = Convert.ToDouble(_mapInfoApp.Eval(
                        String.Format("Area({0}.obj, \"{1}\")", layer, _unitArea)));
                    break;
            }
        }

        public override void ZoomSelected()
        {
            _mapInfoApp.RunMenuCommand((int)MapInfoConstants.MenuDef95Query.M_QUERY_FIND_SELECTION);
        }

        private void CreateIndexes(string layerName, string[] columnNames)
        {
            string indexStatement = String.Format("Create Index On {0}({1})", layerName, "{0}");

            // Create indexes on the specified columns.
            _mapInfoApp.Do("Set ProgressBars Off");
            foreach (string colName in columnNames)
            {
                try { _mapInfoApp.Do(String.Format(indexStatement, colName)); }
                catch { }
            }
            _mapInfoApp.Do("Set ProgressBars On");
        }

        private void DropIndexes(string layerName, string[] columnNames)
        {
            string dropIndexStatement = String.Format("Drop Index {0}({1})", layerName, "{0}");

            // Drop any indexes found on the specified columns.
            _mapInfoApp.Do("Set ProgressBars Off");
            foreach (string colName in columnNames)
            {
                try { _mapInfoApp.Do(String.Format(dropIndexStatement, colName)); }
                catch { }
            }
            _mapInfoApp.Do("Set ProgressBars On");
        }

        private void ToggleLayerEditability(bool on, string layerName)
        {
            string editability = "F";

            // Determine which layer is currently editable in MapInfo
            int editableLayerNo = Int32.Parse(_mapInfoApp.Eval(String.Format("MapperInfo({0}, {1})",
                _hluMapWindowID, (int)MapInfoConstants.MapperInfo.MAPPER_INFO_EDIT_LAYER)));

            switch (editableLayerNo)
            {
                case -1: // no layer editable
                    editability = "F";
                    break;
                case 0: // cosmetic layer
                    if (layerName.ToLower().StartsWith("cosmetic") || layerName == "0")
                    {
                        layerName = "0";
                        editability = "T";
                    }
                    else
                    {
                        editability = "F";
                    }
                    break;
                default:
                    // Determine the name of the currently editable layer
                    string editableLayerName = _mapInfoApp.Eval(String.Format("LayerInfo({0}, {1}, {2})",
                        _hluMapWindowID, editableLayerNo, (int)MapInfoConstants.LayerInfo.LAYER_INFO_NAME));
                    if (editableLayerName == layerName)
                    {
                        editability = _mapInfoApp.Eval(String.Format("LayerInfo({0}, {1}, {2})",
                            _hluMapWindowID, editableLayerNo, (int)MapInfoConstants.LayerInfo.LAYER_INFO_EDITABLE));
                    }
                    break;
            }

            if (on)
            {
                if (editability == "F")
                    _mapInfoApp.Do(String.Format("Set Map Layer {0} Editable On", layerName));
            }
            else
            {
                if (editability == "T")
                    _mapInfoApp.Do(String.Format("Set Map Layer {0} Editable Off", layerName));
            }
        }

        /// <summary>
        /// Merges two or more selected features that must share the same incid and toid. The new feature receives the
        /// toid_fragment_id passed in as parameter newToidFragmentID. 
        /// The output feature is identified by the resultWhereClause; the features to be merged with it by the mergeWhereClause.
        /// </summary>
        /// <param name="newToidFragmentID">toid_fragment_id to be assigned to the new, merged feature.</param>
        /// <param name="resultWhereClause">Where clause identifying the result feature (previously chosen by user) 
        /// that the other features are to be merged with.</param>
        /// <param name="mergeWhereClause">Where clause identifying the features to be merged with result feature.</param>
        /// <param name="historyColumns">The columns of the history table returned by the method.</param>
        /// <returns>A DataTable with the columns in the historyColumns parameter containing the attributes of the 
        /// merged features.</returns>
        public override DataTable MergeFeatures(string newToidFragmentID, 
            List<SqlFilterCondition> resultWhereClause, DataColumn[] historyColumns)
        {
            bool rollbackChanges = false;

            try
            {
                // Get the name of the table that the current selection is based on
                if (String.IsNullOrEmpty(_selName)) _selName = _mapInfoApp.Eval(String.Format(
                    "SelectionInfo({0})", (int)MapInfoConstants.SelectionInfo.SEL_INFO_TABLENAME));
                if (!TableExists(_selName)) return null;

                DataTable selTable = CreateHistoryTable(_selName, false, historyColumns);
                DataColumn[] selColumns = selTable.Columns.Cast<DataColumn>().ToArray();

                ChangeWhereClauseTable(new KeyValuePair<DataTable, DataTable>[] 
                { new KeyValuePair<DataTable, DataTable>(_hluLayerStructure, selTable) },
                    ref resultWhereClause);

                // Select the result feature as the target
                DataTable resultTable = SqlSelect(false, false, selColumns, _selName, false, 
                    false, resultWhereClause, null);
                if (String.IsNullOrEmpty(_selName)) return null;
                _mapInfoApp.Do("Set Target On");

                // Select all the other features (excluding the target)
                DataTable historyTable = SqlSelect(false, false, selColumns, _selName, true, 
                    true, resultWhereClause, null);
                if (historyTable == null) return null;

                // delete result feature row from history
                DataRow delRow = historyTable.AsEnumerable().FirstOrDefault(r => 
                    r.Field<string>(_hluLayerStructure.toid_fragment_idColumn.ColumnName) == newToidFragmentID);
                if (delRow != null) delRow.Delete();
                historyTable.AcceptChanges();

                // set HLU layer editable
                ToggleLayerEditability(true, _hluLayer);
                
                // Set the column expression for the 'Objects Combine ...' statement below
                // to set the toid_fragment_id as required
                string toidFragIdClause = String.Format("{0} = {1}",
                    GetFieldName(_hluLayerStructure.toid_fragment_idColumn.Ordinal), 
                    QuoteValue(newToidFragmentID));

                rollbackChanges = true;

                // merge selected features into the target feature
                //---------------------------------------------------------------------
                // FIX: 026 Hide progress bars during MapInfo processing
                _mapInfoApp.Do("Set ProgressBars Off");
                _mapInfoApp.Do(String.Format("Objects Combine Into Target Data {0}", toidFragIdClause));
                _mapInfoApp.Do("Set ProgressBars On");

                rollbackChanges = false;

                //---------------------------------------------------------------------
                // FIXED: KI100 (Physical Merge)

                // Clear the target feature
                _mapInfoApp.Do("Set Target Off");

                // Don't set the HLU layer as non editable after a physical merge
                    // set HLU layer non editable
                    //ToggleLayerEditability(false, _hluLayer);
                //---------------------------------------------------------------------
                
                // Close the table containing the old (pre-merge) selection
                _mapInfoApp.Do(String.Format("Close Table {0}", _selName));

                // add result feature to history (for updating incid_mm_polygon geometry data)
                // merged object is current selection immediately following merge
                _selName = _mapInfoApp.Eval(String.Format("SelectionInfo({0})",
                    (int)MapInfoConstants.SelectionInfo.SEL_INFO_SELNAME));
                if (string.IsNullOrEmpty(_selName)) return null;

                int ixGeom1 = historyColumns.Length;
                int ixGeom2 = historyColumns.Length + 1;
                string fetchCommand = String.Format("Fetch Rec {0} From {1}", "{0}", _selName);
                string readCommandTemplate = String.Format("{0}.{1}", _selName, "{0}");
                string[] historyColumnNames = historyTable.Columns.Cast<DataColumn>()
                    .Select(c => _hluFieldNames.Contains(c.ColumnName) ?
                    GetFieldName(_hluLayerStructure.Columns[c.ColumnName].Ordinal) : c.ColumnName).ToArray();

                // There should be only one row in the table after the merge
                if (Int32.Parse(_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})",
                   _selName, (int)MapInfoConstants.TableInfo.TAB_INFO_NROWS))) == 1)
                {
                    _mapInfoApp.RunCommand(String.Format(fetchCommand, 1));
                    
                    //---------------------------------------------------------------------
                    // FIXED: KI106 (Shape area and length values)
                    // Include updates for the geom1 and geom2 columns automatically
                    string lastColName = GetFieldName(_hluLayerStructure.ihs_summaryColumn.Ordinal);
                    int indGeom1 = Int32.Parse(_mapInfoApp.Eval(String.Format("ColumnInfo({0}, {1}, {2})", _selName,
                        QuoteValue(lastColName), (int)MapInfoConstants.ColumnInfo.COL_INFO_NUM))) + 1;
                    int indGeom2 = indGeom1 + 1;
                    string geomCol1Name = _mapInfoApp.Eval(String.Format("ColumnInfo({0}, {1}, {2})", _selName,
                        QuoteValue(String.Format("Col{0}", indGeom1)), (int)MapInfoConstants.ColumnInfo.COL_INFO_NAME));
                    string geomCol2Name = _mapInfoApp.Eval(String.Format("ColumnInfo({0}, {1}, {2})", _selName,
                        QuoteValue(String.Format("Col{0}", indGeom2)), (int)MapInfoConstants.ColumnInfo.COL_INFO_NAME));

                    // Get the geometry values from the layer
                    double geom1;
                    double geom2;
                    GetGeometryInfo(_selName, out geom1, out geom2);

                    // Update the geometry values of the new merged feature
                    _mapInfoApp.Do(String.Format("Update {0} Set {1} = {3}, {2} = {4} Where RowID = {5}",
                        _selName, QuoteIdentifier(geomCol1Name), QuoteIdentifier(geomCol2Name), geom1, geom2, 1));
                    //---------------------------------------------------------------------

                    CollectHistory(ixGeom1, ixGeom2, readCommandTemplate, historyColumnNames, ref historyTable);
                }

                if (CommitChanges())
                    return historyTable;
                else
                    return null;
            }
            catch
            {
                if (rollbackChanges) 
                    _mapInfoApp.RunCommand(String.Format("Rollback Table {0}", _hluLayer));
                return null;
            }
        }

        /// <summary>
        /// Merges the features logically.
        /// </summary>
        /// <param name="keepIncid">The incid to keep.</param>
        /// <param name="historyColumns">The history columns.</param>
        /// <returns></returns>
        public override DataTable MergeFeaturesLogically(string keepIncid, DataColumn[] historyColumns)
        {
            string[] indexColumns = null;

            try
            {
                if (!TableExists(_selName)) return null;

                DataTable historyTable = CreateHistoryTable(_selName, true, historyColumns);

                int ixGeom1 = historyColumns.Length;
                int ixGeom2 = historyColumns.Length + 1;

                string[] historyColumnNames = historyTable.Columns.Cast<DataColumn>().Select(c => c.Ordinal < ixGeom1 ?
                    GetFieldName(_hluLayerStructure.Columns[c.ColumnName].Ordinal) : c.ColumnName).ToArray();

                string numFormat = String.Format("D{0}", _hluLayerStructure.toid_fragment_idColumn.MaxLength);

                string fetchCommand = String.Format("Fetch Rec {0} From {1}", "{0}", _selName);
                string incidCommand = String.Format("{0}.{1}", _selName,
                    GetFieldName(_hluLayerStructure.incidColumn.Ordinal));
                string toidCommand = String.Format("{0}.{1}", _selName,
                    GetFieldName(_hluLayerStructure.toidColumn.Ordinal));
                string readCommandTemplate = String.Format("{0}.{1}", _selName, "{0}");

                string incidFieldName = GetFieldName(_hluLayerStructure.incidColumn.Ordinal);
                string toidFieldName = GetFieldName(_hluLayerStructure.toidColumn.Ordinal);
                string toidFragFieldName = GetFieldName(_hluLayerStructure.toid_fragment_idColumn.Ordinal);

                int numRows = Int32.Parse(_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})",
                    _selName, (int)MapInfoConstants.TableInfo.TAB_INFO_NROWS)));

                string updateCommandTemplate;
                if (_hluFieldNames.Length > 2)
                {
                    updateCommandTemplate = String.Format("Update {0} Set {1} = {2}", _selName, incidFieldName, QuoteValue(keepIncid));

                    for (int i = 1; i <= numRows; i++)
                    {
                        _mapInfoApp.RunCommand(String.Format(fetchCommand, i));

                        if (_mapInfoApp.Eval(incidCommand) == keepIncid)
                        {
                            //---------------------------------------------------------------------
                            // FIX: 013 Don't overwrite the geometry fields during logical merge
                            // Don't include the geometry fields in the update string when updating
                            // all the rows that are logically merging into the keep incid.
                            updateCommandTemplate = new StringBuilder(updateCommandTemplate).Append(_hluFieldNames.Where(fn =>
                                (fn != incidFieldName) && (fn != toidFieldName) && (fn != toidFragFieldName) && !fn.StartsWith("shape_")).Aggregate(new StringBuilder(), 
                                (sb, fn) => sb.Append(", " + fn + " = " + QuoteValue(_mapInfoApp.Eval(String.Format(readCommandTemplate, fn))))))
                                    .Append(" Where RowID = {0}").ToString();
                            //---------------------------------------------------------------------
                            break;
                        }
                    }
                }
                else
                {
                    updateCommandTemplate = String.Format("Update {0} Set {1} = {2} Where RowID = {3}",
                        _selName, incidFieldName, QuoteValue(keepIncid), "{0}");
                }

                //---------------------------------------------------------------------
                // FIX: 014 Don't drop indexes unless processing > 500 rows
                // Dropping (and rebuilding) the indexes takes a long time
                // so it is not worth it unless a lot of rows are going to
                // be processed.
                //
                //if (numRows > 2)
                if (numRows > 500)
                //---------------------------------------------------------------------
                {
                    indexColumns = new string[] { 
                        GetFieldName(_hluLayerStructure.incidColumn.Ordinal),
                        GetFieldName(_hluLayerStructure.toidColumn.Ordinal),
                        GetFieldName(_hluLayerStructure.toid_fragment_idColumn.Ordinal) };
                    DropIndexes(_hluLayer, indexColumns);
                }

                for (int i = 1; i <= numRows; i++)
                {
                    _mapInfoApp.RunCommand(String.Format(fetchCommand, i));
                    if (_mapInfoApp.Eval(incidCommand) != keepIncid)
                    {
                        CollectHistory(ixGeom1, ixGeom2, readCommandTemplate, historyColumnNames, ref historyTable);
                        _mapInfoApp.Do(String.Format(updateCommandTemplate, i));
                    }
                }

                if (historyTable.Rows.Count == 0) historyTable = null;

                if ((historyTable != null) && !CommitChanges())
                    historyTable = null;

                return historyTable;
            }
            catch
            {
                _mapInfoApp.Do(String.Format("Rollback Table {0}", _hluLayer));
                return null;
            }
            finally { if (indexColumns != null) CreateIndexes(_hluLayer, indexColumns); }
        }

        public override bool Export(string tempMdbPathName, string attributeDatasetName, int attributesLength, bool selectedOnly)
        {
            int outFeatureCount = 0;
            string attributeTable = String.Empty;
            string outTable = String.Empty;
            string selTable = String.Empty;
            string joinLayer = String.Empty;
            string[] indexColumns = null;

            try
            {
                if (!File.Exists(tempMdbPathName)) throw new IOException("File not found");

                //---------------------------------------------------------------------
                // FIX: 035 Only export selected features, not all features for
                // selected incids.
                //
                // Get the number of features to export.
                if (selectedOnly)
                {
                    // Get the number of features selected.
                    outFeatureCount = Int32.Parse(_mapInfoApp.Eval(String.Format(
                        "SelectionInfo({0})", (int)MapInfoConstants.SelectionInfo.SEL_INFO_NROWS)));

                    // Get the name of the table that the current selection is based on.
                    if (!TableExists(_selName)) _selName = _mapInfoApp.Eval(String.Format(
                        "SelectionInfo({0})", (int)MapInfoConstants.SelectionInfo.SEL_INFO_TABLENAME));

                    // If there is no selection cancel the export.
                    if (string.IsNullOrEmpty(_selName) ||
                        (outFeatureCount == 0))
                    {
                        MessageBox.Show("Export cancelled. No features selected.",
                            "HLU: Export", MessageBoxButton.OK, MessageBoxImage.Error);
                        return false;
                    }
                }
                else
                {
                    // Get the total number of features in the table.
                    outFeatureCount = Int32.Parse(_mapInfoApp.Eval(String.Format(
                        "TableInfo({0}, {1})", QuoteValue(_hluLayer), (int)MapInfoConstants.TableInfo.TAB_INFO_NROWS)));
                }
                //---------------------------------------------------------------------

                //---------------------------------------------------------------------
                // FIX: 039 Check export layer won't exceed MapInfo maximum record
                // length or file sizes.
                //
                // Check the total attributes length doesn't exceed the MapInfo
                // maximum record length.
                if (attributesLength > Settings.Default.MapInfoMaxRecordLength)
                    throw new Exception(String.Format("The export format record length ({0} bytes) exceeds the maximum allowed for MapInfo (4000 bytes)", attributesLength));

                // Check the total export layer won't exceed the MapInfo
                // maximum .tab file size.
                if ((attributesLength * outFeatureCount)/1024 > Settings.Default.MapInfoMaxTableSize)
                    throw new Exception(String.Format("The export table size ({0} Kb) will exceed the maximum allowed for MapInfo .tab files (2 Gb)", (attributesLength * outFeatureCount)/1024));
                //---------------------------------------------------------------------

                //---------------------------------------------------------------------
                // FIX: 040 Enable MapInfo users to set a default export
                // folder path.
                //
                // Prompt the user for where to save the export layer
                SaveFileDialog saveFileFlg = new SaveFileDialog();
                saveFileFlg.Title = "HLU Export";
                saveFileFlg.ValidateNames = true;
                saveFileFlg.OverwritePrompt = true;
                saveFileFlg.RestoreDirectory = false;
                saveFileFlg.InitialDirectory = Settings.Default.ExportPath;
                saveFileFlg.Filter = "MapInfo Tables (*.tab)|*.tab";
                //---------------------------------------------------------------------

                // If no export dataset name was chosen by the user then cancel the export.
                if (saveFileFlg.ShowDialog() != true)
                {
                    MessageBox.Show("Export cancelled. No output table selected.",
                        "HLU: Export", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }

                //---------------------------------------------------------------------
                // FIX: 035 Only export selected features, not all features for
                // selected incids.
                //
                // Get the name of the feature layer to join to.
                if (selectedOnly)
                {
                    // Set the path and name of the temporary table.
                    string tempTabPath = Regex.Replace(saveFileFlg.FileName, ".tab", "_tmp.tab", RegexOptions.IgnoreCase);

                    // Set the name of the temporary table (without the path or extension).
                    string tempTabName = Path.GetFileNameWithoutExtension(tempTabPath);

                    // Close the new temporary table if it is already open.
                    if (TableExists(tempTabName))
                        _mapInfoApp.Do(String.Format("Close Table {0}", QuoteIdentifier(tempTabName)));

                    // Save the selection as a new temporary table.
                    _mapInfoApp.Do("Set ProgressBars Off");
                    _mapInfoApp.Do(String.Format("Commit Table {0} As {1}",
                        QuoteIdentifier(_selName), QuoteValue(tempTabPath)));
                    _mapInfoApp.Do("Set ProgressBars On");
                    //---------------------------------------------------------------------

                    // Close the new temporary selection table.
                    _mapInfoApp.Do(String.Format("Close Table {0}", QuoteValue(_selName)));

                    // Re-open the new temporary table (so that it can be used).
                    _mapInfoApp.Do(String.Format("Open Table {0}", QuoteValue(tempTabPath)));

                    // Get the name of the new temporary table
                    string tempTable = _mapInfoApp.Eval(String.Format("TableInfo(0, {0})",
                        (int)MapInfoConstants.TableInfo.TAB_INFO_NAME));

                    // Set the name of the join table to the new temporary table.
                    joinLayer = tempTable;
                }
                else
                {
                    // Close the previous temporary selection table.
                    CloseTable(ref _selName);

                    // Set the name of the join table to the whole feature layer.
                    joinLayer = _hluLayer;
                }
                //---------------------------------------------------------------------

                // Set the name of the export file.
                string outTabPath = saveFileFlg.FileName;

                //Register (build) a MapInfo table from the Access attribute table
                _mapInfoApp.Do(String.Format("Register Table {0} Type {1} Table {2}",
                    QuoteValue(tempMdbPathName), QuoteValue("Access"), QuoteValue(attributeDatasetName)));

                // Open the registered attribute table
                _mapInfoApp.Do(String.Format("Open Table {0}", QuoteValue(tempMdbPathName)));

                // Get the name of the opened attribute table
                attributeTable = _mapInfoApp.Eval(String.Format("TableInfo(0, {0})",
                    (int)MapInfoConstants.TableInfo.TAB_INFO_NAME));

                //---------------------------------------------------------------------
                // CHANGED: CR13 (Export features performance)
                // Drop any unnecessary indexes on the join table to improve
                // performance.
                //
                // Get a list of the columns with indexes, excluding the
                // incid column.
                indexColumns = ColumnIndexList(joinLayer, new string[] { _hluLayerStructure.incidColumn.ColumnName });

                // Drop the unneccessary indexes from the join table.
                DropIndexes(joinLayer, indexColumns);
                //---------------------------------------------------------------------

                // Check if the geometry fields are in the join layer.
                string geomCol1Name = "Shape_length";
                string geomCol2Name = "Shape_area";
                string geomColNames = ColumnList(joinLayer, new string[] { geomCol1Name, geomCol2Name });

                bool geomColsFound = false;
                if (!String.IsNullOrEmpty(geomColNames)) geomColsFound = true;

                // Perform a sql join between the feature layer and the attribute table
                //---------------------------------------------------------------------
                // FIX: 026 Hide progress bars during MapInfo processing
                // FIX: 037 Move the geometry length and area fields to the end.
                _mapInfoApp.Do("Set ProgressBars Off");
                _mapInfoApp.Do(String.Format("Select {0}, {1}{2} From {3}, {4} Where {3}.{5} = {4}.{6}",
                    ColumnList(joinLayer, new string[] { geomCol1Name, geomCol2Name }, false),
                    ColumnList(attributeTable, new string[] { _hluLayerStructure.incidColumn.ColumnName }, false),
                    geomColsFound ? string.Format(", {0}", geomColNames) : string.Empty,
                    QuoteIdentifier(joinLayer),
                    QuoteIdentifier(attributeTable), 
                    QuoteIdentifier(GetFieldName(_hluLayerStructure.incidColumn.Ordinal)),
                    QuoteIdentifier(_hluLayerStructure.incidColumn.ColumnName)));
                _mapInfoApp.Do("Set ProgressBars On");
                //---------------------------------------------------------------------
 
                // Get the name of the new selection table
                selTable = _mapInfoApp.Eval(String.Format("SelectionInfo({0})",
                    (int)MapInfoConstants.SelectionInfo.SEL_INFO_SELNAME));

                // If the join failed (i.e. no joined features existed) then abort the export
                if (string.IsNullOrEmpty(selTable)) return false;

                // Save the joined table as the new export table
                //---------------------------------------------------------------------
                // FIX: 026 Hide progress bars during MapInfo processing
                _mapInfoApp.Do("Set ProgressBars Off");
                _mapInfoApp.Do(String.Format("Commit Table {0} As {1}",
                    QuoteIdentifier(selTable), QuoteValue(outTabPath)));
                _mapInfoApp.Do("Set ProgressBars On");
                //---------------------------------------------------------------------

                // Close the exported table.
                _mapInfoApp.Do(String.Format("Close Table {0}", QuoteValue(selTable)));

                // Re-open the exported table (so that it can be updated)
                _mapInfoApp.Do(String.Format("Open Table {0}", QuoteValue(outTabPath)));

                // Get the name of the exported table
                outTable = _mapInfoApp.Eval(String.Format("TableInfo(0, {0})",
                    (int)MapInfoConstants.TableInfo.TAB_INFO_NAME));

                //---------------------------------------------------------------------
                // CHANGED: CR13 (Export features performance)
                // Only update the geometry columns for all exported rows if new
                // geometry columns need to be added to the table because they were
                // not found earlier.
                if (!geomColsFound)
                {
                    //---------------------------------------------------------------------
                    // FIX: 026 Hide progress bars during MapInfo processing
                    _mapInfoApp.Do("Set ProgressBars Off");
                    //---------------------------------------------------------------------

                    // Set the geometry column names to the correct case (as
                    // best as can be determined).
                    if (HasColumn(joinLayer, GetFieldName(_hluLayerStructure.incidColumn.Ordinal).ToUpper(), false))
                    {
                        geomCol1Name = "SHAPE_LENGTH";
                        geomCol2Name = "SHAPE_AREA";
                    }
                    else if (HasColumn(joinLayer, GetFieldName(_hluLayerStructure.incidColumn.Ordinal).ToLower(), false))
                    {
                        geomCol1Name = "shape_length";
                        geomCol2Name = "shape_area";
                    }

                    int doubleTypeInt;
                    _typeMapSystemToSQL.TryGetValue(typeof(double), out doubleTypeInt);
                    string doubleMIType = ((MapInfoConstants.ColumnType)doubleTypeInt).ToString().Replace("COL_TYPE_", "");
                    _mapInfoApp.Do(String.Format("Alter Table {0} (Add {1} {3}, {2} {3})",
                        outTable, QuoteIdentifier(geomCol1Name), QuoteIdentifier(geomCol2Name), doubleMIType));

                    // Count the number of rows in the export table
                    int numRows = Int32.Parse(_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})",
                        outTable, (int)MapInfoConstants.TableInfo.TAB_INFO_NROWS)));

                    // Speed up edits ('Undo Off' and 'FastEdit On').
                    _mapInfoApp.Do(String.Format("Set Table {0} Undo Off", QuoteIdentifier(outTable)));
                    _mapInfoApp.Do(String.Format("Set Table {0} FastEdit On", QuoteIdentifier(outTable)));

                    // Fetch the first row from the export table
                    double geom1;
                    double geom2;
                    string fetchCommand = String.Format("Fetch Next From {0}", outTable);
                    _mapInfoApp.Do(String.Format("Fetch First From {0}", outTable));

                    // Loop through all the export table rows updating their geometry
                    for (int i = 1; i <= numRows; i++)
                    {
                        GetGeometryInfo(outTable, out geom1, out geom2);

                        _mapInfoApp.Do(String.Format("Update {0} Set {1} = {3}, {2} = {4} Where RowID = {5}",
                            outTable, QuoteIdentifier(geomCol1Name), QuoteIdentifier(geomCol2Name), geom1, geom2, i));

                        _mapInfoApp.Do(fetchCommand);
                    }

                    // Save the export table updates
                    _mapInfoApp.Do(String.Format("Commit Table {0}", QuoteIdentifier(outTable)));

                    // Set undo back on and fastedit off again.
                    _mapInfoApp.Do(String.Format("Set Table {0} Undo On", QuoteIdentifier(outTable)));
                    _mapInfoApp.Do(String.Format("Set Table {0} FastEdit Off", QuoteIdentifier(outTable)));

                    //---------------------------------------------------------------------
                    // FIX: 026 Hide progress bars during MapInfo processing
                    _mapInfoApp.Do("Set ProgressBars On");
                    //---------------------------------------------------------------------
                }
                //---------------------------------------------------------------------

                //---------------------------------------------------------------------
                // CHANGED: CR16 (Adding exported features)
                // Ask the user if they want to add the new export table to the
                // current map window (i.e. keep it open).
                if (TableExists(outTable))
                {
                    MessageBoxResult userResponse = MessageBoxResult.No;
                    userResponse = MessageBox.Show("The export operation succeeded.\n\nAdd the exported layer to the current map?", "HLU: Export",
                        MessageBoxButton.YesNo, MessageBoxImage.Question);
                    // Add the already open layer to the map window.
                    if (userResponse == MessageBoxResult.Yes)
                    {
                        _mapInfoApp.Do(String.Format("Add Map Window {0} Layer {1}", _hluMapWindowID, outTable));
                    }
                    else
                        // Close the export table.
                        _mapInfoApp.Do(String.Format("Close Table {0}", QuoteIdentifier(outTable)));
                }
                else
                    throw new Exception("The export operation failed.");
                //---------------------------------------------------------------------
                return true;
            }
            catch (Exception ex)
            {
                // Close the export table (if it exists).
                if (TableExists(outTable))
                    _mapInfoApp.Do(String.Format("Close Table {0}", QuoteIdentifier(outTable)));

                MessageBox.Show("Export failed. The error message returned was:\n\n" + ex.Message,
                    "HLU: Export", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
            finally
            {
                // Close the temporary attribute, temporary selection tables
                // (if they exist).
                if (TableExists(attributeTable))
                    _mapInfoApp.Do(String.Format("Close Table {0}", QuoteIdentifier(attributeTable)));
                if (TableExists(selTable))
                    _mapInfoApp.Do(String.Format("Close Table {0}", QuoteIdentifier(selTable)));

                if (selectedOnly)
                {
                    // Close the temporary selection table.
                    if (TableExists(joinLayer))
                        _mapInfoApp.Do(String.Format("Drop Table {0}", QuoteValue(joinLayer)));
                }
                else
                {
                    // Reinstate any indexes removed from the hlu layer.
                    if (indexColumns != null) CreateIndexes(_hluLayer, indexColumns);
                }

            }
        }

        public override DataTable UpdateFeatures(DataColumn[] updateColumns, object[] updateValues, 
            DataColumn[] historyColumns)
        {
            try
            {
                if (!TableExists(_selName)) return null;

                DataTable historyTable = CreateHistoryTable(_selName, true, historyColumns);
                ReadSelectedRows(false, false, false, true, _selName, ref historyTable);
                return UpdateFeaturesAction(updateColumns, updateValues, historyTable);
            }
            catch { return null; }
        }

        /// <summary>
        /// Updates chosen attributes of map features identified by a where clause. The update values are 
        /// the same for all updated rows.
        /// </summary>
        /// <param name="tempMdbPathName">Path to Mdb containing table of id values 
        /// identifying the features to be updated.</param>
        /// <param name="selectionTableName">Table of id values identifying the features to be updated.</param>
        /// <param name="selectionWhereClause">Where clause identifying the features to be updated.</param>
        /// <param name="updateColumns">The columns to be updated. Column names and data types (per type maps) 
        /// match those of GIS layer.</param>
        /// <param name="updateValues">The values to be assigned to the update columns of each updated row.</param>
        /// <param name="historyColumns">The columns of the history table returned by the method.</param>
        /// <returns>A DataTable with the columns in the historyColumns parameter containing the attributes of the 
        /// updated features prior to the update.</returns>
        public override DataTable UpdateFeatures(DataColumn[] updateColumns, object[] updateValues, 
            DataColumn[] historyColumns, string tempMdbPathName, string selectionTableName)
        {
            try
            {
                return UpdateFeaturesAction(updateColumns, updateValues, 
                    SqlSelect(tempMdbPathName, selectionTableName, historyColumns));
            }
            catch { return null; }
        }

        /// <summary>
        /// Updates chosen attributes of map features identified by a where clause. The update values are 
        /// the same for all updated rows.
        /// </summary>
        /// <param name="selectionWhereClause">Where clause identifying the features to be updated.</param>
        /// <param name="updateColumns">The columns to be updated. Column names and data types (per type maps) 
        /// match those of GIS layer.</param>
        /// <param name="updateValues">The values to be assigned to the update columns of each updated row.</param>
        /// <param name="historyColumns">The columns of the history table returned by the method.</param>
        /// <returns>A DataTable with the columns in the historyColumns parameter containing the attributes of the 
        /// updated features prior to the update.</returns>
        public override DataTable UpdateFeatures(DataColumn[] updateColumns, object[] updateValues, 
            DataColumn[] historyColumns, List<SqlFilterCondition> selectionWhereClause)
        {
            try
            {
                //---------------------------------------------------------------------
                // FIXED: KI99 (Modification History)
                // Read the geometry data columns in the selected rows and return
                // them in the datatable which is then used to update the history
                // table so that they are recorded in the history.
                return UpdateFeaturesAction(updateColumns, updateValues, 
                    SqlSelect(false, true, historyColumns, selectionWhereClause));
                //---------------------------------------------------------------------
            }
            catch { return null; }
        }

        private DataTable UpdateFeaturesAction(DataColumn[] updateColumns, 
            object[] updateValues, DataTable historyTable)
        {
            if (!TableExists(_selName))
                throw new Exception("Error selecting update features from MapInfo layer.");

            //---------------------------------------------------------------------
            // QUERY: KI106 (Shape area and length values)
            // Should this include updates for the geom1 and geom2 columns automatically?
            //---------------------------------------------------------------------

            // update selection
            _mapInfoApp.Do(String.Format("Update {0} Set {1}", _selName,
                String.Join(",", updateColumns.Select((c, index) => String.Format("{0} = {1}",
                   GetFieldName(_hluLayerStructure.Columns[c.ColumnName].Ordinal),
                   QuoteValue(updateValues[index]))).ToArray())));

            if (!CommitChanges())
                return null;
            else
                return historyTable;
        }

        public override string WhereClause(bool includeWhere, bool quoteIdentifiers, 
            bool qualifyColumns, List<SqlFilterCondition> whereConds)
        {
            return _whereNullRegex.Replace(base.WhereClause(includeWhere, quoteIdentifiers, 
                qualifyColumns, whereConds), WhereClauseNullReplacer);
        }

        private string WhereClauseNullReplacer(Match m)
        {
            if (!String.IsNullOrEmpty(m.Groups["isnot"].Value) && !String.IsNullOrEmpty(m.Groups["null"].Value))
                return String.Format(" <> {0}{0}", StringLiteralDelimiter);

            if (!String.IsNullOrEmpty(m.Groups["is"].Value) && !String.IsNullOrEmpty(m.Groups["null"].Value))
                return String.Format(" = {0}{0}", StringLiteralDelimiter);

            if (!String.IsNullOrEmpty(m.Groups["null"].Value))
                return String.Format(" {0}{0}", StringLiteralDelimiter);

            return m.Value;
        }

        private bool CommitChanges()
        {
            //---------------------------------------------------------------------
            // FIX: 026 Hide progress bars during MapInfo processing
            _mapInfoApp.Do("Set ProgressBars Off");
            _mapInfoApp.Do(String.Format("Commit Table {0}", _hluLayer));
            _mapInfoApp.Do("Set ProgressBars On");
            //---------------------------------------------------------------------

            // if _hluLayer still flagged as edited save must have failed
            if (_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})", _hluLayer,
                (int)MapInfoConstants.TableInfo.TAB_INFO_EDITED)) == "T")
            {
                _mapInfoApp.Do(String.Format("Rollback Table {0}", _hluLayer));
                return false;
            }
            else
            {
                return true;
            }
        }
        
        protected override string GetFieldName(int columnOrdinal)
        {
            return columnOrdinal > -1 && columnOrdinal < _hluFieldMap.Length ? _hluFieldNames[_hluFieldMap[columnOrdinal] - 1] : null;
        }

        public override bool Start(ProcessWindowStyle windowStyle)
        {
            try
            {
                // get any running MapInfo processes
                _mapInfoProcsPreStart = GetMapInfoProcesses();

                // if there are already any MapInfo processes running then tell the
                // user that they must close all instanced before starting the tool
                if (_mapInfoProcsPreStart.Count() != 0)
                {
                    MessageBox.Show("MapInfo is already running.\n\nAll existing instances of MapInfo must be stopped before the tool can be launched.",
                        "Error Starting MapInfo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }

                //---------------------------------------------------------------------
                // FIXED: KI94 (MapInfo Layer Control)
                // FIXED: KI98 (MapInfo user interface)
                // Start MapInfo as a process rather than as a COM object so that it
                // starts correctly (e.g. all the menu bars, etc. where the user wants).

                // Determine the default version of MapInfo
                String miver = GetDefaultOLE_MIVer();

                // Start the default version of MapInfo
                LaunchMI(miver);

                // Connect to the running version of MapInfo
                _mapInfoApp = (DMapInfo)ConnectToRunningMI(miver);
                //---------------------------------------------------------------------

                // open the HLU workspace (returns false if it is not found or not valid)
                if (!OpenWorkspace(_mapPath)) return false;

                // size MapInfo window
                Window(windowStyle, IntPtr.Zero);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Starting MapInfo", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        /// <summary>
        /// Get the default OLE version of MapInfo from the registry
        /// </summary>
        /// <returns></returns>
        public static string GetDefaultOLE_MIVer()
        {
            string ver = "";

            using (Microsoft.Win32.RegistryKey prokey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("MapInfo.Application\\CurVer"))
            {
                ver = prokey.GetValue("").ToString();
            }

            var verSplit = ver.Split('.');
            ver = verSplit[verSplit.GetUpperBound(0)];

            return ver;

        }

        /// <summary>
        /// Start a new process of MapInfo
        /// </summary>
        public void LaunchMI(string ver)
        {
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo procInfo = new System.Diagnostics.ProcessStartInfo();
            procInfo.UseShellExecute = true;
            procInfo.WindowStyle = ProcessWindowStyle.Normal;
            procInfo.FileName = "MapInfoW.exe";

            // Pass the name of an empty workspace in the same directory as the executing assembly
            // so that MapInfo opens it automatically (and hence doesn't display the quickstart
            // window.
            string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            string assemblyDir = Path.GetDirectoryName(assemblyPath);
            string wksPath = new Uri(assemblyDir).LocalPath + "\\Empty.wor";
            //procInfo.Arguments = @"""C:\Program Files (x86)\HLU\HLU GIS Tool\Empty.wor""";
            procInfo.Arguments = @"""" + wksPath + @"""";

            proc.StartInfo = procInfo;
            proc.Start();

            //int miProcID = Shell(GetMIPath(ver) + "\\mapinfow.exe", Microsoft.VisualBasic.AppWinStyle.NormalFocus, false, -1);
        }

        /// <summary>
        /// Connects to an existing instance of MapInfo
        /// </summary>
        /// <remarks>Doesn't find instances started with /Automation parameter...
        /// Also only recognises the CLSID associated with ProgID MapInfo.Application, this is overwritten by different versions
        /// every time one is run (certainly true in MI 8.5, but not in more recent versions!)</remarks>
        public static object ConnectToRunningMI(string ver)
        {
            object MIObj1 = null;
            int stoploop = 0;
            int countloop = 0;
            //---------------------------------------------------------------------
            // FIX: 002 Increase the max timeout waiting for MI to start
            // Wait up to 30 seconds for MapInfo to start.
            while (stoploop != 1 || countloop > 30)
            //---------------------------------------------------------------------
            {
                try
                {
                    MIObj1 = System.Runtime.InteropServices.Marshal.GetActiveObject("MapInfo.Application." + ver);
                    stoploop = 1;
                    //MIObj1.Visible = true;
                    //MIObj1.Do("print \"OLE connection to existing instance\" & chr$(13) & \"Hello World\"");
                }
                catch (COMException)
                {
                    if (MIObj1 == null)
                    {
                        System.Threading.Thread.Sleep(1000);
                        countloop += 1;
                    }

                }
            }

            if (MIObj1 == null)
            {
                MessageBox.Show("No Running instances of MapInfo version " + ver, "Connect to MapInfo",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }

            //---------------------------------------------------------------------
            // FIX: 019 Pause before connecting to MapInfo to let it finish opening
            // Wait a couple of seconds after connecting to MapInfo to let it finish opening.
            System.Threading.Thread.Sleep(2000);
            //---------------------------------------------------------------------
            return MIObj1;

        }

        /// <summary>
        /// Finding installed version of MI
        /// </summary>
        /// <returns>Path to latest installed version of MI</returns>
        /// <remarks></remarks>
        public static string GetMIPath(string ver)
        {
            string miPath = null;

            using (Microsoft.Win32.RegistryKey prodirKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\\MapInfo\\MapInfo\\Professional\\" + ver))
            {
                miPath = prodirKey.GetValue("ProgramDirectory").ToString() + "\\";
            }

            return miPath;

        }

        private void EnableStandardTools()
        {
            _mapInfoApp.Do("Menu Bar Show");
            _mapInfoApp.Do(String.Format("Alter ButtonPad {0} Show {1}", QuoteValue("Standard"), "Fixed"));
            _mapInfoApp.Do(String.Format("Alter ButtonPad {0} Show {1}", QuoteValue("Main"), "Float"));
            _mapInfoApp.Do(String.Format("Alter ButtonPad {0} Show {1}", QuoteValue("Drawing"), "Float"));
            _mapInfoApp.Do(String.Format("Alter ButtonPad {0} Show {1}", QuoteValue("DBMS"), "Float"));
            _mapInfoApp.Do(String.Format("Alter ButtonPad {0} Show {1}", QuoteValue("Web Services"), "Float"));
        }

        private void SizeWindow(int winID, bool forceSize)
        {
            double wh = double.Parse(_mapInfoApp.Eval(String.Format("WindowInfo({0}, {1})",
                winID, (int)MapInfoConstants.WindowInfo.WIN_INFO_HEIGHT)));
            double ww = double.Parse(_mapInfoApp.Eval(String.Format("WindowInfo({0}, {1})",
                winID, (int)MapInfoConstants.WindowInfo.WIN_INFO_WIDTH)));
            if (forceSize || (wh < 1) || (ww < 4))
                _mapInfoApp.Do(String.Format("Set Window {0} Width {1} Units {3} Height {2} Units {3}",
                   winID, 20, 15, QuoteValue("cm")));
        }

        /// <summary>
        /// Returns the MapInfo process that wasn't running already when the tool started
        /// </summary>
        /// <param name="miProcs"></param>
        /// <returns></returns>
        private Process GetMapInfoProcess(Process[] miProcs)
        {
            if (miProcs == null) return null;
            var q = Process.GetProcesses().Where(p => (Regex.IsMatch(p.ProcessName, "MapInfow", RegexOptions.IgnoreCase)) && 
                (p.MainWindowHandle != IntPtr.Zero)).Where(p => (miProcs.Count(mip => mip.Id == p.Id) == 0));

            if (q.Count() == 1)
            {
                Process miProcess = q.ElementAt(0);
                if (miProcess.MainModule.FileName == _mapInfoApp.FullName)
                    return miProcess;
            }
            return null;
        }

        private Process[] GetMapInfoProcesses()
        {
            return Process.GetProcesses().Where(p => 
                Regex.IsMatch(p.ProcessName, "MapInfow", RegexOptions.IgnoreCase)).ToArray();
        }

        #region Implementation of SqlBuilder

        public override string QuotePrefix { get { return String.Empty; } }
        
        public override string QuoteSuffix { get { return String.Empty; } }
        
        public override string StringLiteralDelimiter { get { return "\""; } }
        
        public override string DateLiteralPrefix { get { return "\""; } }
        
        public override string DateLiteralSuffix { get { return "\""; } }
        
        public override string WildcardSingleMatch { get { return "_"; } }
        
        public override string WildcardManyMatch { get { return "%"; } }

        public override string ConcatenateOperator { get { return "+"; } }

        private void PopulateTypeMaps()
        {
            _typeMapSystemToSQL = new Dictionary<Type, int>();
            _typeMapSystemToSQL.Add(typeof(System.String), (int)MapInfoConstants.ColumnType.COL_TYPE_CHAR);
            _typeMapSystemToSQL.Add(typeof(System.Decimal), (int)MapInfoConstants.ColumnType.COL_TYPE_DECIMAL);
            _typeMapSystemToSQL.Add(typeof(System.Int64), (int)MapInfoConstants.ColumnType.COL_TYPE_INTEGER);
            _typeMapSystemToSQL.Add(typeof(System.Int32), (int)MapInfoConstants.ColumnType.COL_TYPE_INTEGER);
            _typeMapSystemToSQL.Add(typeof(System.Int16), (int)MapInfoConstants.ColumnType.COL_TYPE_SMALLINT);
            _typeMapSystemToSQL.Add(typeof(System.Boolean), (int)MapInfoConstants.ColumnType.COL_TYPE_LOGICAL);
            _typeMapSystemToSQL.Add(typeof(System.Single), (int)MapInfoConstants.ColumnType.COL_TYPE_FLOAT);
            _typeMapSystemToSQL.Add(typeof(System.Double), (int)MapInfoConstants.ColumnType.COL_TYPE_FLOAT);
            _typeMapSystemToSQL.Add(typeof(System.DateTime), (int)MapInfoConstants.ColumnType.COL_TYPE_DATETIME);

            _typeMapSQLToSystem = new Dictionary<int, Type>();
            _typeMapSQLToSystem.Add((int)MapInfoConstants.ColumnType.COL_TYPE_CHAR, typeof(System.String));
            _typeMapSQLToSystem.Add((int)MapInfoConstants.ColumnType.COL_TYPE_DECIMAL, typeof(System.Decimal));
            _typeMapSQLToSystem.Add((int)MapInfoConstants.ColumnType.COL_TYPE_INTEGER, typeof(System.Int64));
            _typeMapSQLToSystem.Add((int)MapInfoConstants.ColumnType.COL_TYPE_SMALLINT, typeof(System.Int32));
            _typeMapSQLToSystem.Add((int)MapInfoConstants.ColumnType.COL_TYPE_DATE, typeof(System.DateTime));
            _typeMapSQLToSystem.Add((int)MapInfoConstants.ColumnType.COL_TYPE_LOGICAL, typeof(System.Boolean));
            _typeMapSQLToSystem.Add((int)MapInfoConstants.ColumnType.COL_TYPE_GRAPHIC, typeof(System.Byte[]));
            _typeMapSQLToSystem.Add((int)MapInfoConstants.ColumnType.COL_TYPE_FLOAT, typeof(System.Double));
            _typeMapSQLToSystem.Add((int)MapInfoConstants.ColumnType.COL_TYPE_TIME, typeof(System.DateTime));
            _typeMapSQLToSystem.Add((int)MapInfoConstants.ColumnType.COL_TYPE_DATETIME, typeof(System.DateTime));
        }

        private string AddJoinTable(string scratchMdbPath, string selectionTableName)
        {
            string selTablePath = Path.GetDirectoryName(scratchMdbPath) +
                Path.DirectorySeparatorChar + selectionTableName + ".TAB";

            _mapInfoApp.Do(String.Format("Register Table {0} Type ACCESS Table {1} Into {2}",
                QuoteValue(scratchMdbPath), QuoteValue(selectionTableName), QuoteValue(selTablePath)));

            _mapInfoApp.Do(String.Format("Open Table {0}", QuoteValue(selTablePath)));

            return _mapInfoApp.Eval(String.Format("TableInfo(0, {0})",
                (int)MapInfoConstants.TableInfo.TAB_INFO_NAME));
        }

        public override DataTable SqlSelect(string scratchMdbPath, string selectionTableName, 
            DataColumn[] targetColumns)
        {
            if (!File.Exists(scratchMdbPath) || String.IsNullOrEmpty(selectionTableName) ||
                (targetColumns == null)) return null;

            string joinTable = String.Empty;

            try
            {
                joinTable = AddJoinTable(scratchMdbPath, selectionTableName);

                // Close the previous temporary selection table.
                CloseTable(ref _selName);

                int numJoinColumns = Int32.Parse(_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})",
                    joinTable, (int)MapInfoConstants.TableInfo.TAB_INFO_NCOLS)));

                string joinTableNameQuoted = QuoteIdentifier(joinTable);
                string hluLayerNameQuoted = QuoteIdentifier(_hluLayer);

                List<string> joinCondList = new List<string>();

                for (int i = 1; i <= numJoinColumns; i++)
                {
                    string colName = _mapInfoApp.Eval(String.Format("ColumnInfo({0}, {1}, {2})", joinTable,
                        QuoteValue(String.Format("Col{0}", i)), (int)MapInfoConstants.ColumnInfo.COL_INFO_NAME));
                    joinCondList.Add(String.Format("{0}.{1} = {2}.{3}", hluLayerNameQuoted,
                        GetFieldName(_hluLayerStructure.Columns[colName].Ordinal), joinTableNameQuoted,
                        QuoteIdentifier(colName)));
                }

                _mapInfoApp.Do(String.Format("Select {0} From {1}, {2} Where {3}", _hluColumnList, 
                    hluLayerNameQuoted, joinTableNameQuoted, String.Join(" AND ", joinCondList.ToArray())));
                
                string selTable = _mapInfoApp.Eval(String.Format("SelectionInfo({0})",
                    (int)MapInfoConstants.SelectionInfo.SEL_INFO_SELNAME));

                if (string.IsNullOrEmpty(selTable)) return new DataTable();

                DataTable resultTable = new DataTable();
                foreach (DataColumn c in targetColumns)
                    resultTable.Columns.Add(new DataColumn(c.ColumnName, c.DataType));

                ReadSelectedRows(true, true, false, false, selTable, ref resultTable);

                return resultTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Map selection failed. MapInfo returned the following error message:\n\n{0}",
                    ex.Message), "HLU: Selection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return new DataTable();
            }
            finally
            {
                if (TableExists(joinTable))
                {
                    try { _mapInfoApp.Do(String.Format("Close Table {0}", joinTable)); }
                    catch { }
                }
            }
        }

        /// <summary>
        /// Joins are supported using WHERE syntax.
        /// Column names must be qualified with table name if multiple tables are joined.
        /// </summary>
        /// <param name="selectDistinct">If set to true a 'DISTINCT' clause is added to the SQL statement.</param>
        /// <param name="targetList">The target list of data columns.</param>
        /// <param name="whereConds">The SQL WHERE conditions.</param>
        /// <returns></returns>
        public override DataTable SqlSelect(bool selectDistinct, 
            DataColumn[] targetList, List<SqlFilterCondition> whereConds)
        {
            if ((_mapInfoApp == null) || (targetList == null) || (targetList.Length == 0)) 
                return new DataTable();

            try
            {
                bool qualifyColumns = false;
                bool additionalTables;
                DataTable resultTable = null;

                StringBuilder sbCommandText = new StringBuilder("SELECT ");
                sbCommandText.Append(TargetList(targetList, true, false, ref qualifyColumns, out resultTable));
                sbCommandText.Append(qualifyColumns ? FromList(true, targetList, true, ref whereConds,
                    out additionalTables) : String.Format(" FROM {0}", _hluLayer));
                sbCommandText.Append(WhereClause(true, true, false,
                    MapWhereClauseFields(_hluLayerStructure, whereConds)));

                _mapInfoApp.Do(sbCommandText.ToString());

                ReadSelectedRows(true, true, qualifyColumns, false, _hluLayer, ref resultTable);

                return resultTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Map selection failed. MapInfo returned the following error message:\n\n{0}",
                    ex.Message), "HLU: Selection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return new DataTable();
            }
        }

        /// <summary>
        /// Joins are supported using WHERE syntax.
        /// Column names must be qualified with table name if multiple tables are joined.
        /// </summary>
        /// <param name="selectDistinct">If set to true a 'DISTINCT' clause is added to the SQL statement.</param>
        /// <param name="addGeometryInfo">If set to true the geometry fields will be added to the returned data table.</param>
        /// <param name="targetList">The target list of data columns.</param>
        /// <param name="whereConds">The SQL WHERE conditions.</param>
        /// <returns></returns>
        public override DataTable SqlSelect(bool selectDistinct, bool addGeometryInfo,
            DataColumn[] targetList, List<SqlFilterCondition> whereConds)
        {
            if ((_mapInfoApp == null) || (targetList == null) || (targetList.Length == 0))
                return new DataTable();

            try
            {
                bool qualifyColumns = false;
                bool additionalTables;
                DataTable resultTable = null;

                StringBuilder sbCommandText = new StringBuilder("SELECT ");
                sbCommandText.Append(TargetList(targetList, true, false, ref qualifyColumns, out resultTable));
                sbCommandText.Append(qualifyColumns ? FromList(true, targetList, true, ref whereConds,
                    out additionalTables) : String.Format(" FROM {0}", _hluLayer));
                sbCommandText.Append(WhereClause(true, true, false,
                    MapWhereClauseFields(_hluLayerStructure, whereConds)));

                _mapInfoApp.Do(sbCommandText.ToString());

                ReadSelectedRows(true, true, qualifyColumns, addGeometryInfo, _hluLayer, ref resultTable);

                return resultTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Map selection failed. MapInfo returned the following error message:\n\n{0}",
                    ex.Message), "HLU: Selection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return new DataTable();
            }
        }

        public override DataTable SqlSelect(bool selectDistinct,
            DataTable[] targetTables, List<SqlFilterCondition> whereConds)
        {
            if ((_mapInfoApp == null) || (targetTables == null) || (targetTables.Length == 0) ||
                (targetTables[0].Columns.Count == 0)) return new DataTable();

            try
            {
                bool qualifyColumns = false;
                bool additionalTables;
                DataTable resultTable = null;
                StringBuilder sbCommandText = new StringBuilder("SELECT ");
                sbCommandText.Append(TargetList(targetTables, true, ref qualifyColumns, out resultTable));
                sbCommandText.Append(FromList(true, false, targetTables, ref whereConds, out additionalTables));
                sbCommandText.Append(WhereClause(true, true, qualifyColumns,
                    MapWhereClauseFields(_hluLayerStructure, whereConds)));

                _mapInfoApp.Do(sbCommandText.ToString());
                ReadSelectedRows(true, true, qualifyColumns, false, _hluLayer, ref resultTable);

                return resultTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Map selection failed. MapInfo returned the following error message:\n\n{0}",
                    ex.Message), "HLU: Selection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return new DataTable();
            }
        }

        private DataTable SqlSelect(bool replaceSelection, bool closePreviousSelection, DataColumn[] targetList, string tableName,
            bool addGeometryInfo, bool negateWhereClause, List<SqlFilterCondition> whereConds, DataColumn[] orderBy)
        {
            if ((_mapInfoApp == null) || (targetList == null) || (targetList.Length == 0))
                return new DataTable();

            try
            {
                bool qualifyColumns = false;
                bool additionalTables;
                DataTable resultTable = null;

                StringBuilder sbCommandText = new StringBuilder("SELECT ");
                sbCommandText.Append(TargetList(targetList, true, true, ref qualifyColumns, out resultTable));

                sbCommandText.Append(FromList(true, targetList, true, ref whereConds, out additionalTables));

                if (negateWhereClause)
                    sbCommandText.Append(" WHERE NOT (").Append(WhereClause(false, true, qualifyColumns,
                        MapWhereClauseFields(_hluLayerStructure, whereConds))).Append(")");
                else
                    sbCommandText.Append(WhereClause(true, true, qualifyColumns,
                        MapWhereClauseFields(_hluLayerStructure, whereConds)));
                
                if ((orderBy != null) && (orderBy.Length > 0))
                {
                    StringBuilder orderByColNames = orderBy.Where(c => _hluLayerStructure.Columns.Contains(c.ColumnName))
                        .Aggregate(new StringBuilder(), (sb, c) => sb.Append(",").Append(qualifyColumns ? 
                            QuoteIdentifier(c.Table.TableName) + "." : String.Empty).Append(
                            QuoteIdentifier(GetFieldName(_hluLayerStructure.Columns[c.ColumnName].Ordinal))));

                    if (orderByColNames.Length > 0)
                        sbCommandText.Append(String.Format(" ORDER BY {0}", orderByColNames.Remove(0, 1)));
                }

                _mapInfoApp.Do(sbCommandText.ToString());
                ReadSelectedRows(replaceSelection, closePreviousSelection, qualifyColumns, addGeometryInfo, tableName, ref resultTable);

                return resultTable;
            }
            catch { return null; }
        }

        //---------------------------------------------------------------------
        // CHANGED: CR12 (Select by attribute performance)
        // Calculate the approximate length of the SQL statement that will be
        // used in GIS so that it can be determined if the selection can be
        // performed using a direct query or if a table join is needed.
        //---------------------------------------------------------------------
        /// <summary>
        /// Calculate the approximate length of the resulting SQL query.
        /// </summary>
        /// <param name="targetList">The target list of data columns.</param>
        /// <param name="whereConds">The SQL WHERE conditions.</param>
        /// <returns>Integer of the approximate length of the SQL statement that will
        /// meet the where conditions.</returns>
        public override int SqlLength(DataColumn[] targetList, List<SqlFilterCondition> whereConds)
        {
            if ((_mapInfoApp == null) || (targetList == null) || (targetList.Length == 0))
                return 0;

            try
            {
                bool qualifyColumns = false;
                bool additionalTables;
                DataTable resultTable = null;

                StringBuilder sbCommandText = new StringBuilder("SELECT ");
                sbCommandText.Append(TargetList(targetList, true, false, ref qualifyColumns, out resultTable));
                sbCommandText.Append(qualifyColumns ? FromList(true, targetList, true, ref whereConds,
                    out additionalTables) : String.Format(" FROM {0}", _hluLayer));
                sbCommandText.Append(WhereClause(true, true, false,
                    MapWhereClauseFields(_hluLayerStructure, whereConds)));

                int sqlLen = sbCommandText.Length;

                return sqlLen;
            }
            catch
            {
                return 0;
            }
        }
        //---------------------------------------------------------------------

        private bool TableExists(string tableName)
        {
            if (String.IsNullOrEmpty(tableName)) return false;

            int numTables = Int32.Parse(_mapInfoApp.Eval("NumTables()"));
            string currTableName;
            for (int i = 1; i <= numTables; i++)
            {

                currTableName = _mapInfoApp.Eval(String.Format("TableInfo({0}, {1})", i,
                    (int)MapInfoConstants.TableInfo.TAB_INFO_NAME));
                if (currTableName == tableName) return true;
            }
            return false;
        }

        private void CloseTable(ref string tableName)
        {
            if (TableExists(tableName))
            {
                try { _mapInfoApp.Do(String.Format("Close Table {0}", tableName)); }
                catch { }
                tableName = null;
            }
        }

        private void ReadSelectedRows(bool replaceSelection, bool closePreviousSelection, bool qualifyColumns, 
            bool addGeometryInfo, string selName, ref DataTable resultTable)
        {

            //---------------------------------------------------------------------
            // CHANGED: CR31 (Switching between GIS layers)
            // Enable the user to switch between different HLU layers, where
            // there is more than one valid layer in the current document.
            //
            // Get the name of the table that the current selection is based on.
            string tableName = _mapInfoApp.Eval(String.Format(
                "SelectionInfo({0})", (int)MapInfoConstants.SelectionInfo.SEL_INFO_TABLENAME));
            if (!TableExists(tableName)) return;

            // Check that the table is the same as the expected table
            // and if not return an empty result table.
            if ((tableName != selName) && (tableName != _hluLayer)) return;
            //---------------------------------------------------------------------
            
            int numSelected = -1;
            string readSelName = null;
            if (replaceSelection)
            {
                // Close the previous temporary selection table.
                if (closePreviousSelection)
                    CloseTable(ref _selName);

                if ((numSelected = Int32.Parse(_mapInfoApp.Eval(String.Format(
                    "SelectionInfo({0})", (int)MapInfoConstants.SelectionInfo.SEL_INFO_NROWS)))) == 0) return;
                readSelName = _mapInfoApp.Eval(String.Format("SelectionInfo({0})",
                    (int)MapInfoConstants.SelectionInfo.SEL_INFO_SELNAME));
                _selName = readSelName;
            }
            else 
            {
                if (TableExists(selName) && ((numSelected = 
                    Int32.Parse(_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})", selName, 
                    (int)MapInfoConstants.TableInfo.TAB_INFO_NROWS)))) > 0))
                    readSelName = selName;
                else
                    return;
            }

            int nonGeomfieldCount = resultTable.Columns.Count;

            if (addGeometryInfo && ((resultTable = AddGeometryInfoColumns(resultTable)) == null)) return;
            int ixGeom1 = nonGeomfieldCount;
            int ixGeom2 = nonGeomfieldCount + 1;
            double geom1;
            double geom2;

            string[] resultTableColumnNames = resultTable.Columns.Cast<DataColumn>()
                .Select(c => c.Ordinal < nonGeomfieldCount ? (qualifyColumns ? ColumnAlias(c) :
                    GetFieldName(_hluLayerStructure.Columns[c.ColumnName].Ordinal)) : c.ColumnName).ToArray();

            _mapInfoApp.Do(String.Format("Fetch First From {0}", readSelName));

            string fetchCommand = String.Format("Fetch Next From {0}", readSelName);

            for (int i = 0; i < numSelected; i++)
            {
                object[] itemArray = new object[resultTable.Columns.Count];
                for (int j = 0; j < nonGeomfieldCount; j++)
                    itemArray[j] = _mapInfoApp.Eval(
                        String.Format("{0}.{1}", readSelName, resultTableColumnNames[j]));
                if (addGeometryInfo)
                {
                    GetGeometryInfo(readSelName, out geom1, out geom2);
                    itemArray[ixGeom1] = geom1;
                    itemArray[ixGeom2] = geom2;
                }
                resultTable.Rows.Add(itemArray);
                _mapInfoApp.Do(fetchCommand);
            }
        }

        public override string QuoteIdentifier(string identifier)
        {
            return identifier;
        }

        /// <summary>
        /// Does not escape string delimiter or other special characters.
        /// Does check if value is already quoted.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public override string QuoteValue(object value)
        {
            if (value == null) return "NULL";
            int colType;
            if (_typeMapSystemToSQL.TryGetValue(value.GetType(), out colType))
            {
                string s = value.ToString();
                switch ((MapInfoConstants.ColumnType)colType)
                {
                    case MapInfoConstants.ColumnType.COL_TYPE_CHAR:
                        if (s.Length == 0) return StringLiteralDelimiter + StringLiteralDelimiter;
                        if (!s.StartsWith(StringLiteralDelimiter)) s = StringLiteralDelimiter + s;
                        if (!s.EndsWith(StringLiteralDelimiter)) s += StringLiteralDelimiter;
                        return s;
                    case MapInfoConstants.ColumnType.COL_TYPE_DATE:
                    case MapInfoConstants.ColumnType.COL_TYPE_DATETIME:
                    case MapInfoConstants.ColumnType.COL_TYPE_TIME:
                        if (s.Length == 0) return DateLiteralPrefix + DateLiteralSuffix;
                        if (!s.StartsWith(DateLiteralPrefix)) s = DateLiteralPrefix + s;
                        if (!s.EndsWith(DateLiteralSuffix)) s += DateLiteralSuffix;
                        return s;
                    case MapInfoConstants.ColumnType.COL_TYPE_LOGICAL:
                        if (s == "true")
                            return StringLiteralDelimiter + "T" + StringLiteralDelimiter;
                        else
                            return StringLiteralDelimiter + "F" + StringLiteralDelimiter;
                    default:
                        return s;
                }
            }
            else
            {
                return value.ToString();
            }
        }

        public override string TargetList(DataColumn[] targetColumns, bool quoteIdentifiers, 
            bool checkQualify, ref bool qualifyColumns, out DataTable resultTable)
        {
            resultTable = new DataTable();

            if ((targetColumns == null) || (targetColumns.Length == 0)) return String.Empty; ;

            StringBuilder sbTargetList = new StringBuilder();

            try
            {
                if (checkQualify) qualifyColumns = QualifyColumnNames(targetColumns);

                string columnAlias;
                foreach (DataColumn c in targetColumns)
                {
                    if (qualifyColumns)
                    {
                        columnAlias = ColumnAlias(c);
                        if (quoteIdentifiers)
                            sbTargetList.Append(String.Format(",{0}.{1}", QuoteIdentifier(c.Table.TableName), 
                                QuoteIdentifier(c.ColumnName)));
                        else
                            sbTargetList.Append(String.Format(",{0}.{1}", c.Table.TableName, c.ColumnName));
                        resultTable.Columns.Add(new DataColumn(columnAlias, c.DataType));
                    }
                    else
                    {
                        if (quoteIdentifiers)
                            sbTargetList.Append(String.Format(",{0}", QuoteIdentifier(GetFieldName(c.Ordinal))));
                        else
                            sbTargetList.Append(String.Format(",{0}", GetFieldName(c.Ordinal)));
                        resultTable.Columns.Add(new DataColumn(c.ColumnName, c.DataType));
                    }
                }
                sbTargetList.Remove(0, 1);
            }
            catch { }

            return sbTargetList.ToString();
        }
        
        #endregion

        public override bool Close()
        {
            try
            {
                if (_mapInfoApp != null)
                {
                    _mapInfoApp.RunMenuCommand((int)MapInfoConstants.MenuDef95File.M_FILE_EXIT);
                    _mapInfoApp = null;
                    //_mapInfoComObj = null;
                }
                return true;
            }
            catch { return false; }
        }

        public override IntPtr hWnd { get { return _mapInfoParentWindow; } }

        public override void Window(ProcessWindowStyle windowStyle, IntPtr sideBySideWith)
        {
            if ((_mapInfoParentWindow == null) || !WinAPI.IsWindow(_mapInfoParentWindow) || !_mapInfoApp.Visible)
            {
                _mapInfoApp.Visible = true;
                _mapInfoProcess = GetMapInfoProcess(_mapInfoProcsPreStart);
                if (_mapInfoProcess == null) return;
                _mapInfoParentWindow = _mapInfoProcess.MainWindowHandle;
            }

            System.Windows.Forms.Screen mapInfoScreen = System.Windows.Forms.Screen.FromHandle(_mapInfoParentWindow);
            Rectangle mapInfoWorkingArea = mapInfoScreen.WorkingArea;

            System.Windows.Forms.Screen hluScreen = 
                System.Windows.Forms.Screen.FromHandle(Process.GetCurrentProcess().MainWindowHandle);

            WinAPI.WINDOWINFO winfo = new WinAPI.WINDOWINFO();
            winfo.cbSize = (uint)Marshal.SizeOf(winfo);
            WinAPI.GetWindowInfo(_mapInfoParentWindow, ref winfo);

            // Set the window style based on the parameter passed to the procedure
            switch (windowStyle)
            {
                case ProcessWindowStyle.Hidden:
                    _mapInfoApp.Visible = false;
                    break;
                case ProcessWindowStyle.Maximized:
                    _mapInfoApp.Visible = true;
                    if ((winfo.rcClient.Width < mapInfoScreen.WorkingArea.Width) ||
                         (winfo.rcClient.Bottom < mapInfoScreen.WorkingArea.Height))
                    {
                        WinAPI.ShowWindow(_mapInfoParentWindow, (int)WinAPI.WindowStates.SW_SHOWNORMAL);
                        WinAPI.ShowWindow(_mapInfoParentWindow, (int)WinAPI.WindowStates.SW_SHOWMAXIMIZED);
                    }
                    break;
                case ProcessWindowStyle.Minimized:
                    _mapInfoApp.Visible = true;
                    WinAPI.ShowWindow(_mapInfoParentWindow, (int)WinAPI.WindowStates.SW_SHOWMINIMIZED);
                    break;
                case ProcessWindowStyle.Normal:
                    _mapInfoApp.Visible = true;
                    if (sideBySideWith != IntPtr.Zero)
                    {
                        WinAPI.RECT sideBySideRect;
                        if (WinAPI.GetWindowRect(sideBySideWith, out sideBySideRect))
                        {
                            int gisWinWidth = hluScreen.WorkingArea.Width - sideBySideRect.Width;
                            if (gisWinWidth <= 0) return;
                            WinAPI.MoveWindow(sideBySideWith, 0, 0, sideBySideRect.Width, sideBySideRect.Height, true);
                            WinAPI.MoveWindow(_mapInfoParentWindow, sideBySideRect.Width, 0, gisWinWidth, 
                                hluScreen.WorkingArea.Height, false);
                        }
                    }
                    else
                    {
                        WinAPI.ShowWindow(_mapInfoParentWindow, (int)WinAPI.WindowStates.SW_SHOWNORMAL);
                    }
                    break;
            }
        }

        public override void Activate()
        {
            WinAPI.SetForegroundWindow(_mapInfoParentWindow);
        }

        public override bool OpenWorkspace(string path)
        {
            if (_mapInfoApp == null) return false;

            // Open the workspace passed to the function
            string browsePath = path;
            OpenWorkspaceDocument("Select HLU Workspace", ref browsePath);

            // Check the workspace is valid (i.e. it contains a valid HLU layer)
            bool ok = IsHluWorkspace();

            // If the workspace is valid and is different to the last used workspace
            // then save the new workspace in the settings
            if (ok && browsePath != path)
            {
                Settings.Default.MapPath = browsePath;
                Settings.Default.Save();
            }
            return ok;
        }

        private bool OpenWorkspaceDocument(string title, ref string path)
        {
            if (_mapInfoApp == null) return false;

            try
            {
                if (!File.Exists(path))
                {
                    OpenFileDialog openFileDlg = new OpenFileDialog();
                    openFileDlg.Filter = "MapInfo Workspaces (*.wor)|*.wor";
                    openFileDlg.Title = title;
                    openFileDlg.CheckPathExists = true;
                    openFileDlg.CheckFileExists = true;
                    openFileDlg.ValidateNames = true;
                    openFileDlg.Multiselect = false;
                    openFileDlg.RestoreDirectory = false;
                    openFileDlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    
                    if (openFileDlg.ShowDialog() == true)
                        path = openFileDlg.FileName;
                    else
                        return false;
                }

                _mapInfoApp.Do(String.Format("Run Application {0}", QuoteValue(path)));

                return true;
            }
            catch { return false; }
        }

        public override bool SaveWorkspace()
        {
            try
            {
                string path;
                SaveFileDialog saveFileDlg = new SaveFileDialog();
                saveFileDlg.Title = "Save New Map Document";
                saveFileDlg.Filter = "MapInfo Workspaces (*.wor)|*.wor";
                saveFileDlg.CheckPathExists = true;
                saveFileDlg.RestoreDirectory = false;
                saveFileDlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (saveFileDlg.ShowDialog() == true)
                {
                    path = saveFileDlg.FileName;
                    FileInfo fInfo = new FileInfo(path);
                    if (Directory.Exists(fInfo.DirectoryName))
                    {
                        _mapInfoApp.Do(String.Format("Save Workspace As {0}", QuoteValue(path)));
                        Settings.Default.MapPath = path;
                        Settings.Default.Save();
                        return true;
                    }
                }
            }
            catch { }

            return false;
        }

        public bool AddLayer(string tabName)
        {
            try
            {
                string path = null;

                if (String.IsNullOrEmpty(tabName))
                {
                    tabName = OpenTable(out path);
                    if (!IsHluLayer(tabName))
                    {
                        if (String.IsNullOrEmpty(tabName))
                        {
                            throw new Exception("No HLU table.");
                        }
                        else
                        {
                            _mapInfoApp.Do(String.Format("Close Table {0}", tabName));
                            throw new Exception(String.Format("Table {0} is not a valid HLU table.", tabName));
                        }
                    }
                }

                _hluMapWindowID = Int32.Parse(_mapInfoApp.Eval("FrontWindow()"));

                if ((_hluMapWindowID <= 0) || Int32.Parse(_mapInfoApp.Eval(String.Format(
                    "WindowInfo({0}, {1})", _hluMapWindowID, (int)MapInfoConstants.WindowInfo.WIN_INFO_TYPE))) !=
                    (int)MapInfoConstants.WindowInfoWindowTypes.WIN_MAPPER) _hluMapWindowID = -1;

                //---------------------------------------------------------------------
                // CHANGED: CR19 (Feature layer position in GIS)
                // Loop through all windows to find the window number (order) of
                // the current window.
                int numWindows = Int32.Parse(_mapInfoApp.Eval("NumWindows()"));
                int windowID;
                int windowNum = 0;

                for (int i = 1; i <= numWindows; i++)
                {
                    windowID = Int32.Parse(_mapInfoApp.Eval(String.Format("WindowID({0})", i)));

                    // If this is a mapper window
                    if (Int32.Parse(_mapInfoApp.Eval(String.Format("WindowInfo({0}, {1})", i,
                        (int)MapInfoConstants.WindowInfo.WIN_INFO_TYPE))) ==
                        (int)MapInfoConstants.WindowInfoWindowTypes.WIN_MAPPER)
                    {
                        // Increment the window number
                        windowNum += 1;

                        // If this window is the current window then use this window number (order)
                        if (windowID == _hluMapWindowID) break;
                    }
                }
                //---------------------------------------------------------------------

                if (_hluMapWindowID != -1)
                {
                    _mapInfoApp.Do(String.Format("Add Map Layer {0}", tabName));
                }
                else
                {
                    _mapInfoApp.Do(String.Format("Map From {0}", tabName));
                    _hluMapWindowID = Convert.ToInt32(_mapInfoApp.Eval("FrontWindow()"));
                    SizeWindow(_hluMapWindowID, true);
                }

                //---------------------------------------------------------------------
                // CHANGED: CR19 (Feature layer position in GIS)
                // Set the current HLU layer to the new layer.
                _hluCurrentLayer = new GISLayer(windowNum, 0, tabName);
                //---------------------------------------------------------------------

                _mapInfoApp.Do(String.Format("Set Window {0} SysMenuClose Off", _hluMapWindowID));

                if (!String.IsNullOrEmpty(path)) return SaveWorkspace();
                return true;
            }
            catch (Exception ex)
            {
                _hluMapWindowID = -1;
                MessageBox.Show(ex.Message, "Error Opening Table", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        private string OpenTable(out string path)
        {
            string tabName;
            path = null;

            OpenFileDialog openFileDlg = new OpenFileDialog();
            openFileDlg.Filter = "MapInfo Tables (*.tab)|*.tab";
            openFileDlg.CheckPathExists = true;
            openFileDlg.CheckFileExists = true;
            openFileDlg.ValidateNames = true;
            openFileDlg.Multiselect = false;
            openFileDlg.RestoreDirectory = false;
            openFileDlg.InitialDirectory =
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (openFileDlg.ShowDialog() != true)
            {
                _hluMapWindowID = -1;
                _hluLayerStructure = null;
                _hluFieldNames = null;
                _hluFieldMap = null;
                _hluLayer = null;
                return null;
            }
            else
            {
                path = openFileDlg.FileName;
                tabName = _mapInfoApp.Eval(String.Format("PathToTableName$({0})", QuoteValue(path)));
                _mapInfoApp.Do(String.Format("Open Table {0}", QuoteValue(path)));
                tabName = _mapInfoApp.Eval(String.Format("TableInfo(0, {0})",
                    (int)MapInfoConstants.TableInfo.TAB_INFO_NAME));
                //_hluLayer = tabName;
                return tabName;
            }
        }

        /// <summary>
        /// List the columns in the table if they are not found in
        /// the list of columns to skip.
        /// </summary>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="skipColumns">The list of columns to skip over if found.</param>
        /// <param name="includeGeom">If set to <c>true</c> include the geometry field.</param>
        /// <returns></returns>
        private string ColumnList(string tableName, string[] skipColumns, bool includeGeom)
        {
            if (skipColumns == null) skipColumns = new string[0];

            int numColumns = Int32.Parse(_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})",
                tableName, (int)MapInfoConstants.TableInfo.TAB_INFO_NCOLS)));

            List<string> columnList = new List<string>();

            string tableNameQuoted = QuoteIdentifier(tableName);

            for (int i = 1; i <= numColumns; i++)
            {
                string colName = _mapInfoApp.Eval(String.Format("ColumnInfo({0}, {1}, {2})", tableName,
                    QuoteValue(String.Format("Col{0}", i)), (int)MapInfoConstants.ColumnInfo.COL_INFO_NAME));

                //---------------------------------------------------------------------
                // FIX: 033 Ignore case in field names during export to avoid duplicate
                // fields.
                if (Array.FindIndex(skipColumns, c => c.Equals(colName, StringComparison.InvariantCultureIgnoreCase)) == -1)
                    columnList.Add(tableNameQuoted + "." + QuoteIdentifier(colName));
                //---------------------------------------------------------------------
            }

            if (includeGeom &&( _mapInfoApp.Eval(String.Format("TableInfo({0}, {1})", _hluLayer,
                (int)MapInfoConstants.TableInfo.TAB_INFO_MAPPABLE)) == "T"))
                return String.Format("{0}, {1}.obj", String.Join(", ", columnList.ToArray()), tableNameQuoted);
            else
                return String.Join(", ", columnList.ToArray());
        }

        /// <summary>
        /// List the columns in the table if they are found in the
        /// list of required columns.
        /// </summary>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="onlyColumns">The list of columns to include if found.</param>
        /// <returns></returns>
        private string ColumnList(string tableName, string[] onlyColumns)
        {
            if (onlyColumns == null) return string.Empty;

            int numColumns = Int32.Parse(_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})",
                tableName, (int)MapInfoConstants.TableInfo.TAB_INFO_NCOLS)));

            List<string> columnList = new List<string>();

            string tableNameQuoted = QuoteIdentifier(tableName);

            for (int i = 1; i <= numColumns; i++)
            {
                string colName = _mapInfoApp.Eval(String.Format("ColumnInfo({0}, {1}, {2})", tableName,
                    QuoteValue(String.Format("Col{0}", i)), (int)MapInfoConstants.ColumnInfo.COL_INFO_NAME));

                // Ignore case in field names during export to ensure
                // fields are found.
                if (Array.FindIndex(onlyColumns, c => c.Equals(colName, StringComparison.InvariantCultureIgnoreCase)) != -1)
                    columnList.Add(tableNameQuoted + "." + QuoteIdentifier(colName));
            }

            return String.Join(", ", columnList.ToArray());
        }

        private bool HasColumn(string tableName, string columnName, bool ignoreCase)
        {
            int numColumns = Int32.Parse(_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})",
                tableName, (int)MapInfoConstants.TableInfo.TAB_INFO_NCOLS)));

            for (int i = 1; i <= numColumns; i++)
            {
                //---------------------------------------------------------------------
                // FIX: 033 Ignore case in field names during export to avoid duplicate
                // fields.
                string colName = _mapInfoApp.Eval(String.Format("ColumnInfo({0}, {1}, {2})",
                    tableName, QuoteValue(String.Format("Col{0}", i)),
                    (int)MapInfoConstants.ColumnInfo.COL_INFO_NAME));
                if (ignoreCase && colName.ToLower() == columnName.ToLower())
                    return true;
                else if (!ignoreCase && colName == columnName)
                    return true;
                //---------------------------------------------------------------------
            }

            return false;
        }

        /// <summary>
        /// Create a list of all the index columns in the table if they
        /// are not found in the list of columns to skip.
        /// </summary>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="skipColumns">The list of columns to skip over if found.</param>
        /// <returns></returns>
        private string[] ColumnIndexList(string tableName, string[] skipColumns)
        {
            if (skipColumns == null) skipColumns = new string[0];

            int numColumns = Int32.Parse(_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})",
                tableName, (int)MapInfoConstants.TableInfo.TAB_INFO_NCOLS)));

            List<string> columnList = new List<string>();

            for (int i = 1; i <= numColumns; i++)
            {
                string colName = _mapInfoApp.Eval(String.Format("ColumnInfo({0}, {1}, {2})", tableName,
                    QuoteValue(String.Format("Col{0}", i)), (int)MapInfoConstants.ColumnInfo.COL_INFO_NAME));

                string colIndexed = _mapInfoApp.Eval(String.Format("ColumnInfo({0}, {1}, {2})", tableName,
                    QuoteValue(String.Format("Col{0}", i)), (int)MapInfoConstants.ColumnInfo.COL_INFO_INDEXED));

                // Add the name of the column to the list if it is indexed
                // and not in the list of columns to skip.
                if ((colIndexed == "T") && (Array.FindIndex(skipColumns, c => c.Equals(colName, StringComparison.InvariantCultureIgnoreCase)) == -1))
                    columnList.Add(QuoteIdentifier(colName));
            }

            return columnList.ToArray();
        }

        /// <summary>
        /// Checks whether the current workspace contains an HLU layer. Also initializes the fields 
        /// _hluMapWindowID and, indirectly (by calling IsHluLayer()), _hluFieldNames and _hluFieldMap.
        /// </summary>
        /// <returns>True if the current workspace contains a valid HLU layer, otherwise false.</returns>
        protected override bool IsHluWorkspace()
        {
            if (_mapInfoApp == null) return false;

            try
            {
                // check if the workspace contains a mapper window with a valid HLU layer
                int numWindows = Int32.Parse(_mapInfoApp.Eval("NumWindows()"));
                int windowID;
                int windowNum = 0;
                string layer;

                // Loop through all windows
                for (int i = 1; i <= numWindows; i++)
                {
                    windowID = Int32.Parse(_mapInfoApp.Eval(String.Format("WindowID({0})", i)));

                    // If this is a mapper window
                    if (Int32.Parse(_mapInfoApp.Eval(String.Format("WindowInfo({0}, {1})", i,
                        (int)MapInfoConstants.WindowInfo.WIN_INFO_TYPE))) ==
                        (int)MapInfoConstants.WindowInfoWindowTypes.WIN_MAPPER)
                    {
                        // Increment the window number counter
                        windowNum += 1;

                        // Store the number of layers in the window
                        int numLayers = Int32.Parse(_mapInfoApp.Eval(String.Format("MapperInfo({0}, {1})",
                            windowID, (int)MapInfoConstants.MapperInfo.MAPPER_INFO_LAYERS)));

                        // Loop through all the layers in the current window
                        for (int j = 1; j <= numLayers; j++)
                        {
                            // If this is not a cosmetic layer
                            if (_mapInfoApp.Eval(String.Format("LayerInfo({0}, {1}, {2})", windowID, 
                                j, (int)MapInfoConstants.LayerInfo.LAYER_INFO_COSMETIC)) == "F")
                            {
                                // Store the name of the layer
                                layer = _mapInfoApp.Eval(String.Format("LayerInfo({0}, {1}, {2})", 
                                    windowID, j, (int)MapInfoConstants.LayerInfo.LAYER_INFO_NAME));

                                // Check to see if this layer is a HLU layer
                                if (IsHluLayer(layer))
                                {
                                    _hluMapWindowID = windowID;
                                    _hluCurrentLayer = new GISLayer(windowNum, 0, layer);

                                    // Disable the Close command in the window's system menu.
                                    _mapInfoApp.Do(String.Format("Set Window {0} SysMenuClose Off", _hluMapWindowID));
                                    // Note: Before version 10.5, you could enable or disable the Close
                                    // button regardless of the toolbar’s floating or docking state.
                                    // As of version 10.5, you cannot enable or disable the Close
                                    // button when the toolbar is docked. You can only change the
                                    // state when it is floating or floating and hidden.
                                    return true;
                                }
                            }
                        }
                    }
                }

                // no mapper window open, but there may still be an HLU table
                int numTables = Int32.Parse(_mapInfoApp.Eval("NumTables()"));

                string tabName;
                layer = null;
                for (int i = 1; i <= numTables; i++)
                {
                    tabName = _mapInfoApp.Eval(String.Format("TableInfo({0}, {1})", i, 
                        (int)MapInfoConstants.TableInfo.TAB_INFO_NAME));
                    if (IsHluLayer(tabName))
                    {
                        layer = tabName;
                        break;
                    }
                }

                // make a map from the open HLU table or browse for a table
                return AddLayer(layer);
            }
            catch { }

            return false;
        }

        //---------------------------------------------------------------------
        // CHANGED: CR31 (Switching between GIS layers)
        // Enable the user to switch between different HLU layers, where
        // there is more than one valid layer in the current document.
        //
        /// <summary>
        /// Determines which of the layers in all the windows are valid HLU layers
        /// and stores these in a list so the user can switch between them.
        /// Called before displaying the list of layers for the user to switch
        /// between.
        /// </summary>
        /// <returns>The number of valid HLU layers in the list</returns>
        public override int ListHluLayers()
        {
            if (_mapInfoApp == null) return 0;

            try
            {
                // check if the workspace contains a mapper window with a valid HLU layer
                int numWindows = Int32.Parse(_mapInfoApp.Eval("NumWindows()"));

                int windowID;
                int windowNum = 0;

                // Initialise the list of valid layers
                if (_hluLayerList == null) _hluLayerList = new List<GISLayer>();

                // Clear the valid HLU layer list
                _hluLayerList.Clear();

                // Loop through all windows
                for (int i = 1; i <= numWindows; i++)
                {
                    windowID = Int32.Parse(_mapInfoApp.Eval(String.Format("WindowID({0})", i)));

                    // If this is a mapper window
                    if (Int32.Parse(_mapInfoApp.Eval(String.Format("WindowInfo({0}, {1})", i,
                        (int)MapInfoConstants.WindowInfo.WIN_INFO_TYPE))) ==
                        (int)MapInfoConstants.WindowInfoWindowTypes.WIN_MAPPER)
                    {
                        // Increment the window number counter
                        windowNum += 1;

                        // Store the number of layers in the window
                        int numLayers = Int32.Parse(_mapInfoApp.Eval(String.Format("MapperInfo({0}, {1})",
                            windowID, (int)MapInfoConstants.MapperInfo.MAPPER_INFO_LAYERS)));

                        // Loop through all the layers in the current window
                        for (int j = 1; j <= numLayers; j++)
                        {
                            // If this is not a cosmetic layer
                            if (_mapInfoApp.Eval(String.Format("LayerInfo({0}, {1}, {2})", windowID,
                                j, (int)MapInfoConstants.LayerInfo.LAYER_INFO_COSMETIC)) == "F")
                            {
                                // Store the name of the layer
                                string layer = _mapInfoApp.Eval(String.Format("LayerInfo({0}, {1}, {2})",
                                    windowID, j, (int)MapInfoConstants.LayerInfo.LAYER_INFO_NAME));

                                // Check to see if this layer is a HLU layer
                                if (IsHluLayer(layer))
                                {
                                    // Add the name of layer to the valid list
                                    _hluLayerList.Add(new GISLayer(windowNum, 0, layer));
                                }
                            }
                        }
                    }
                }

                //// no mapper window open, but there may still be an HLU table
                //int numTables = Int32.Parse(_mapInfoApp.Eval("NumTables()"));

                //for (int i = 1; i <= numTables; i++)
                //{
                //    _hluLayer = _mapInfoApp.Eval(String.Format("TableInfo({0}, {1})", i,
                //        (int)MapInfoConstants.TableInfo.TAB_INFO_NAME));
                //    if (IsHluLayer()) break;
                //}

                //// make a map from the open HLU table or browse for a table
                //return AddLayer();
            }
            catch { }

            if (_hluCurrentLayer == null)
                _hluCurrentLayer = _hluLayerList[0];
            return _hluLayerList.Count();
        }
        //---------------------------------------------------------------------

        //---------------------------------------------------------------------
        // CHANGED: CR31 (Switching between GIS layers)
        // Enable the user to switch between different HLU layers, where
        // there is more than one valid layer in the current document.
        //
        /// <summary>
        /// Checks whether the layer/table that is passed to the method is an HLU layer, 
        /// as determined by its data structure. The data structure must follow the template of the 
        /// HLU.Data.Model.HluGISLayer.incid_mm_polygonsDataTable (same column names and data types, 
        /// as per type maps _typeMapSystemToSQL and _typeMapSQLToSystem).
        /// If _hluLayer points to a proper HLU layer _hluLayer, _hluFieldNames and _hluFieldMap 
        /// are initialized to their correct values; otherwise they are left as their old values.
        /// </summary>
        /// <returns>True if the new GIS is a valid HLU layer, otherwise False</returns>
        public override bool IsHluLayer(GISLayer newGISLayer)
        {
            int hluMapWindowIDBak = _hluMapWindowID;
            int[] hluFieldMapBak = _hluFieldMap;
            string[] hluFieldNamesBak = _hluFieldNames;
            string hluLayerBak = _hluLayer;
            string hluColumnListBak = _hluColumnList;

            try
            {
                // Store the previous HLU layer so that the user can remove it
                // if they wish.
                _hluLayerOld = _hluLayer;

                // Initialise the current HLU layer so that it will be set when
                // if the passed layer is a valid HLU layer.
                _hluLayer = null;

                string hluLayer = newGISLayer.LayerName;
                int windowID = Int32.Parse(_mapInfoApp.Eval(String.Format("WindowID({0})", newGISLayer.MapNum)));
                
                if (IsHluLayer(hluLayer))
                {
                    // If this layer is in a different window to the current HLU layer
                    if (_hluMapWindowID != windowID)
                    {
                        // Enable the Close command in the previous window's system menu.
                        _mapInfoApp.Do(String.Format("Set Window {0} SysMenuClose On", _hluMapWindowID));
                    }

                    _hluMapWindowID = windowID;
                    _hluCurrentLayer = newGISLayer;

                    // Disable the Close command in the new window's system menu.
                    _mapInfoApp.Do(String.Format("Set Window {0} SysMenuClose Off", _hluMapWindowID));
                }
            }
            catch { }

            // Reset the class properties if the layer is not valid.
            if (_hluLayer == null)
            {
                _hluMapWindowID = hluMapWindowIDBak;
                _hluFieldNames = hluFieldNamesBak;
                _hluFieldMap = hluFieldMapBak;
                _hluLayer = hluLayerBak;
                _hluColumnList = hluColumnListBak;
                return false;
            }
            else
                return true;
        }
        //---------------------------------------------------------------------

        /// <summary>
        /// Checks whether the layer/table that field hluLayer points to is an HLU layer,
        /// as determined by its data structure. The data structure must follow the template of the
        /// HLU.Data.Model.HluGISLayer.incid_mm_polygonsDataTable (same column names and data types,
        /// as per type maps _typeMapSystemToSQL and _typeMapSQLToSystem).
        /// If hluLayer points to a proper HLU layer _hluLayer, _hluFieldNames and _hluFieldMap
        /// are initialized to their correct values; otherwise they are set to null.
        /// </summary>
        /// <param name="hluLayer">The HLU layer that will be switched to.</param>
        /// <returns></returns>
        protected bool IsHluLayer(string hluLayer)
        {
            try
            {
                if (_mapInfoApp == null)
                    throw new Exception("No MapInfo application.");

                if (String.IsNullOrEmpty(hluLayer))
                    throw new Exception("No HLU layer.");

                // Check various characteristics of the layer and throw exceptions if not valid
                if (Int32.Parse(_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})", hluLayer,
                    (int)MapInfoConstants.TableInfo.TAB_INFO_TYPE))) != (int)MapInfoConstants.TableInfoType.TAB_TYPE_BASE)
                    throw new Exception(String.Format("Table {0} is not a base table.", hluLayer));

                if (_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})", hluLayer,
                    (int)MapInfoConstants.TableInfo.TAB_INFO_READONLY)) == "T")
                    throw new Exception(String.Format("Table {0} is read only.", hluLayer));

                if (_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})", hluLayer,
                    (int)MapInfoConstants.TableInfo.TAB_INFO_MAPPABLE)) == "F")
                    throw new Exception(String.Format("Table {0} is not mappable.", hluLayer));

                // Store the number of columns in the layer
                int numColumns = Int32.Parse(_mapInfoApp.Eval(String.Format("TableInfo({0}, {1})",
                    hluLayer, (int)MapInfoConstants.TableInfo.TAB_INFO_NCOLS)));

                if (_hluLayerStructure == null)
                    _hluLayerStructure = new HluGISLayer.incid_mm_polygonsDataTable();

                int[] hluFieldMap = _hluLayerStructure.Columns.Cast<DataColumn>().Select(c => -1).ToArray();
                string[] hluFieldNames = new string[numColumns];

                // Loop through all the columns in the layer
                for (int i = 1; i <= numColumns; i++)
                {
                    // Store the column field name
                    hluFieldNames[i - 1] = _mapInfoApp.Eval(String.Format("ColumnInfo({0}, {1}, {2})",
                        hluLayer, QuoteValue(String.Format("Col{0}", i)), (int)MapInfoConstants.ColumnInfo.COL_INFO_NAME));

                    DataColumn hluColumn = _hluLayerStructure.Columns.Contains(hluFieldNames[i - 1]) ?
                        _hluLayerStructure.Columns[hluFieldNames[i - 1]] : null;

                    if (hluColumn != null)
                    {
                        hluFieldMap[hluColumn.Ordinal] = i;

                        // Check the field type and length
                        Type colSysType;
                        if (!_typeMapSQLToSystem.TryGetValue(Int32.Parse(_mapInfoApp.Eval(String.Format(
                            "ColumnInfo({0}, {1}, {2})", hluLayer, QuoteValue(String.Format("Col{0}", i)),
                            (int)MapInfoConstants.ColumnInfo.COL_INFO_TYPE))), out colSysType) ||
                            (hluColumn.DataType != colSysType))
                            throw new Exception("Field type does not match the HLU GIS layer structure.");

                        if ((colSysType == typeof(string)) && (Int32.Parse(_mapInfoApp.Eval(
                            String.Format("ColumnInfo({0}, {1}, {2})", hluLayer, QuoteValue(String.Format("Col{0}", i)),
                            (int)MapInfoConstants.ColumnInfo.COL_INFO_WIDTH))) > hluColumn.MaxLength))
                            throw new Exception("Field length does not match the HLU GIS layer structure.");
                    }
                }

                if (!hluFieldMap.All(o => o != -1))
                    throw new Exception("Layer is missing some fields of the HLU GIS layer structure.");

                //---------------------------------------------------------------------
                // CHANGED: CR31 (Switching between GIS layers)
                // Enable the user to switch between different HLU layers, where
                // there is more than one valid layer in the current document.
                //
                // Only set the current HLU layer properties (i.e. _hluLayer, etc)
                // if this valid HLU layer is to replace an existing layer.
                if (_hluLayer == null)
                {
                    // The layer is a valid HLU layer so store the layer properties.
                    _hluLayer = hluLayer;
                    _hluFieldMap = hluFieldMap;
                    _hluFieldNames = hluFieldNames;

                    if (_hluLayerOld != null)
                    {
                        // Turn off speed edits ('Undo Off' and 'FastEdit Off') on the old HLU layer
                        // and allow the user to remove if from the map or close it.
                        _mapInfoApp.Do(String.Format("Set Table {0} Undo On", _hluLayerOld));
                        _mapInfoApp.Do(String.Format("Set Table {0} FastEdit Off", _hluLayerOld));
                        _mapInfoApp.Do(String.Format("Set Table {0} UserRemoveMap On", _hluLayerOld));
                        _mapInfoApp.Do(String.Format("Set Table {0} UserClose On", _hluLayerOld));
                        _mapInfoApp.Do(String.Format("Set Table {0} UserMap On", _hluLayerOld));
                    }

                    // The layer is a valid HLU layer so speed up edits ('Undo Off' and 'FastEdit Off')
                    // and stop the user from being able to remove the layer from the map or close it.
                    _mapInfoApp.Do(String.Format("Set Table {0} Undo Off", _hluLayer));
                    _mapInfoApp.Do(String.Format("Set Table {0} FastEdit On", _hluLayer));
                    _mapInfoApp.Do(String.Format("Set Table {0} UserRemoveMap Off", _hluLayer));
                    _mapInfoApp.Do(String.Format("Set Table {0} UserClose Off", _hluLayer));
                    _mapInfoApp.Do(String.Format("Set Table {0} UserMap Off", _hluLayer));

                    _hluColumnList = ColumnList(_hluLayer, null, true);
                }
                //---------------------------------------------------------------------

                return true;
            }
            catch
            {
                //_hluMapWindowID = -1;
                //_hluLayerStructure = null;
                //_hluFieldNames = null;
                //_hluFieldMap = null;
                //_hluLayer = null;
                //_hluColumnList = null;
                return false;
            }
        }

        #region Callback

        void miCallback_OnWindowChanged(int obj)
        {
            //MessageBox.Show(String.Format("obj = {0}", obj), "OnWindowChanged", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        void miCallback_OnMenuItemClick(string obj)
        {
            //MessageBox.Show(obj, "OnMenuItemClick", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        void miCallback_OnStatusChanged(string obj)
        {
            //MessageBox.Show(obj, "OnStatusChanged", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        #endregion
    }
}
