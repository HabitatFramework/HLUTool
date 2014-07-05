// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2013 Andy Foy
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
using HLU.Data;
using HLU.Data.Model;
using HLU.Properties;

namespace HLU.GISApplication
{
    public abstract class GISApp : SqlBuilder
    {
        public enum DistanceUnits
        {
            Chains, Centimeters, Feet, Inches, Kilometers, Links, Meters, 
            Miles, Millimeters, NauticalMiles, Rods, SurveyFeet, Yards
        }

        public enum AreaUnits
        {
            Acres, Hectares, Perches, Roods, SquareChains, SquareCentimeters,
            SquareFeet, SquareInches, SquareKilometers, SquareLinks, SquareMeters,
            SquareMiles, SquareMillimeters, SquareRods, SquareSurveyFeet, SquareYards
        }

        public abstract DistanceUnits DistanceUnit { set; }

        public abstract AreaUnits AreaUnit { set; }

        /// <summary>
        /// Type of GIS application.
        /// </summary>
        public abstract GISApplications ApplicationType { get; }

        /// <summary>
        /// Reference to the running GIS application object.
        /// </summary>
        public abstract object ApplicationObject { get; }

        /// <summary>
        /// True if the GIS application is running, otherwise false.
        /// </summary>
        public abstract bool IsRunning { get; }

        /// <summary>
        /// True if GIS application is in editing session.
        /// This is an ArcGIS concept that has no equivalent in the MapInfo world.
        /// </summary>
        public abstract bool IsEditing { get; }

        /// <summary>
        /// Launches an instance of the GIS application.
        /// </summary>
        /// <param name="waitSeconds">Number of seconds to wait for the GIS application process to load before an exception is thrown.</param>
        /// <returns>true if GIS application launched ok, otherwise false.</returns>
        public abstract bool Start(ProcessWindowStyle windowStyle);

        public abstract void Window(ProcessWindowStyle windowStyle, IntPtr sideBySideWith);

        public abstract IntPtr hWnd { get; }

        public abstract void Activate();

        public abstract bool Close();

        public abstract bool OpenWorkspace(string path);

        public abstract bool SaveWorkspace();

        public abstract string HluLayerName { get; }

        public abstract string IncidFieldName { get; }

        public abstract void ReadMapSelection(ref DataTable resultTable);

        public abstract void FlashSelectedFeature(List<SqlFilterCondition> whereClause);

        public abstract void FlashSelectedFeatures(List<List<SqlFilterCondition>> whereClauses);

        public abstract void ZoomSelected();

        public abstract int HluLayerCount { get; }

        public abstract List<GISLayer> ValidHluLayers { get; }

        public abstract GISLayer CurrentHluLayer { get; }

        public abstract DataTable UpdateFeatures(DataColumn[] updateColumns, 
            object[] updateValues, DataColumn[] historyColumns);

        public abstract DataTable UpdateFeatures(DataColumn[] updateColumns, object[] updateValues, 
            DataColumn[] historyColumns, List<SqlFilterCondition> selectionWhereClause);

        public abstract DataTable UpdateFeatures(DataColumn[] updateColumns, object[] updateValues, 
            DataColumn[] historyColumns, string tempMdbPathName, string selectionTableName);

        /// <summary>
        /// Following a feature split performed by the user with regular GIS tools, updates the toid_fragment_id of 
        /// the newly created feature(s), incrementing from the highest toid_fragment_id for that toid passed as 
        /// parameter lastToidFragmentID.
        /// </summary>
        /// <param name="currentToidFragmentID">Current toid_fragment_id for the selected toid.</param>
        /// <param name="lastToidFragmentID">Highest toid_fragment_id for the selected toid.</param>
        /// <param name="selectionWhereClause">Where clause underlying the current selection set.</param>
        /// <param name="historyColumns">All columns of the GIS layer structure except the geomtry property columns.</param>
        /// <returns>DataTable of all the features resulting from the split, with all the columns in historyColumns parameter
        /// and the original feature in the first row.</returns>
        //---------------------------------------------------------------------
        // FIXED: KI110 (Physical split)
        // Pass the currentToidFragmentID (pre-split) to the function so that
        // it can re-select it, and the new ToidFragmentIDs, after the split.
        //---------------------------------------------------------------------
        public abstract DataTable SplitFeature(string currentToidFragmentID, string lastToidFragmentID,
            List<SqlFilterCondition> selectionWhereClause, DataColumn[] historyColumns);

        public abstract DataTable SplitFeaturesLogically(string oldIncid, string newIncid, DataColumn[] historyColumns);

        public abstract DataTable MergeFeatures(string newToidFragmentID, 
            List<SqlFilterCondition> resultWhereClause, DataColumn[] historyColumns);

        public abstract DataTable MergeFeaturesLogically(string keepIncid, DataColumn[] historyColumns);

        public abstract bool Export(string tempMdbPathName, string attributeDatasetName, int exportRowCount);

        public abstract DataTable SqlSelect(string scratchMdbPath, string selectionTableName, DataColumn[] targetColumns);

        public abstract DataTable SqlSelect(bool selectDistinct, bool addGeometryInfo, DataColumn[] targetColumns, List<SqlFilterCondition> whereConds);

        protected abstract bool IsHluWorkspace();

        public abstract int ListHluLayers();

        public abstract bool IsHluLayer(GISLayer newGISLayer);

        protected abstract string GetFieldName(int columnOrdinal);

        public static readonly string HistoryAdditionalFieldsDelimiter = Settings.Default.HistoryAdditionalFieldsDelimiter;

        protected List<SqlFilterCondition> MapWhereClauseFields(
            HluGISLayer.incid_mm_polygonsDataTable _hluLayerStructure, List<SqlFilterCondition> whereClause)
        {
            List<SqlFilterCondition> outWhereClause = new List<SqlFilterCondition>();
            for (int i = 0; i < whereClause.Count; i++)
            {
                SqlFilterCondition cond = whereClause[i];
                if (!_hluLayerStructure.Columns.Contains(cond.Column.ColumnName))
                {
                    if ((!String.IsNullOrEmpty(cond.CloseParentheses)) && (outWhereClause.Count > 0))
                    {
                        SqlFilterCondition condPrev = outWhereClause[outWhereClause.Count - 1];
                        condPrev.CloseParentheses += cond.CloseParentheses;
                        outWhereClause[outWhereClause.Count - 1] = condPrev;
                    }
                    if ((!String.IsNullOrEmpty(cond.OpenParentheses)) && (i < whereClause.Count - 1))
                    {
                        SqlFilterCondition condNext = whereClause[i + 1];
                        condNext.OpenParentheses += cond.OpenParentheses;
                        whereClause[i + 1] = condNext;
                    }
                    continue;
                }
                string columnName = GetFieldName(_hluLayerStructure.Columns[cond.Column.ColumnName].Ordinal);
                if (!String.IsNullOrEmpty(columnName))
                {
                    cond.Column = new DataColumn(columnName, cond.Column.DataType);
                    outWhereClause.Add(cond);
                }
            }
            return outWhereClause;
        }
    }
}
