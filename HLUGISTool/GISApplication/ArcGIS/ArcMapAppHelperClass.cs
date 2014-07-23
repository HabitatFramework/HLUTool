// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// Copyright © 2013 Thames Valley Environmental Records Centre
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
using System.Text;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using HLU.Data.Model;
using HLU.Properties;
using ESRI.ArcGIS.Carto;

namespace HLU.GISApplication.ArcGIS
{
    /// <summary>
    /// Shared by ArcMapApp and HluArcMapExtension
    /// </summary>
    public class ArcMapAppHelperClass
    {
        /// <summary>
        /// Retrieves the field of _hluFeatureClass that corresponds to the column of _hluLayerStructure whose name is passed in.
        /// </summary>
        /// <param name="columnName">Name of the column of _hluLayerStructure.</param>
        /// <returns>The field of _hluFeatureClass corresponding to column _hluLayerStructure[columnName].</returns>
        public static IField GetField(string columnName, IFeatureClass _hluFeatureClass,
            HluGISLayer.incid_mm_polygonsDataTable _hluLayerStructure, int[] _hluFieldMap)
        {
            if ((_hluLayerStructure == null) || (_hluFieldMap == null) ||
                (_hluFeatureClass == null) || String.IsNullOrEmpty(columnName)) return null;
            DataColumn c = _hluLayerStructure.Columns[columnName.Trim()];
            if ((c == null) || (c.Ordinal >= _hluFieldMap.Length)) return null;
            int fieldOrdinal = _hluFieldMap[c.Ordinal];
            if ((fieldOrdinal >= 0) && (fieldOrdinal <= _hluFieldMap.Length))
                return _hluFeatureClass.Fields.get_Field(fieldOrdinal);
            else
                return null;
        }

        public static bool IsHluLayer(IFeatureLayer layer, IFields origFields,
            IFieldChecker fieldChecker, object[] validWorkspaces, Dictionary<Type, int> typeMapSystemToSQL,
            ref HluGISLayer.incid_mm_polygonsDataTable hluLayerStructure,
            out int[] hluFieldMap, out string[] hluFieldNames)
        {
            hluFieldMap = null;
            hluFieldNames = null;

            if ((layer == null) || !layer.Valid || (((IFeatureLayerDefinition)layer).DefinitionSelectionSet != null))
                return false;

            bool isHlu = true;

            try
            {
                IFeatureClass testFeatureClass = layer.FeatureClass;

                if (hluLayerStructure == null) hluLayerStructure = new HluGISLayer.incid_mm_polygonsDataTable();
                IFeatureWorkspace testWorkspace = ((IDataset)testFeatureClass).Workspace as IFeatureWorkspace;
                if (testWorkspace == null)
                    throw (new Exception("Invalid feature workspace."));
                if (System.Array.IndexOf(validWorkspaces, ((IWorkspace)testWorkspace).WorkspaceFactory.GetClassID().Value) == -1)
                    throw (new Exception("Invalid workspace type."));

                if (testFeatureClass.ShapeType != esriGeometryType.esriGeometryPolygon)
                    throw (new Exception("Invalid geometry type."));

                IFieldsEdit fieldsEdit = (IFieldsEdit)origFields;
                foreach (DataColumn c in hluLayerStructure.Columns)
                {
                    IField newField = new FieldClass();
                    IFieldEdit2 newFieldEdit = (IFieldEdit2)newField;
                    newFieldEdit.Name_2 = c.ColumnName;
                    int fieldType;
                    if (!typeMapSystemToSQL.TryGetValue(c.DataType, out fieldType))
                        throw (new Exception("Invalid field type."));
                    newFieldEdit.Type_2 = (esriFieldType)fieldType;
                    if ((c.MaxLength != -1) && (newField.Type == esriFieldType.esriFieldTypeString))
                        newFieldEdit.Length_2 = c.MaxLength;
                    fieldsEdit.AddField(newField);
                }

                fieldChecker.ValidateWorkspace = (IWorkspace)testWorkspace;

                IEnumFieldError error;
                IFields fixedFields;
                fieldChecker.Validate(origFields, out error, out fixedFields);

                hluFieldMap = new int[origFields.FieldCount];
                hluFieldNames = new string[hluFieldMap.Length];

                for (int i = 0; i < fixedFields.FieldCount; i++)
                {
                    IField fixedField = fixedFields.get_Field(i);

                    int ordinal = testFeatureClass.Fields.FindField(fixedField.Name);

                    if (ordinal == -1)
                        throw (new Exception("Field name does not match the HLU GIS layer structure."));
                    IField fcField = testFeatureClass.Fields.get_Field(ordinal);
                    if (fcField.Type != fixedField.Type)
                        throw (new Exception("Field type does not match the HLU GIS layer structure."));
                    if ((fcField.Type == esriFieldType.esriFieldTypeString) && (fcField.Length > fixedField.Length))
                        throw (new Exception("Field length does not match the HLU GIS layer structure."));

                    hluFieldMap[i] = ordinal;
                    //---------------------------------------------------------------------
                    // FIXED: KI107 (GIS layer column names)
                    // Use the field names from the GIS layer so that they can be found
                    // when performing any SELECT statements.
                    hluFieldNames[i] = fcField.Name;
                    //hluFieldNames[i] = fixedField.Name;
                    //---------------------------------------------------------------------
                }
            }
            catch { return false; }

            return isHlu;
        }

        public static string WhereClauseFromCursor(int oidOrdinalCursor, string oidColumnAlias, ICursor cursor)
        {
            StringBuilder sbIDs = new StringBuilder();
            StringBuilder sbBetween = new StringBuilder();
            string betweenTemplate = " OR (" + oidColumnAlias + " BETWEEN {0} AND {1})";
            int currOid = -1;
            int nextOid = -1;
            int countContinuous = 0;
            IRow row = cursor.NextRow();

            while (row != null)
            {
                currOid = (int)row.get_Value(oidOrdinalCursor);
                nextOid = currOid;
                countContinuous = 1;
                do
                {
                    row = cursor.NextRow();
                    if (row != null)
                    {
                        nextOid = (int)row.get_Value(oidOrdinalCursor);
                        if (nextOid != currOid + countContinuous)
                            break;
                        else
                            countContinuous++;
                    }
                    else
                    {
                        break;
                    }
                }
                while (true);
                switch (countContinuous)
                {
                    case 1:
                        sbIDs.Append(",").Append(currOid);
                        break;
                    case 2:
                        sbIDs.Append(",").Append(currOid);
                        if (nextOid != currOid) sbIDs.Append(",").Append(nextOid);
                        break;
                    default:
                        sbBetween.Append(String.Format(betweenTemplate, currOid, currOid + countContinuous - 1));
                        break;
                }
            }

            if (sbIDs.Length > 1) sbIDs.Remove(0, 1).Insert(0, oidColumnAlias + " IN (").Append(")");
            return sbIDs.Append(sbBetween).ToString();
        }

        public static IQueryDef CreateQueryDef(IFeatureWorkspace featureWorkspace,
            String tables, String subFields, String whereClause)
        {
            // Create the query definition.
            IQueryDef queryDef = featureWorkspace.CreateQueryDef();

            // Provide a list of table(s) to join.
            queryDef.Tables = tables;

            // Declare the subfields to retrieve.
            queryDef.SubFields = subFields; // must be qualified if multiple tables !!

            // Assign a where clause to filter the results.
            queryDef.WhereClause = whereClause;

            return queryDef;
        }

        public static void GetTypeMaps(out Dictionary<string, esriSQLPredicates> _sqlPredicates,
            out Dictionary<Type, int> _typeMapSystemToSQL, out Dictionary<int, Type> _typeMapSQLToSystem)
        {
            _sqlPredicates = new Dictionary<string, esriSQLPredicates>();
            _sqlPredicates.Add("=", esriSQLPredicates.esriSQL_EQ);
            _sqlPredicates.Add("<>", esriSQLPredicates.esriSQL_NOT_EQ);
            _sqlPredicates.Add("<", esriSQLPredicates.esriSQL_LT);
            _sqlPredicates.Add(">", esriSQLPredicates.esriSQL_GT);
            _sqlPredicates.Add("<=", esriSQLPredicates.esriSQL_LE);
            _sqlPredicates.Add(">=", esriSQLPredicates.esriSQL_GE);
            _sqlPredicates.Add("BETWEEN", esriSQLPredicates.esriSQL_BETWEEN);
            _sqlPredicates.Add("NOT BETWEEN", esriSQLPredicates.esriSQL_NOT_BETWEEN);
            _sqlPredicates.Add("IN", esriSQLPredicates.esriSQL_IN);
            _sqlPredicates.Add("NOT IN", esriSQLPredicates.esriSQL_NOT_IN);
            _sqlPredicates.Add("LIKE", esriSQLPredicates.esriSQL_LIKE);
            _sqlPredicates.Add("NOT LIKE", esriSQLPredicates.esriSQL_NOT_LIKE);
            _sqlPredicates.Add("IS NULL", esriSQLPredicates.esriSQL_IS_NULL);
            _sqlPredicates.Add("IS NOT NULL", esriSQLPredicates.esriSQL_IS_NOT_NULL);
            _sqlPredicates.Add("EXISTS", esriSQLPredicates.esriSQL_EXISTS);
            _sqlPredicates.Add("NOT EXISTS", esriSQLPredicates.esriSQL_NOT_EXISTS);

            _typeMapSystemToSQL = new Dictionary<Type, int>();
            _typeMapSystemToSQL.Add(typeof(System.String), (int)esriFieldType.esriFieldTypeString);
            _typeMapSystemToSQL.Add(typeof(System.Decimal), (int)esriFieldType.esriFieldTypeSingle);
            _typeMapSystemToSQL.Add(typeof(System.Int64), (int)esriFieldType.esriFieldTypeInteger);
            _typeMapSystemToSQL.Add(typeof(System.Int32), (int)esriFieldType.esriFieldTypeSmallInteger);
            _typeMapSystemToSQL.Add(typeof(System.Int16), (int)esriFieldType.esriFieldTypeSmallInteger);
            _typeMapSystemToSQL.Add(typeof(System.Boolean), (int)esriFieldType.esriFieldTypeSmallInteger);
            _typeMapSystemToSQL.Add(typeof(System.Single), (int)esriFieldType.esriFieldTypeSingle);
            _typeMapSystemToSQL.Add(typeof(System.Double), (int)esriFieldType.esriFieldTypeDouble);
            _typeMapSystemToSQL.Add(typeof(System.DateTime), (int)esriFieldType.esriFieldTypeDate);

            _typeMapSQLToSystem = new Dictionary<int, Type>();
            _typeMapSQLToSystem.Add((int)esriFieldType.esriFieldTypeBlob, typeof(System.Byte[]));
            _typeMapSQLToSystem.Add((int)esriFieldType.esriFieldTypeDate, typeof(System.DateTime));
            _typeMapSQLToSystem.Add((int)esriFieldType.esriFieldTypeDouble, typeof(System.Double));
            _typeMapSQLToSystem.Add((int)esriFieldType.esriFieldTypeGeometry, typeof(System.Byte[]));
            _typeMapSQLToSystem.Add((int)esriFieldType.esriFieldTypeGlobalID, typeof(System.Guid));
            _typeMapSQLToSystem.Add((int)esriFieldType.esriFieldTypeGUID, typeof(System.Guid));
            _typeMapSQLToSystem.Add((int)esriFieldType.esriFieldTypeInteger, typeof(System.Int64));
            _typeMapSQLToSystem.Add((int)esriFieldType.esriFieldTypeOID, typeof(System.Int64));
            _typeMapSQLToSystem.Add((int)esriFieldType.esriFieldTypeRaster, typeof(System.Byte[]));
            _typeMapSQLToSystem.Add((int)esriFieldType.esriFieldTypeSingle, typeof(System.Single));
            _typeMapSQLToSystem.Add((int)esriFieldType.esriFieldTypeSmallInteger, typeof(System.Int32));
            _typeMapSQLToSystem.Add((int)esriFieldType.esriFieldTypeString, typeof(System.String));
            _typeMapSQLToSystem.Add((int)esriFieldType.esriFieldTypeXML, typeof(System.String));

            /*
            Shapefile -> Geodatabase
            Date NA -> Date 
            String 1–255 -> Text 
            Boolean NA -> Short integer 
            Number 1–4 (decimals=0) -> Short integer 
            Number 5–9 (decimals=0) -> Long integer 
            Number 10–19 (decimals=0) -> Double 
            Float 1–13 -> Float 
            Float 14–19 -> Double 
            Number 1–8 (decimals>0) -> Float 
            Number 9–19 (decimals>0) -> Double 
            
            Access Workspace:
            esriFieldType.esriFieldTypeOID, "Long Integer" (AutoNumber) 
            esriFieldType.esriFieldTypeSmallInteger, "Integer" 
            esriFieldType.esriFieldTypeInteger, "Long Integer" 
            esriFieldType.esriFieldTypeSingle, "Single" 
            esriFieldType.esriFieldTypeDouble, "Double" 
            esriFieldType.esriFieldTypeString, "Text" 
            esriFieldType.esriFieldTypeDate, "Date/Time" 
            esriFieldType.esriFieldTypeBlob, "OLE Object" * 
            esriFieldType.esriFieldTypeGUID, "Number" 
            esriFieldType.esriFieldTypeGeometry, "OLE Object" * 
            esriFieldType.esriFieldTypeRaster, "Memo"

            DB2 Workspace
            esriFieldType.esriFieldTypeOID , "INTEGER(4) NOT NULL" 
            esriFieldType.esriFieldTypeSmallInteger, "SMALLINT(2)" 
            esriFieldType.esriFieldTypeInteger, "INTEGER(4)" 
            esriFieldType.esriFieldTypeSingle, "DECIMAL(31,8)" 
            esriFieldType.esriFieldTypeDouble, "DECIMAL(31,8)" 
            esriFieldType.esriFieldTypeString, "VARCHAR(n)" 
            esriFieldType.esriFieldTypeDate, "TIMESTAMP" 
            esriFieldType.esriFieldTypeBlob, "BLOB" 
            esriFieldType.esriFieldTypeGUID, "CHARACTER(38) 
            esriFieldType.esriFieldTypeGeometry, "ST_GEOMETRY" // ST_Geometry is a superclass. The actual data subtype created (such as ST_Multilinestring or ST_Point) depends on what type of feature class you create such as polygon, line, point, multipoint, and multipatch.)
            esriFieldType.esriFieldTypeRaster, "INTEGER(4)" 

            Informix Workspace
            esriFieldType.esriFieldTypeOID, "INT(4) NOT NULL" 
            esriFieldType.esriFieldTypeSmallInteger, "INTEGER SMALLINT(2)" 
            esriFieldType.esriFieldTypeInteger, "INTEGER INT(4)" 
            esriFieldType.esriFieldTypeSingle, "DECIMAL(32)" 
            esriFieldType.esriFieldTypeDouble, "DECIMAL(32)" 
            esriFieldType.esriFieldTypeString, "VARCHAR(n)" 
            esriFieldType.esriFieldTypeDate, "DATETIME" 
            esriFieldType.esriFieldTypeBlob, "BLOB" 
            esriFieldType.esriFieldTypeGUID, "CHAR(8)" 
            esriFieldType.esriFieldTypeGeometry, "ST_GEOMETRY" // ST_Geometry is a superclass. The actual data subtype created (such as ST_Multilinestring or ST_Point) depends on what type of feature class you create such as polygon, line, point, multipoint, and multipatch. 
            esriFieldType.esriFieldTypeRaster, "INT(4)" 
             
            Oracle Workspace
            esriFieldType.esriFieldTypeOID, "NUMBER(38) NOT NULL" 
            esriFieldType.esriFieldTypeSmallInteger, "INTEGER NUMBER(4)" 
            esriFieldType.esriFieldTypeInteger, "INTEGER NUMBER(38)" 
            esriFieldType.esriFieldTypeSingle, "NUMBER(38,8)" 
            esriFieldType.esriFieldTypeDouble, "NUMBER(38,8)" 
            esriFieldType.esriFieldTypeString, "VARCHAR2(50)" 
            esriFieldType.esriFieldTypeDate, "DATE" 
            esriFieldType.esriFieldTypeBlob, "BLOB" 
            esriFieldType.esriFieldTypeGUID, "CHAR(38)" 
            esriFieldType.esriFieldTypeGeometry, "ST_GEOMETRY" or "NUMBER(38)" or "SDO_GEOMETRY" // Oracle data type depends on the geometry storage specified for the layer.  
            esriFieldType.esriFieldTypeRaster, "BLOB" or "LONG_RAW" 

            PostgreSQL Workspace
            esriFieldType.esriFieldTypeOID, "INTEGER" 
            esriFieldType.esriFieldTypeSmallInteger, "SMALLINT" 
            esriFieldType.esriFieldTypeInteger, "INTEGER" 
            esriFieldType.esriFieldTypeSingle, "NUMERIC(38,8)" 
            esriFieldType.esriFieldTypeDouble, "NUMERIC(38,8)" 
            esriFieldType.esriFieldTypeString, "CHARACTER VARYING" 
            esriFieldType.esriFieldTypeDate, "TIME WITHOUT TIME ZONE" 
            esriFieldType.esriFieldTypeBlob BYTEA" 
            esriFieldType.esriFieldTypeGUID, "CHARACTER VARYING(38)" 
            esriFieldType.esriFieldTypeGeometry, "ST_GEOMETRY" or "geometry" // Data type depends on the geometry storage specified for the layer.  
            esriFieldType.esriFieldTypeRaster, "INTEGER" 

            SQL Server Workspace
            esriFieldType.esriFieldTypeOID, "INT(4) NOT NULL"
            esriFieldType.esriFieldTypeSmallInteger, "SMALLINT(2)"
            esriFieldType.esriFieldTypeInteger, "INT(4)"
            esriFieldType.esriFieldTypeSingle, "REAL"
            esriFieldType.esriFieldTypeDouble, "DOUBLE" or "NUMERIC"
            esriFieldType.esriFieldTypeString, "VARCHAR(n)"
            esriFieldType.esriFieldTypeDate, "DATETIME"
            esriFieldType.esriFieldTypeBlob, "IMAGE"
            esriFieldType.esriFieldTypeGUID, "UNIQUEIDENTIFIER(16)"
            esriFieldType.esriFieldTypeGeometry, "IMAGE"
            esriFieldType.esriFieldTypeRaster, "IMAGE"
             */
        }

        public static void GetValidWorkspaces(out object[] _validWorkspaces)
        {
            _validWorkspaces = new object[Settings.Default.ArcGISValidWorkspaces.Count];
            for (int i = 0; i < _validWorkspaces.Length; i++)
            {
                UID uid = new UIDClass();
                uid.Value = Settings.Default.ArcGISValidWorkspaces[i];
                _validWorkspaces[i] = uid.Value;
            }
        }
    }
}
