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
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using System.Globalization;
using HLU.Data.Connection;
using HLU.Data.Model;
using HLU.GISApplication;
using HLU.Properties;
using HLU.UI.View;
using HLU.Data;
using HLU.Date;

namespace HLU.UI.ViewModel
{
    class ViewModelWindowMainExport
    {
        ViewModelWindowMain _viewModelMain;
        private WindowExport _windowExport;
        private ViewModelExport _viewModelExport;

        private string _lastTableName;
        private int _tableCount;
        private int _fieldCount;
        private int _sourceIdOrdinal;
        private List<int> _sourceDateStartOrdinals;
        private List<int> _sourceDateEndOrdinals;
        private List<int> _sourceDateTypeOrdinals;
        private int _attributesLength;

        public ViewModelWindowMainExport(ViewModelWindowMain viewModelMain)
        {
            _viewModelMain = viewModelMain;
        }

        public void InitiateExport()
        {
            _windowExport = new WindowExport();
            _windowExport.Owner = App.Current.MainWindow;
            _windowExport.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            // Fill all export formats if there are any export fields
            // defined for the export format.
            _viewModelMain.HluTableAdapterManager.exportsTableAdapter.ClearBeforeFill = true;
            _viewModelMain.HluTableAdapterManager.exportsTableAdapter.Fill(_viewModelMain.HluDataset.exports,
                String.Format("EXISTS (SELECT {0}.{1} FROM {0} WHERE {0}.{1} = {2}.{3})",
                _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.exports_fields.TableName),
                _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.exports_fields.export_idColumn.ColumnName),
                _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.exports.TableName),
                _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.exports.export_idColumn.ColumnName)));

            // If there are no exports formats defined that have any
            // export fields then exit.
            if (_viewModelMain.HluDataset.exports.Count == 0)
            {
                MessageBox.Show("Cannot export: there are no export formats defined.",
                    "HLU: Export", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // Display the export interface to prompt the user
            // to select which export format they want to use.
            _viewModelExport = new ViewModelExport(_viewModelMain.GisSelection == null ? 0 :
                _viewModelMain.GisSelection.Rows.Count, _viewModelMain.GISApplication.HluLayerName,
                _viewModelMain.GISApplication.ApplicationType, _viewModelMain.HluDataset.exports);
            _viewModelExport.DisplayName = "Export";
            _viewModelExport.RequestClose += new ViewModelExport.RequestCloseEventHandler(_viewModelExport_RequestClose);

            _windowExport.DataContext = _viewModelExport;

            _windowExport.ShowDialog();
        }

        private void _viewModelExport_RequestClose(int exportID, bool selectedOnly)
        {
            _viewModelExport.RequestClose -= _viewModelExport_RequestClose;
            _windowExport.Close();

            // If the user selected an export format then
            // perform the export using that format.
            if (exportID != -1)
            {
                DispatcherHelper.DoEvents();
                Export(exportID, selectedOnly);
            }
        }

        /// <summary>
        /// Exports the combined GIS and database data using
        /// the specified export format.
        /// </summary>
        /// <param name="userExportId">The export format selected by the user.</param>
        /// <param name="selectedOnly">If set to <c>true</c> export only selected incids/features.</param>
        /// <exception cref="System.Exception">
        /// Failed to find a table alias that does not match a table name in the HLU dataset
        /// or
        /// No export fields are defined for format 'x'
        /// or
        /// Export query did not retrieve any rows
        /// </exception>
        private void Export(int userExportId, bool selectedOnly)
        {
            string tempPath = null;

            try
            {
                _viewModelMain.ChangeCursor(Cursors.Wait, "Preparing ...");

                // Create a new unique table name to export to.
                string tableAlias = GetTableAlias();
                if (tableAlias == null)
                    throw new Exception("Failed to find a table alias that does not match a table name in the HLU dataset");

                // Retrieve the export fields for the export format
                // selected by the user from the database.
                _viewModelMain.HluTableAdapterManager.exportsFieldsTableAdapter.ClearBeforeFill = true;
                _viewModelMain.HluTableAdapterManager.exportsFieldsTableAdapter.Fill(
                    _viewModelMain.HluDataset.exports_fields, String.Format("{0} = {1} ORDER BY {2}, {3}",
                    _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.exports_fields.export_idColumn.ColumnName),
                    userExportId, 
                    _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.exports_fields.table_nameColumn.ColumnName),
                    _viewModelMain.DataBase.QuoteIdentifier(_viewModelMain.HluDataset.exports_fields.field_ordinalColumn.ColumnName)));

                // Exit if there are no export fields for this format
                // (this should not be possible).
                if (_viewModelMain.HluDataset.exports_fields.Count == 0)
                    throw new Exception(String.Format("No export fields are defined for format '{0}'",
                        _viewModelMain.HluDataset.exports.FindByexport_id(userExportId).export_name));

                // Build a new export data table and also determine
                // the number of output fields, a string of the output
                // field names, the from SQL clause, and the incid
                // field ordinal.
                DataTable exportTable;
                int[][] fieldMapTemplate;
                StringBuilder targetList;
                StringBuilder fromClause;
                int incidOrdinal;
                int[] dupsAllowed;
                List<ExportField> exportFields = new List<ExportField>();
                ExportJoins(tableAlias, ref exportFields, out exportTable,
                    out fieldMapTemplate, out targetList, out fromClause, out incidOrdinal, out dupsAllowed);

                // Add the length of the GIS layer fields to the length of
                // field attributes exported from the database (excluding
                // the incid field which is already included).
                _attributesLength += 93;

                // Check the total attributes length doesn't exceed the MapInfo
                // maximum record length.
                if ((_viewModelMain.GISApplication.ApplicationType == GISApplications.MapInfo) &&
                    (_attributesLength > Settings.Default.MapInfoMaxRecordLength))
                    throw new Exception(String.Format("The export format record length ({0} bytes) exceeds the maximum allowed for MapInfo (4000 bytes)", _attributesLength));

                // Set the export filter conditions, depending if all
                // the records are to be exported or only the selected
                // features.
                List<List<SqlFilterCondition>> exportFilter = null;
                if (!selectedOnly)
                {
                    SqlFilterCondition cond = new SqlFilterCondition("AND",
                        _viewModelMain.IncidTable, _viewModelMain.IncidTable.incidColumn, null);
                    cond.Operator = "IS NOT NULL";
                    exportFilter = new List<List<SqlFilterCondition>>(new List<SqlFilterCondition>[] { 
                        new List<SqlFilterCondition>(new SqlFilterCondition[] { cond }) });
                }
                else
                {
                    if ((_viewModelMain.IncidSelectionWhereClause == null) &&
                        (_viewModelMain.GisSelection != null) && (_viewModelMain.GisSelection.Rows.Count > 0))
                        _viewModelMain.IncidSelectionWhereClause = ViewModelWindowMainHelpers.GisSelectionToWhereClause(
                            _viewModelMain.GisSelection.Select(), _viewModelMain.GisIDColumnOrdinals,
                            ViewModelWindowMain.IncidPageSize, _viewModelMain.HluDataset.incid);
                    exportFilter = _viewModelMain.IncidSelectionWhereClause;
                }

                // Union the constituent parts of the export query
                // together into a single SQL string.
                string sql = ScratchDb.UnionQuery(targetList.ToString(), fromClause.ToString(),
                    incidOrdinal, exportFilter, _viewModelMain.DataBase);

                // Export the attribute data to a temporary database.
                int exportRowCount;
                tempPath = ExportMdb(sql, exportFields, exportTable, incidOrdinal, dupsAllowed, ref fieldMapTemplate, out exportRowCount);

                // Call the GIS application export method to join the
                // temporary attribute data to the GIS feature layer
                // and save them as a new GIS layer.
                if (!String.IsNullOrEmpty(tempPath) && (exportRowCount > 0))
                {
                    _viewModelMain.ChangeCursor(Cursors.Wait, "Exporting ...");

                    _viewModelMain.GISApplication.Export(tempPath, exportTable.TableName, exportRowCount);
                }
                else
                    throw new Exception("Export query did not retrieve any rows");

                // Remove the current record filter.
                _viewModelMain.IncidSelection = null;
                _viewModelMain.GisSelection = null;
                _viewModelMain.OnPropertyChanged("IsFiltered");
                _viewModelMain.OnPropertyChanged("StatusIncid");
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Export failed. The error message was:\n\n{0}.",
                    ex.Message), "HLU: Export", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (!String.IsNullOrEmpty(tempPath))
                {
                    string[] tempFiles = Directory.GetFiles(Path.GetDirectoryName(tempPath),
                        Path.GetFileNameWithoutExtension(tempPath) + ".*");
                    foreach (string fName in tempFiles)
                    {
                        try { File.Delete(fName); }
                        catch { _viewModelMain.ExportMdbs.Add(fName); }
                    }
                }
                _viewModelMain.ChangeCursor(Cursors.Arrow, null);
            }
        }

        private void ExportJoins(string tableAlias, ref List<ExportField> exportFields, out DataTable exportTable,
            out int[][] fieldMapTemplate, out StringBuilder targetList, out StringBuilder fromClause,
            out int incidOrdinal, out int[] dupsAllowed)
        {
            exportTable = new DataTable("HluExport");
            targetList = new StringBuilder();
            List<string> fromList = new List<string>();
            List<string> leftJoined = new List<string>();
            fromClause = new StringBuilder();
            dupsAllowed = null;

            int tableAliasNum = 1;
            bool firstJoin = true;
            incidOrdinal = -1;
            _lastTableName = null;
            _fieldCount = 0;
            _attributesLength = 0;

            //
            foreach (HluDataSet.exports_fieldsRow r in
                _viewModelMain.HluDataset.exports_fields.OrderBy(r => r.field_ordinal))
            {
                // If the field is not sourced from the database then
                // add it to the export table but don't add it to the
                // sql target list or from clause.
                if (r.table_name.ToLower() == "<none>")
                {
                    AddExportColumn(0, r.table_name, r.column_name, r.field_name,
                        r.field_type, r.field_length, !r.IsNull(_viewModelMain.HluDataset.exports_fields.field_formatColumn) ? r.field_format : null,
                        ref exportFields);
                    continue;
                }

                // Determine if this field is to be output multiple
                // times, once for each row in the relevant table
                // up to the maximum fields_count value.
                bool multipleFields = false;
                if (!r.IsNull(_viewModelMain.HluDataset.exports_fields.fields_countColumn))
                    multipleFields = true;

                // Add the required table to the list of sql tables in
                // the from clause.
                string currTable = _viewModelMain.DataBase.QualifyTableName(r.table_name);
                if (!fromList.Contains(currTable))
                {
                    fromList.Add(currTable);

                    var incidRelation = _viewModelMain.HluDataset.incid.ChildRelations.Cast<DataRelation>()
                        .Where(dr => dr.ChildTable.TableName == r.table_name);

                    if (incidRelation.Count() == 0)
                    {
                        fromClause.Append(currTable);
                    }
                    else
                    {
                        DataRelation incidRel = incidRelation.ElementAt(0);
                        if (firstJoin)
                            firstJoin = false;
                        else
                            fromClause.Insert(0, "(").Append(")");
                        fromClause.Append(RelationJoinClause("LEFT", currTable, true,
                            _viewModelMain.DataBase.QuoteIdentifier(
                            incidRel.ParentTable.TableName), incidRel, fromList));
                        leftJoined.Add(currTable);
                    }
                }

                // Get the relationships for the table/column if
                // a value from a lookup table is required.
                string fieldFormat = !r.IsNull(_viewModelMain.HluDataset.exports_fields.field_formatColumn) ? r.field_format : null;
                var relations = ((fieldFormat != null) && (fieldFormat == "Both" || fieldFormat == "Lookup")) ? _viewModelMain.HluDataRelations.Where(rel =>
                    rel.ChildTable.TableName == r.table_name && rel.ChildColumns
                    .Count(ch => ch.ColumnName == r.column_name) == 1) : new DataRelation[0];

                switch (relations.Count())
                {
                    case 0:     // If this field does not have any related lookup tables.

                        // Add the field to the sql target list.
                        targetList.Append(String.Format(",{0}.{1} AS {2}", currTable,
                            _viewModelMain.DataBase.QuoteIdentifier(r.column_name), r.field_name.Replace("<no>", "")));

                        // Add the field to the sql list of export table columns.
                        AddExportColumn(multipleFields ? r.fields_count : 0, r.table_name, r.column_name, r.field_name,
                            r.field_type, r.field_length, !r.IsNull(_viewModelMain.HluDataset.exports_fields.field_formatColumn) ? r.field_format : null,
                            ref exportFields);
                        break;
                    case 1:     // If this field has a related lookup tables.

                        DataRelation lutRelation = relations.ElementAt(0);
                        string parentTable = _viewModelMain.DataBase.QualifyTableName(lutRelation.ParentTable.TableName);

                        string parentTableAlias = tableAlias + tableAliasNum++;
                        fromList.Add(parentTable);

                        // Determine the related lookup table field name and
                        // field ordinal.
                        string lutFieldName;
                        int lutFieldOrdinal;
                        if ((r.table_name == _viewModelMain.HluDataset.incid_sources.TableName) && (Regex.IsMatch(r.column_name, @"(_id)", RegexOptions.IgnoreCase)))
                        {
                            lutFieldName = ViewModelWindowMain.LutSourceFieldName;
                            lutFieldOrdinal = ViewModelWindowMain.LutSourceFieldOrdinal - 1;
                        }
                        else if ((r.table_name == _viewModelMain.HluDataset.incid.TableName) && (Regex.IsMatch(r.column_name, @"(_user_id)", RegexOptions.IgnoreCase)))
                        {
                            lutFieldName = ViewModelWindowMain.LutUserFieldName;
                            lutFieldOrdinal = ViewModelWindowMain.LutUserFieldOrdinal - 1;
                        }
                        else
                        {
                            lutFieldName = ViewModelWindowMain.LutDescriptionFieldName;
                            lutFieldOrdinal = ViewModelWindowMain.LutDescriptionFieldOrdinal - 1;
                        }

                        // If the lookup table contains the required field name.
                        if (lutRelation.ParentTable.Columns.Contains(lutFieldName))
                        {
                            // If both the original field and it's corresponding lookup
                            // table field are required then add them both to the sql
                            // target list.
                            if ((fieldFormat != null) && (fieldFormat == "Both"))
                                // Add the corresponding lookup table field to the sql
                                // target list.
                                targetList.Append(String.Format(",{0}.{1} {5} {6} {5} {2}.{3} AS {4}",
                                    currTable,
                                    _viewModelMain.DataBase.QuoteIdentifier(r.column_name),
                                    parentTableAlias,
                                    _viewModelMain.DataBase.QuoteIdentifier(lutFieldName),
                                    r.field_name.Replace("<no>", ""),
                                    _viewModelMain.DataBase.ConcatenateOperator,
                                    _viewModelMain.DataBase.QuoteValue(" : ")));
                            else
                                // Add the corresponding lookup table field to the sql
                                // target list.
                                targetList.Append(String.Format(",{0}.{1} AS {2}",
                                    parentTableAlias,
                                    _viewModelMain.DataBase.QuoteIdentifier(lutFieldName),
                                    r.field_name.Replace("<no>", "")));

                            // Add the field to the sql list of export table columns.
                            AddExportColumn(multipleFields ? r.fields_count : 0, r.table_name, r.column_name, r.field_name,
                                r.field_type, r.field_length, !r.IsNull(_viewModelMain.HluDataset.exports_fields.field_formatColumn) ? r.field_format : null,
                                ref exportFields);
                        }
                        // If the lookup table does not contains the required field
                        // name, but does contain the required field ordinal.
                        else if (lutRelation.ParentTable.Columns.Count >= lutFieldOrdinal)
                        {
                            // If both the original field and it's corresponding lookup
                            // table field are required then add them both to the sql
                            // target list.
                            if ((fieldFormat != null) && (fieldFormat == "Both"))
                                // Add the corresponding lookup table field to the sql
                                // target list.
                                targetList.Append(String.Format(",{0}.{1} {5} {6} {5} {2}.{3} AS {4}",
                                    currTable,
                                    _viewModelMain.DataBase.QuoteIdentifier(r.column_name),
                                    parentTableAlias,
                                    _viewModelMain.DataBase.QuoteIdentifier(lutRelation.ParentTable.Columns[lutFieldOrdinal].ColumnName),
                                    r.field_name.Replace("<no>", ""),
                                    _viewModelMain.DataBase.ConcatenateOperator,
                                    _viewModelMain.DataBase.QuoteValue(" : ")));
                            else
                                // Add the corresponding lookup table field to the sql
                                // target list.
                                targetList.Append(String.Format(",{0}.{1} AS {2}",
                                    parentTableAlias,
                                    _viewModelMain.DataBase.QuoteIdentifier(lutRelation.ParentTable.Columns[lutFieldOrdinal].ColumnName),
                                    r.field_name.Replace("<no>", "")));

                            // Add the field to the sql list of export table columns.
                            AddExportColumn(multipleFields ? r.fields_count : 0, r.table_name, r.column_name, r.field_name,
                                r.field_type, r.field_length, !r.IsNull(_viewModelMain.HluDataset.exports_fields.field_formatColumn) ? r.field_format : null,
                                ref exportFields);
                        }
                        else
                        {
                            continue;
                        }

                        string joinType;
                        if (leftJoined.Contains(currTable))
                        {
                            joinType = "LEFT";
                            leftJoined.Add(parentTableAlias);
                        }
                        else
                        {
                            joinType = "INNER";
                        }

                        if (firstJoin)
                            firstJoin = false;
                        else
                            fromClause.Insert(0, "(").Append(")");

                        fromClause.Append(RelationJoinClause(joinType, currTable,
                            false, parentTableAlias, lutRelation, fromList));

                        break;
                }
            }

            // Create a new field map template with as many items
            // as there are input fields.
            fieldMapTemplate = new int[exportFields.Max(e => e.FieldOrdinal) + 1][];

            // Loop through all the export fields, adding them as columns
            // in the export table and adding them to the field map template.
            int fieldTotal = 0;
            int primaryKeyOrdinal = -1;
            _sourceIdOrdinal = -1;
            _sourceDateStartOrdinals = new List<int>();
            _sourceDateEndOrdinals = new List<int>();
            _sourceDateTypeOrdinals = new List<int>();
            List<int> dupFields = new List<int>();
            foreach (ExportField f in exportFields.OrderBy(f => f.FieldOrder))
            {
                // Create a new data column for the field.
                DataColumn c = new DataColumn(f.FieldName, f.FieldType);

                // If the field is an autonumber set the relevant
                // auto increment properties.
                if (f.AutoNum == true)
                {
                    c.AutoIncrement = true;
                    c.AutoIncrementSeed = 1;
                    c.AutoIncrementStep = 1;
                }

                // If the field has a maximum length (i.e. it is a text field)
                // then set the maximum length property.
                if (f.FieldLength > 0)
                    c.MaxLength = f.FieldLength;

                // Add the field as a new column in the export table.
                exportTable.Columns.Add(c);

                // If the field will not be sourced from the database.
                if (f.FieldOrdinal == -1)
                {
                    // Increment the total number of fields to be exported.
                    fieldTotal += 1;

                    // Skip adding the field to the field map template.
                    continue;
                }

                // If the field is not repeated and refers to the incid column
                // in the incid table.
                if ((f.FieldsCount == 0) && ((f.TableName == _viewModelMain.HluDataset.incid.TableName) &&
                    (f.ColumnName == _viewModelMain.HluDataset.incid.incidColumn.ColumnName)))
                {
                    // Store the input field position for use later 
                    // when exporting the data.
                    incidOrdinal = f.FieldOrdinal;

                    // Store the output field position for use later 
                    // as the primary index field ordinal.
                    primaryKeyOrdinal = fieldTotal;
                }

                // If the table is the incid_sources table.
                if (f.TableName == _viewModelMain.HluDataset.incid_sources.TableName)
                {
                    // Add the output field position to the list of fields
                    // that can have duplicate values.
                    dupFields.Add(fieldTotal);

                    // If the field refers to the source_id column then
                    // store the input field ordinal for use later as the
                    // unique incid_source field ordinal.
                    if (f.ColumnName == _viewModelMain.HluDataset.incid_sources.source_idColumn.ColumnName)
                        _sourceIdOrdinal = f.FieldOrdinal;

                    // If the field refers to the source_date_start column then
                    // store the input field ordinal for use later.
                    if (f.ColumnName == _viewModelMain.HluDataset.incid_sources.source_date_startColumn.ColumnName)
                        _sourceDateStartOrdinals.Add(f.FieldOrdinal);

                    // If the field refers to the source_date_end column then
                    // store the input field ordinal for use later.
                    if (f.ColumnName == _viewModelMain.HluDataset.incid_sources.source_date_endColumn.ColumnName)
                        _sourceDateEndOrdinals.Add(f.FieldOrdinal);

                    // If the field refers to the source_date_type column then
                    // store the input field ordinal for use later.
                    if (f.ColumnName == _viewModelMain.HluDataset.incid_sources.source_date_typeColumn.ColumnName)
                        _sourceDateTypeOrdinals.Add(f.FieldOrdinal);
                }

                // Set the field mapping for the current field ordinal.
                List<int> fieldMap;
                if ((fieldMapTemplate[f.FieldOrdinal] != null) && (f.AutoNum != true))
                    fieldMap = fieldMapTemplate[f.FieldOrdinal].ToList();
                else
                {
                    fieldMap = new List<int>();
                    fieldMap.Add(f.FieldOrdinal);
                }

                // Add the current field number to the field map for
                // this field ordinal.
                fieldMap.Add(fieldTotal);

                // Update the field map template for this field ordinal.
                fieldMapTemplate[f.FieldOrdinal] = fieldMap.ToArray();

                // Increment the total number of fields to be exported.
                fieldTotal += 1;
            }

            // Store which export fields can be allowed to have duplicate
            // values (i.e. the incid_source fields).
            dupsAllowed = dupFields.ToArray();

            // If any incid_source fields are in the export file.
            if ((exportFields.Count(f => f.TableName == _viewModelMain.HluDataset.incid_sources.TableName) != 0))
            {
                // Get the last input field ordinal.
                int lastFieldOrdinal = exportFields.Max(e => e.FieldOrdinal);

                // If the source_id column is not included then add
                // it so that different sources can be identified.
                if (_sourceIdOrdinal == -1)
                {
                    // Add the field to the input table.
                    targetList.Append(String.Format(",{0}.{1} AS {2}", _viewModelMain.HluDataset.incid_sources.TableName,
                        _viewModelMain.HluDataset.incid_sources.source_idColumn.ColumnName, _viewModelMain.HluDataset.incid_sources.source_idColumn.ColumnName));

                    // Store the input field ordinal for use
                    // later as the unique incid_source field ordinal.
                    _sourceIdOrdinal = lastFieldOrdinal += 1;
                }

                // If the source_date_start column is not included then add
                // it for use later.
                if ((_sourceDateStartOrdinals == null) || (_sourceDateStartOrdinals.Count() == 0))
                {
                    // Add the field to the input table.
                    targetList.Append(String.Format(",{0}.{1} AS {2}", _viewModelMain.HluDataset.incid_sources.TableName,
                        _viewModelMain.HluDataset.incid_sources.source_date_startColumn.ColumnName, _viewModelMain.HluDataset.incid_sources.source_date_startColumn.ColumnName));

                    // Store the input field ordinal for use later.
                    _sourceDateStartOrdinals.Add(lastFieldOrdinal += 1);
                }

                // If the source_date_end column is not included then add
                // it for use later.
                if ((_sourceDateEndOrdinals == null) || (_sourceDateEndOrdinals.Count() == 0))
                {
                    // Add the field to the input table.
                    targetList.Append(String.Format(",{0}.{1} AS {2}", _viewModelMain.HluDataset.incid_sources.TableName,
                        _viewModelMain.HluDataset.incid_sources.source_date_endColumn.ColumnName, _viewModelMain.HluDataset.incid_sources.source_date_endColumn.ColumnName));

                    // Store the input field ordinal for use later.
                    _sourceDateEndOrdinals.Add(lastFieldOrdinal += 1);
                }

                // If the source_date_type column is not included then add
                // it for use later.
                if ((_sourceDateTypeOrdinals == null) || (_sourceDateTypeOrdinals.Count() == 0))
                {
                    // Add the field to the input table.
                    targetList.Append(String.Format(",{0}.{1} AS {2}", _viewModelMain.HluDataset.incid_sources.TableName,
                        _viewModelMain.HluDataset.incid_sources.source_date_typeColumn.ColumnName, _viewModelMain.HluDataset.incid_sources.source_date_typeColumn.ColumnName));

                    // Store the input field ordinal for use later.
                    _sourceDateTypeOrdinals.Add(lastFieldOrdinal += 1);
                }
            }

            // Set the incid field as the primary key to the table.
            if (primaryKeyOrdinal != -1)
                exportTable.PrimaryKey = new DataColumn[] { exportTable.Columns[primaryKeyOrdinal] };

            // Remove the leading comma from the target list of fields.
            if (targetList.Length > 1) targetList.Remove(0, 1);
        }

        private string ExportMdb(string sql, List<ExportField> exportFields, DataTable exportTable, int incidOrdinal, int[]dupsAllowed, 
            ref int[][] fieldMap, out int exportRowCount)
        {
            exportRowCount = -1;
            DbOleDb dbOut = null;
            string tempPath = String.Empty;
            try { tempPath = Path.GetTempPath(); }
            catch { tempPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData); }
            tempPath = Path.Combine(tempPath, Path.GetFileNameWithoutExtension(Path.GetRandomFileName()) + ".mdb");

            try
            {
                if (File.Exists(tempPath)) File.Delete(tempPath);
                OdbcCP32 odbc = new OdbcCP32();
                odbc.CreateDatabase(tempPath);
                string connString = String.Format(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};", tempPath);
                string defaultSchema = "";
                bool promptPwd = false;
                dbOut = new DbOleDb(ref connString, ref defaultSchema, ref promptPwd,
                    Properties.Resources.PasswordMaskString, Settings.Default.UseAutomaticCommandBuilders,
                    true, Settings.Default.DbIsUnicode, Settings.Default.DbUseTimeZone, 255,
                    Settings.Default.DbBinaryLength, Settings.Default.DbTimePrecision,
                    Settings.Default.DbNumericPrecision, Settings.Default.DbNumericScale);
                dbOut.CreateTable(exportTable);

                DataSet datasetOut = new DataSet("Export");

                IDbDataAdapter adapterOut = dbOut.CreateAdapter(exportTable);
                adapterOut.MissingSchemaAction = MissingSchemaAction.AddWithKey;
                adapterOut.Fill(datasetOut);
                int[] pkOrdinals = exportTable.PrimaryKey.Select(c => c.Ordinal).ToArray();
                exportTable.PrimaryKey = pkOrdinals.Select(o => exportTable.Columns[o]).ToArray();
                adapterOut.TableMappings.Clear();
                adapterOut.TableMappings.Add(exportTable.TableName, datasetOut.Tables[0].TableName);

                // Save the export table structure so that the field
                // properties can be examined during the export.
                DataTable exportTable2 = exportTable;

                exportTable = datasetOut.Tables[0];

                DataRow exportRow = null;
                bool rowAdded = false;

                using (IDataReader reader = _viewModelMain.DataBase.ExecuteReader(sql, 
                    _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text))
                {
                    string currIncid = String.Empty;
                    string prevIncid = String.Empty;
                    object currSourceId = String.Empty;
                    object prevSourceId = String.Empty;
                    int fieldIndex = 1;

                    while (reader.Read())
                    {
                        // Get the current incid.
                        currIncid = reader.GetString(incidOrdinal);

                        // Get the current source id (or equivalent lookup table field).
                        if (_sourceIdOrdinal != -1)
                            currSourceId = reader.GetValue(_sourceIdOrdinal);

                        // If this incid is different to the last record's incid
                        // then process all the fields.
                        if (currIncid != prevIncid)
                        {
                            // Store the last incid.
                            prevIncid = currIncid;

                            // Reset the field map index to the start of the array.
                            fieldIndex = 1;

                            // Store the last source id.
                            prevSourceId = currSourceId;

                            // If the last export row has not been saved then
                            // save it now.
                            if (exportRow != null)
                            {
                                exportTable.Rows.Add(exportRow);
                                rowAdded = true;
                            }
                            
                            // Create a new export row ready for the next values.
                            exportRow = exportTable.NewRow();
                            rowAdded = false;

                            // Loop through all the fields in the field map
                            // to transfer the values from the input reader
                            // to the correct field in the export row.
                            for (int i = 0; i < fieldMap.GetLength(0); i++)
                            {
                                // If this field is not mapped from the input reader
                                // set the export table value to null.
                                if (fieldMap[i][0] == -1)
                                {
                                    exportRow[fieldMap[i][fieldIndex]] = null;
                                    continue;
                                }

                                // Get the properties for the current export field.
                                ExportField exportField = exportFields.Find(f => f.FieldOrdinal == i);

                                // Store the input value of the current column.
                                object inValue = reader.GetValue(fieldMap[i][0]);

                                // If the value is null then skip this field.
                                if (inValue == DBNull.Value)
                                    continue;

                                // Convert the input value to the output value data type and format.
                                object outValue;
                                outValue = ConvertInput(fieldMap[i][0], inValue, reader.GetFieldType(fieldMap[i][0]),
                                    exportTable2.Columns[fieldMap[i][fieldIndex]].DataType, (exportField != null) ? exportField.FieldFormat : null);

                                // If the value is not null.
                                if (outValue != null)
                                {
                                    // Get the maximum length of the column.
                                    int fieldLength = exportTable2.Columns[fieldMap[i][fieldIndex]].MaxLength;

                                    // If the maximum length of the column is shorter
                                    // than the value then truncate the value as it
                                    // is transferred  to the export row.
                                    if ((fieldLength != -1) && (fieldLength < outValue.ToString().Length))
                                        exportRow[fieldMap[i][fieldIndex]] = outValue.ToString().Substring(0, fieldLength);
                                    else
                                        exportRow[fieldMap[i][fieldIndex]] = outValue;
                                }
                            }
                        }
                        else
                        {
                            // Increment the index position for the field map.
                            fieldIndex += 1;

                            // Loop through all the fields in the field map
                            // to transfer the values from the input reader
                            // to the correct field in the export row.
                            for (int i = 0; i < fieldMap.GetLength(0); i++)
                            {
                                // Only process fields that have multiple outputs
                                // specified in the field map.
                                if (fieldIndex < fieldMap[i].Length)
                                {

                                    // Get the properties for the current export field.
                                    ExportField exportField = exportFields.Find(f => f.FieldOrdinal == i);

                                    // Store the input value of the current column.
                                    object inValue = reader.GetValue(fieldMap[i][0]);

                                    // If the value is null then skip this field.
                                    if (inValue == DBNull.Value)
                                        continue;

                                    // Convert the input value to the output value data type and format.
                                    object outValue;
                                    outValue = ConvertInput(fieldMap[i][0], inValue, reader.GetFieldType(fieldMap[i][0]),
                                        exportTable2.Columns[fieldMap[i][fieldIndex]].DataType, (exportField != null) ? exportField.FieldFormat : null);

                                    // Get the current and previous string values of the
                                    // current column so they can be compared later.
                                    object itemStr = outValue.ToString();
                                    object lastItemStr = exportRow[fieldMap[i][fieldIndex - 1]].ToString();

                                    // If the value is not null and the string value is different
                                    // to the last string value for this incid, or, the column is
                                    // allowed to have duplicates and the source is different
                                    // to the last source, then output the value.
                                    if ((!itemStr.Equals(lastItemStr) ||
                                        ((Array.IndexOf(dupsAllowed, fieldMap[i][fieldIndex]) != -1) && (currSourceId != prevSourceId))))
                                    {
                                        // Get the maximum length of the column.
                                        int fieldLength = exportTable2.Columns[fieldMap[i][fieldIndex]].MaxLength;

                                        // If the maximum length of the column is shorter
                                        // than the value then truncate the value as it
                                        // is transferred  to the export row.
                                        if ((fieldLength != -1) && (fieldLength < outValue.ToString().Length))
                                            exportRow[fieldMap[i][fieldIndex]] = outValue.ToString().Substring(0, fieldLength);
                                        else
                                            exportRow[fieldMap[i][fieldIndex]] = outValue;
                                    }
                                }
                            }
                        }
                        // Store the last source id.
                        prevSourceId = currSourceId;
                    }
                }

                // If the last export row has not been saved then
                // save it now.
                if (!rowAdded && (exportRow != null)) exportTable.Rows.Add(exportRow);

                exportRowCount = adapterOut.Update(datasetOut);

                return exportRowCount != -1 ? tempPath : null;
            }
            catch
            {
                if (File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); }
                    catch { _viewModelMain.ExportMdbs.Add(tempPath); }
                }
                return null;
            }
            finally
            {
                if ((dbOut != null) && (dbOut.Connection.State != ConnectionState.Closed))
                {
                    try { dbOut.Connection.Close(); }
                    catch { }
                }
            }
        }

        private object ConvertInput(int inOrdinal, object inValue, System.Type inType, System.Type outType, string outFormat)
        {
            // If the output field is a DateTime.
            if (outType == System.Type.GetType("System.DateTime"))
            {
                // If the input field is also a DateTime
                if (inType == System.Type.GetType("System.DateTime"))
                {
                    // Returns the value as a DateTime.
                    if (inValue is DateTime)
                        return inValue;
                    else
                        return null;
                }
                // If the input field is an integer and is part of
                // the source date.
                else if ((inType == System.Type.GetType("System.Int32")) &&
                    (_sourceDateStartOrdinals.Contains(inOrdinal) || _sourceDateEndOrdinals.Contains(inOrdinal)))
                {
                    // Convert the value to an integer.
                    int inInt = (int)inValue;

                    // Convert the value to a vague date instance.
                    string vt = "D";
                    Date.VagueDateInstance vd = new Date.VagueDateInstance(inInt, inInt, vt);

                    // If the vague date is invalid then return null.
                    if ((vd == null) || (vd.IsBad) || (vd.IsUnknown))
                        return null;
                    else
                    {
                        // If the vague date is valid then parse it into
                        // a date format.
                        string itemStr = Date.VagueDate.FromVagueDateInstance(vd, VagueDate.DateType.Vague);
                        DateTime inDate;
                        if (DateTime.TryParseExact(itemStr, "dd/MM/yyyy",
                            null, DateTimeStyles.None, out inDate))
                            return inDate;
                        else
                            return null;
                    }
                }
                else
                {
                    // Otherwise, try and parse the input value as
                    // if it was a date string and return the date value
                    // if it is valid, or the raw value if not.
                    string inStr = inValue.ToString();
                    DateTime inDate;
                    if (DateTime.TryParseExact(inStr, "dd/MM/yyyy",
                        null, DateTimeStyles.None, out inDate))
                        return inDate;
                    else
                        return null;
                }
            }
            // If the output field is a string and there is
            // a required output format.
            else if ((outType == System.Type.GetType("System.String")) &&
                (outFormat != null))
            {
                // If the input field is a DateTime field.
                if (inType == System.Type.GetType("System.DateTime"))
                {
                    // If the input value is a valid DateTime then
                    // convert it to a string of the required format.
                    if (inValue is DateTime)
                    {
                        DateTime inDate = (DateTime)inValue;
                        string inStr = inDate.ToString(outFormat);
                        if (inStr != null)
                            return inStr;
                        else
                            return null;

                        //// If the date string formats exactly then return
                        //// the formatted value.
                        //if (DateTime.TryParseExact(inDate.ToString(), outFormat,
                        //    null, DateTimeStyles.None, out inDate))
                        //    return inDate.ToString();
                        //else
                        //    return null;
                    }
                    else
                        return null;
                }
                // If the input field is an integer and is part of
                // the source date.
                else if ((inType == System.Type.GetType("System.Int32")) &&
                    (_sourceDateStartOrdinals.Contains(inOrdinal) || _sourceDateEndOrdinals.Contains(inOrdinal)))
                {
                    // Convert the value to an integer.
                    int inInt = (int)inValue;

                    // Convert the value to a vague date instance.
                    string vt = "D";
                    Date.VagueDateInstance vd = new Date.VagueDateInstance(inInt, inInt, vt);

                    // If the vague date is invalid then set the output
                    // field to null.
                    if ((vd == null) || (vd.IsBad))
                        return null;
                    else if (vd.IsUnknown)
                        return VagueDate.VagueDateTypes.Unknown.ToString();
                    else
                    {
                        // If the vague date is valid then parse it into
                        // a date format using the required format.
                        string inStr = Date.VagueDate.FromVagueDateInstance(vd, VagueDate.DateType.Vague);
                        DateTime inDate;
                        if (!DateTime.TryParseExact(inStr, "dd/MM/yyyy",
                            null, DateTimeStyles.None, out inDate))
                            return null;

                        // If the input value is a valid DateTime then
                        // convert it to a string of the required format.
                        string inStr2 = inDate.ToString(outFormat);
                        if (inStr2 != null)
                            return inStr2;
                        else
                            return null;
                    }
                }
                else
                {
                    // Otherwise, try and parse the input value as
                    // if it was a date string and return the value
                    // as a string if it is valid.
                    string inStr = inValue.ToString();

                    DateTime inDate;
                    if (DateTime.TryParse(inStr,
                        null, DateTimeStyles.None, out inDate))
                        return inDate.ToString();
                    else
                        return inValue;
                }
            }
            else
                return inValue;

        }

        /// <summary>
        /// Adds the export column to the export table.
        /// </summary>
        /// <param name="numFields">The number of occurrences of this field.</param>
        /// <param name="columnName">The name of the exported column.</param>
        /// <param name="dataType">The data type of the column.</param>
        /// <param name="maxLength">The maximum length of the column.</param>
        /// <param name="exportTable">The export table.</param>
        /// <param name="lutDescrColOrdinals">The lut description col ordinals.</param>
        private void AddExportColumn(int numFields, string tableName, string columnName, string fieldName, int fieldType, int maxLength,
            string fieldFormat, ref List<ExportField> exportFields)
        {
            Type dataType = null;
            int fieldLength = 0;
            bool autoNum = false;
            int attributeLength = 0;

            switch (fieldType)
            {
                case 3:     // Integer
                    dataType = System.Type.GetType("System.Int32");
                    attributeLength = 2;
                    break;
                case 6:     // Single
                    dataType = System.Type.GetType("System.Single");
                    attributeLength = 4;
                    break;
                case 7:     // Double
                    dataType = System.Type.GetType("System.Double");
                    attributeLength = 8;
                    break;
                case 8:     // Date/Time
                    dataType = System.Type.GetType("System.DateTime");
                    attributeLength = 8;
                    break;
                case 10:    // Text
                    dataType = System.Type.GetType("System.String");
                    fieldLength = maxLength;
                    attributeLength = maxLength;
                    break;
                case 99:    // Autonumber
                    dataType = System.Type.GetType("System.Int32");
                    autoNum = true;
                    attributeLength = 4;
                    break;
                default:
                    dataType = System.Type.GetType("System.String");
                    fieldLength = maxLength;
                    attributeLength = maxLength;
                    break;
            }

            // If this field has multiple occurrences.
            if (numFields > 0)
            {
                // Increment the number of times the same table has
                // been referenced.
                if (tableName == _lastTableName)
                    _tableCount += 1;
                else
                    _tableCount = 1;

                for (int i = 1; i <= numFields; i++)
                {
                    ExportField fld = new ExportField();

                    if (tableName.ToLower() == "<none>")
                        fld.FieldOrdinal = -1;
                    else
                        fld.FieldOrdinal = _fieldCount;
                    fld.TableName = tableName;
                    fld.ColumnName = columnName;

                    // Include the occurrence counter in the field name, either
                    // where the user chooses or at the end.
                    if (Regex.IsMatch(fieldName, @"(<no>)", RegexOptions.IgnoreCase))
                        fld.FieldName = fieldName.Replace("<no>", i.ToString());
                    else
                        fld.FieldName = String.Format("{0}_{1}", fieldName, i);
 
                    // If the table is the same as the last table then
                    // interweave the repeated fields.
                    int fieldNum;
                    int fieldCount = exportFields.Count + 1;
                    if (_tableCount == 1)
                        fieldNum = fieldCount * 10;
                    else
                        fieldNum = ((fieldCount - (numFields * (_tableCount - 1))) * 10) + (_tableCount - 1);

                    fld.FieldType = dataType;
                    fld.FieldOrder = fieldNum;
                    fld.FieldLength = fieldLength;
                    fld.FieldsCount = numFields;
                    fld.FieldFormat = fieldFormat;
                    fld.AutoNum = autoNum;

                    exportFields.Add(fld);

                    // Add the field attribute length to the running total.
                    _attributesLength += attributeLength;
                }
            }
            else
            {
                ExportField fld = new ExportField();
                if (tableName.ToLower() == "<none>")
                    fld.FieldOrdinal = -1;
                else
                    fld.FieldOrdinal = _fieldCount;
                fld.TableName = tableName;
                fld.ColumnName = columnName;
                fld.FieldName = fieldName;
                fld.FieldType = dataType;
                fld.FieldOrder = (exportFields.Count + 1) * 10;
                fld.FieldLength = fieldLength;
                fld.FieldsCount = numFields;
                fld.FieldFormat = fieldFormat;
                fld.AutoNum = autoNum;

                exportFields.Add(fld);

                // Add the field attribute length to the running total.
                _attributesLength += attributeLength;
            }

            // Store the last table referenced.
            _lastTableName = tableName;

            // Increment the field counter.
            if (tableName.ToLower() != "<none>")
                _fieldCount += 1;

        }

        private string RelationJoinClause(string joinType, string currTable, bool parentLeft,
            string parentTableAlias, DataRelation rel, List<string> fromList)
        {
            StringBuilder joinClausePart = new StringBuilder();

            for (int i = 0; i < rel.ParentColumns.Length; i++)
            {
                joinClausePart.Append(String.Format(" AND {0}.{2} = {1}.{3}", parentTableAlias, 
                    currTable, _viewModelMain.DataBase.QuoteIdentifier(rel.ParentColumns[i].ColumnName),
                    _viewModelMain.DataBase.QuoteIdentifier(rel.ChildColumns[i].ColumnName)));
            }

            if (parentTableAlias == _viewModelMain.DataBase.QuoteIdentifier(rel.ParentTable.TableName))
                parentTableAlias = String.Empty;
            else
                parentTableAlias = " " + parentTableAlias;

            string leftTable = String.Empty;
            string rightTable = string.Empty;
            if (parentLeft)
            {
                leftTable = _viewModelMain.DataBase.QuoteIdentifier(rel.ParentTable.TableName) + parentTableAlias;
                rightTable = currTable;
            }
            else
            {
                leftTable = currTable;
                rightTable = _viewModelMain.DataBase.QuoteIdentifier(rel.ParentTable.TableName) + parentTableAlias;
            }

            if (!fromList.Contains(currTable))
                return joinClausePart.Remove(0, 5).Insert(0, String.Format(" {0} {1} JOIN {2} ON ",
                    leftTable, joinType, rightTable)).ToString();
            else
                return joinClausePart.Remove(0, 5).Insert(0, String.Format(" {0} JOIN {1} ON ",
                    joinType, rightTable)).ToString();
        }

        private string GetTableAlias()
        {
            for (int i = 1; i < 5; i++)
            {
                for (int j = 122; j > 96; j--)
                {
                    char[] testCharArray = new char[i];
                    for (int k = 0; k < i; k++)
                        testCharArray[k] = (char)j;
                    string testString = new string(testCharArray);
                    if (_viewModelMain.HluDataset.Tables.Cast<DataTable>().Count(t => Regex.IsMatch(t.TableName,
                        testString + "[0-9]+", RegexOptions.IgnoreCase)) == 0)
                    {
                        return testString;
                    }
                }
            }
            return null;
        }
    }
}
