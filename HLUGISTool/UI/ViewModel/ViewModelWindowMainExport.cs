// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
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
        public static HluDataSet HluDatasetStatic = null;

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

        //---------------------------------------------------------------------
        // CHANGED: CR14 (Exporting IHS codes or descriptions)
        // Enable users to specify if individual fields should be
        // exported with descriptions, rather than the whole export,
        // by moving this option to the exports_fields table.
        //
        private void _viewModelExport_RequestClose(int exportID, bool selectedOnly)
        //---------------------------------------------------------------------
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
        /// Exports the combined GIS and database data using the
        /// specified export format.
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
                _viewModelMain.ChangeCursor(Cursors.Wait, "Creating export table ...");

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

                //---------------------------------------------------------------------
                // FIX: 041 Check the selected export format contains the incid column.
                //
                // Exit if there is no incid field for this format.
                if (_viewModelMain.HluDataset.exports_fields.Count(f => f.column_name == _viewModelMain.IncidTable.incidColumn.ColumnName) == 0)
                    throw new Exception(String.Format("The export format '{0}' does not contain the column 'incid'",
                        _viewModelMain.HluDataset.exports.FindByexport_id(userExportId).export_name));
                //---------------------------------------------------------------------

                // Build a new export data table and also determine the
                // number of output fields, a string of the output field
                // names, the from SQL clause, and the incid
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

                // Set the export filter conditions, depending if all the
                // records are to be exported or only the selected features.
                List<List<SqlFilterCondition>> exportFilter = null;
                if (selectedOnly)
                {
                    // If the where clause is not already set then get it
                    // using the GIS selection.
                    if ((_viewModelMain.IncidSelectionWhereClause == null) &&
                        (_viewModelMain.GisSelection != null) && (_viewModelMain.GisSelection.Rows.Count > 0))
                    {
                        _viewModelMain.IncidSelectionWhereClause = ViewModelWindowMainHelpers.GisSelectionToWhereClause(
                            _viewModelMain.GisSelection.Select(), _viewModelMain.GisIDColumnOrdinals,
                            250, _viewModelMain.HluDataset.incid);

                        // Set the export filter to the where clause.
                        exportFilter = _viewModelMain.IncidSelectionWhereClause;
                    }
                    else
                    {
                        // Combine all the where clauses into a single list (so
                        // that it can be re-chunked later into larger chunks
                        // than the standard chunk based on the IncidPageSize.
                        exportFilter = new List<List<SqlFilterCondition>>();
                        exportFilter.Add(_viewModelMain.IncidSelectionWhereClause.SelectMany(l => l).ToList());
                    }
                }
                else
                {
                    SqlFilterCondition cond = new SqlFilterCondition("AND",
                        _viewModelMain.IncidTable, _viewModelMain.IncidTable.incidColumn, null);
                    cond.Operator = "IS NOT NULL";
                    exportFilter = new List<List<SqlFilterCondition>>(new List<SqlFilterCondition>[] { 
                        new List<SqlFilterCondition>(new SqlFilterCondition[] { cond }) });
                }

                //---------------------------------------------------------------------
                // FIX: 042 Warn the user when the export will be very large.
                //
                // Count the number of incids to be exported.
                int rowCount = 0;
                if (selectedOnly)
                    rowCount = _viewModelMain.IncidsSelectedMapCount;
                else
                    rowCount = _viewModelMain.IncidRowCount(false);
                
                // Warn the user if the export is very large.
                if (rowCount > 5000)
                {
                    MessageBoxResult userResponse = MessageBoxResult.No;
                    userResponse = MessageBox.Show("This export operation may take some time.\n\nDo you wish to proceed?", "HLU: Export",
                        MessageBoxButton.YesNo, MessageBoxImage.Question);

                    // Cancel the export.
                    if (userResponse != MessageBoxResult.Yes)
                        return;
                }
                //---------------------------------------------------------------------

                // Export the attribute data to a temporary database.
                int exportRowCount;
                tempPath = ExportMdb(targetList.ToString(), fromClause.ToString(), exportFilter,
                    _viewModelMain.DataBase, exportFields, exportTable, incidOrdinal, dupsAllowed, ref fieldMapTemplate, out exportRowCount);

                if (String.IsNullOrEmpty(tempPath))
                    throw new Exception("Error creating the export table");

                if (exportRowCount < 1)
                    throw new Exception("Export query did not retrieve any rows");

                _viewModelMain.ChangeCursor(Cursors.Wait, "Exporting from GIS ...");

                // Call the GIS application export method to join the
                // temporary attribute data to the GIS feature layer
                // and save them as a new GIS layer.
                _viewModelMain.GISApplication.Export(tempPath, exportTable.TableName, _attributesLength, selectedOnly);

                // Remove the current record filter.
                //_viewModelMain.IncidSelection = null;
                //_viewModelMain.GisSelection = null;
                //_viewModelMain.OnPropertyChanged("IsFiltered");
                //_viewModelMain.OnPropertyChanged("StatusIncid");
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
            int fieldLength = 0;
            _attributesLength = 0;
            _tableCount = 0;

            //
            foreach (HluDataSet.exports_fieldsRow r in
                _viewModelMain.HluDataset.exports_fields.OrderBy(r => r.field_ordinal))
            {
                // Get the field length of the source table/column.
                fieldLength = GetFieldLength(r.table_name, r.column_ordinal);

                //---------------------------------------------------------------------
                // FIX: 043 Enable new 'empty' fields to be included in exports.
                if (r.table_name.ToLower() == "<none>")
                {
                    // Override the source field length(s) if an export
                    // field length has been set.
                    if (!r.IsNull(_viewModelMain.HluDataset.exports_fields.field_lengthColumn) &&
                        r.field_length > 0)
                        fieldLength = r.field_length;

                    AddExportColumn(0, r.table_name, r.column_name, r.field_name,
                        r.field_type, fieldLength, !r.IsNull(_viewModelMain.HluDataset.exports_fields.field_formatColumn) ? r.field_format : null,
                        ref exportFields);
                    continue;
                }
                //---------------------------------------------------------------------

                // Determine if this field is to be output multiple times,
                // once for each row in the relevant table up to the
                // maximum fields_count value.
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

                // Get the relationships for the table/column if a
                // value from a lookup table is required.
                string fieldFormat = !r.IsNull(_viewModelMain.HluDataset.exports_fields.field_formatColumn) ? r.field_format : null;
                //---------------------------------------------------------------------
                // CHANGED: CR14 (Exporting IHS codes or descriptions)
                // Enable users to specify if individual fields should be
                // exported with descriptions in the exports_fields table.
                //
                var relations = ((fieldFormat != null) && (fieldFormat == "Both" || fieldFormat == "Lookup")) ? _viewModelMain.HluDataRelations.Where(rel =>
                    rel.ChildTable.TableName == r.table_name && rel.ChildColumns
                    .Count(ch => ch.ColumnName == r.column_name) == 1) : new DataRelation[0];
                //---------------------------------------------------------------------

                switch (relations.Count())
                {
                    case 0:     // If this field does not have any related lookup tables.

                        // Add the field to the sql target list.
                        targetList.Append(String.Format(",{0}.{1} AS {2}", currTable,
                            _viewModelMain.DataBase.QuoteIdentifier(r.column_name), r.field_name.Replace("<no>", "")));

                        //---------------------------------------------------------------------
                        // FIX: 044 Enable text field lengths to be specified in
                        // the export format.
                        //
                        // Override the source field length(s) if an export
                        // field length has been set.
                        if (!r.IsNull(_viewModelMain.HluDataset.exports_fields.field_lengthColumn) &&
                            r.field_length > 0)
                            fieldLength = r.field_length;
                        //---------------------------------------------------------------------

                        // Add the field to the sql list of export table columns.
                        AddExportColumn(multipleFields ? r.fields_count : 0, r.table_name, r.column_name, r.field_name,
                            r.field_type, fieldLength,
                            !r.IsNull(_viewModelMain.HluDataset.exports_fields.field_formatColumn) ? r.field_format : String.Empty,
                            ref exportFields);
                        break;
                    case 1:     // If this field has a related lookup table.

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

                        // Get the list of columns for the lookup table.
                        DataColumn[] lutColumns = new DataColumn[lutRelation.ParentTable.Columns.Count];
                        lutRelation.ParentTable.Columns.CopyTo(lutColumns, 0);

                        // If the lookup table contains the required field name.
                        if (lutRelation.ParentTable.Columns.Contains(lutFieldName))
                        {
                            //---------------------------------------------------------------------
                            // CHANGED: CR15 (Concatenate IHS codes and descriptions)
                            // Enable users to specify if individual fields should be
                            // exported with both codes and descriptions concatenated
                            // together.
                            //
                            // If both the original field and it's corresponding lookup
                            // table field are required then add them both to the sql
                            // target list.
                            if ((fieldFormat != null) && (fieldFormat == "Both"))
                            {
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

                                // Set the field length of the export field to the source
                                // field length plus the lookup table field length plus 3
                                // for the concatenation string length.
                                fieldLength += lutColumns.First(c => c.ColumnName == lutFieldName).MaxLength + 3;
                            }
                            //---------------------------------------------------------------------
                            else
                            {
                                // Add the corresponding lookup table field to the sql
                                // target list.
                                targetList.Append(String.Format(",{0}.{1} AS {2}",
                                    parentTableAlias,
                                    _viewModelMain.DataBase.QuoteIdentifier(lutFieldName),
                                    r.field_name.Replace("<no>", "")));

                                // Set the field length of the lookup table field.
                                fieldLength = lutColumns.First(c => c.ColumnName == lutFieldName).MaxLength;
                            }

                            //---------------------------------------------------------------------
                            // FIX: 044 Enable text field lengths to be specified in
                            // the export format.
                            //
                            // Override the source field length(s) if an export
                            // field length has been set.
                            if (!r.IsNull(_viewModelMain.HluDataset.exports_fields.field_lengthColumn) &&
                                r.field_length > 0)
                                fieldLength = r.field_length;
                            //---------------------------------------------------------------------

                            // Add the field to the sql list of export table columns.
                            AddExportColumn(multipleFields ? r.fields_count : 0, r.table_name, r.column_name, r.field_name,
                                r.field_type, fieldLength, !r.IsNull(_viewModelMain.HluDataset.exports_fields.field_formatColumn) ? r.field_format : String.Empty,
                                ref exportFields);
                        }
                        // If the lookup table does not contains the required field
                        // name, but does contain the required field ordinal.
                        else if (lutRelation.ParentTable.Columns.Count >= lutFieldOrdinal)
                        {
                            //---------------------------------------------------------------------
                            // CHANGED: CR15 (Concatenate IHS codes and descriptions)
                            // Enable users to specify if individual fields should be
                            // exported with both codes and descriptions concatenated
                            // together.
                            //
                            // If both the original field and it's corresponding lookup
                            // table field are required then add them both to the sql
                            // target list.
                            if ((fieldFormat != null) && (fieldFormat == "Both"))
                            {
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

                                // Set the field length of the lookup table field.
                                fieldLength = lutColumns.First(c => c.ColumnName == lutFieldName).MaxLength;
                            }
                            //---------------------------------------------------------------------
                            else
                            {
                                // Add the corresponding lookup table field to the sql
                                // target list.
                                targetList.Append(String.Format(",{0}.{1} AS {2}",
                                    parentTableAlias,
                                    _viewModelMain.DataBase.QuoteIdentifier(lutRelation.ParentTable.Columns[lutFieldOrdinal].ColumnName),
                                    r.field_name.Replace("<no>", "")));

                                // Set the field length of the lookup table field.
                                fieldLength = lutColumns.First(c => c.ColumnName == lutFieldName).MaxLength;
                            }

                            //---------------------------------------------------------------------
                            // FIX: 044 Enable text field lengths to be specified in
                            // the export format.
                            //
                            // Override the source field length(s) if an export
                            // field length has been set.
                            if (!r.IsNull(_viewModelMain.HluDataset.exports_fields.field_lengthColumn) &&
                                r.field_length > 0)
                                fieldLength = r.field_length;
                            //---------------------------------------------------------------------

                            // Add the field to the sql list of export table columns.
                            AddExportColumn(multipleFields ? r.fields_count : 0, r.table_name, r.column_name, r.field_name,
                                r.field_type, fieldLength, !r.IsNull(_viewModelMain.HluDataset.exports_fields.field_formatColumn) ? r.field_format : null,
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

            //---------------------------------------------------------------------
            // FIX: 045 Interweave multiple record fields from the same
            // table together.
            //
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
                if (f.AutoNum == true) c.AutoIncrement = true;

                // If the field is a text field and has a maximum length
                // then set the maximum length property.
                if ((f.FieldType == System.Type.GetType("System.String")) &&
                    (f.FieldLength != null) && (f.FieldLength > 0)) c.MaxLength = f.FieldLength;

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

                    // If the field refers to the source_id column then store
                    // the input field ordinal for use later as the unique
                    // incid_source field ordinal.
                    if (f.ColumnName == _viewModelMain.HluDataset.incid_sources.source_idColumn.ColumnName)
                        _sourceIdOrdinal = f.FieldOrdinal;
                    // If the field refers to the source_date_start column then
                    // store the input field ordinal for use later.
                    else if (f.ColumnName == _viewModelMain.HluDataset.incid_sources.source_date_startColumn.ColumnName)
                        _sourceDateStartOrdinals.Add(f.FieldOrdinal);
                    // If the field refers to the source_date_end column then
                    // store the input field ordinal for use later.
                    else if (f.ColumnName == _viewModelMain.HluDataset.incid_sources.source_date_endColumn.ColumnName)
                        _sourceDateEndOrdinals.Add(f.FieldOrdinal);
                    // If the field refers to the source_date_type column then
                    // store the input field ordinal for use later.
                    else if (f.ColumnName == _viewModelMain.HluDataset.incid_sources.source_date_typeColumn.ColumnName)
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
            //---------------------------------------------------------------------

            // Store which export fields can be allowed to have duplicate
            // values (i.e. the incid_source fields).
            dupsAllowed = dupFields.ToArray();

            // If any incid_source fields are in the export file.
            if ((exportFields.Count(f => f.TableName == _viewModelMain.HluDataset.incid_sources.TableName) != 0))
            {
                // Get the last input field ordinal.
                int lastFieldOrdinal = exportFields.Max(e => e.FieldOrdinal);

                //---------------------------------------------------------------------
                // FIX: 046 Don't export multiple source details for the
                // same source.
                //
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
                //---------------------------------------------------------------------

                //---------------------------------------------------------------------
                // CHANGED: CR17 (Exporting date fields)
                // Store all of the source date fields for use later when
                // formatting the attribute data.
                //
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
                //---------------------------------------------------------------------
            }

            // Set the incid field as the primary key to the table.
            if (primaryKeyOrdinal != -1)
                exportTable.PrimaryKey = new DataColumn[] { exportTable.Columns[primaryKeyOrdinal] };

            // Remove the leading comma from the target list of fields.
            if (targetList.Length > 1) targetList.Remove(0, 1);
        }

        private string ExportMdb(string targetListStr, string fromClauseStr, List<List<SqlFilterCondition>> exportFilter,
            DbBase dataBase, List<ExportField> exportFields, DataTable exportTable, int incidOrdinal, int[]dupsAllowed, 
            ref int[][] fieldMap, out int exportRowCount)
        {
            exportRowCount = -1;
            int outputRowCount = 0;
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

                // Return a null database path if the table cannot be created.
                if (!dbOut.CreateTable(exportTable))
                    return null;
                
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

                //---------------------------------------------------------------------
                // CHANGED: CR13 (Export features performance)
                // Improve the performance as much as possible when
                // creating the attribute database.
                //
                // Turn off notifications and index maintenance whilst
                // inserting the records.
                exportTable.BeginLoadData();
                //---------------------------------------------------------------------

                // If there is only one long list then chunk
                // it up into smaller lists.
                if (exportFilter.Count == 1)
                {
                    try
                    {
                        List<SqlFilterCondition> whereCond = new List<SqlFilterCondition>();
                        whereCond = exportFilter[0];
                        exportFilter = whereCond.ChunkClause(250).ToList();
                    }
                    catch { }
                }

                //---------------------------------------------------------------------
                // FIX: 047 Break exporting attributes into chunks to avoid errors
                // with excessive sql lengths.
                //
                outputRowCount = 0;
                exportRowCount = 0;
                for (int j = 0; j < exportFilter.Count; j++)
                {
                    DataRow exportRow = exportTable.NewRow();
                    bool rowAdded = false;

                    // Union the constituent parts of the export query
                    // together into a single SQL string.
                    string sql = ScratchDb.UnionQuery(targetListStr, fromClauseStr,
                        incidOrdinal, exportFilter[j], dataBase);

                    // Execute the sql to retrieve the records.
                    using (IDataReader reader = _viewModelMain.DataBase.ExecuteReader(sql,
                        _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text))
                    {
                        string currIncid = String.Empty;
                        string prevIncid = String.Empty;
                        object currSourceId = String.Empty;
                        object prevSourceId = String.Empty;
                        int currSourceDateStart = 0;
                        int currSourceDateEnd = 0;
                        string currSourceDateType = String.Empty;
                        int fieldIndex = 1;

                        // Read each record and process the contents.
                        while (reader.Read())
                        {
                            // Get the current incid.
                            currIncid = reader.GetString(incidOrdinal);

                            //---------------------------------------------------------------------
                            // FIX: 046 Don't export multiple source details for the
                            // same source.
                            //
                            // Get the current source id (or equivalent lookup table field).
                            if (_sourceIdOrdinal != -1)
                                currSourceId = reader.GetValue(_sourceIdOrdinal);
                            //---------------------------------------------------------------------

                            //---------------------------------------------------------------------
                            // CHANGED: CR17 (Exporting date fields)
                            // Store all of the source date fields for use later when
                            // formatting the attribute data.
                            //
                            // Get the current source date start.
                            if ((_sourceDateStartOrdinals.Count() > 0) &&
                                !reader.IsDBNull(_sourceDateStartOrdinals[0]))
                                currSourceDateStart = reader.GetInt32(_sourceDateStartOrdinals[0]);

                            // Get the current source date type.
                            if ((_sourceDateEndOrdinals.Count() > 0) &&
                                !reader.IsDBNull(_sourceDateEndOrdinals[0]))
                                currSourceDateEnd = reader.GetInt32(_sourceDateEndOrdinals[0]);

                            // Get the current source date type.
                            if ((_sourceDateTypeOrdinals.Count() > 0) &&
                                !reader.IsDBNull(_sourceDateTypeOrdinals[0]))
                                currSourceDateType = reader.GetString(_sourceDateTypeOrdinals[0]);
                            //---------------------------------------------------------------------

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

                                // If the last export row has not been added then
                                // add it now.
                                if (!exportRow.IsNull(fieldMap[0][1]))
                                {
                                    exportTable.Rows.Add(exportRow);
                                    rowAdded = true;

                                    // Increment the output row count.
                                    outputRowCount += 1;

                                    // Commit the outputs and update the export row count
                                    // every 10,000 records to avoid excessive memory use.
                                    if (outputRowCount >= 10000)
                                    {
                                        exportRowCount += adapterOut.Update(datasetOut);
                                        outputRowCount = 0;
                                    }
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

                                    //---------------------------------------------------------------------
                                    // FIX: 048 Enable fields to be exported using a different
                                    // data type.
                                    //
                                    // Convert the input value to the output value data type and format.
                                    object outValue;
                                    outValue = ConvertInput(fieldMap[i][0], inValue, reader.GetFieldType(fieldMap[i][0]),
                                        exportTable2.Columns[fieldMap[i][fieldIndex]].DataType,
                                        (exportField != null) ? exportField.FieldFormat : null,
                                        currSourceDateStart, currSourceDateEnd, currSourceDateType);
                                    //---------------------------------------------------------------------

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

                                        //---------------------------------------------------------------------
                                        // FIX: 048 Enable fields to be exported using a different
                                        // data type.
                                        //
                                        // Convert the input value to the output value data type and format.
                                        object outValue;
                                        outValue = ConvertInput(fieldMap[i][0], inValue, reader.GetFieldType(fieldMap[i][0]),
                                            exportTable2.Columns[fieldMap[i][fieldIndex]].DataType,
                                            (exportField != null) ? exportField.FieldFormat : null,
                                            currSourceDateStart, currSourceDateEnd, currSourceDateType);
                                        //---------------------------------------------------------------------

                                        // Get the current and previous string values of the
                                        // current column so they can be compared later.
                                        string itemStr;
                                        if (outValue != null)
                                            itemStr = outValue.ToString();
                                        else
                                            itemStr = string.Empty;
                                        object lastItemStr = exportRow[fieldMap[i][fieldIndex - 1]].ToString();

                                        //---------------------------------------------------------------------
                                        // FIX: 046 Don't export multiple source details for the
                                        // same source.
                                        //
                                        // If the value is not null and the string value is different
                                        // to the last string value for this incid, or, the column is
                                        // allowed to have duplicates and the source is different
                                        // to the last source, then output the value.
                                        if ((!itemStr.Equals(lastItemStr) ||
                                            ((Array.IndexOf(dupsAllowed, fieldMap[i][fieldIndex]) != -1) && (currSourceId != prevSourceId))))
                                        //---------------------------------------------------------------------
                                        {
                                            // Get the maximum length of the column.
                                            int fieldLength = exportTable2.Columns[fieldMap[i][fieldIndex]].MaxLength;

                                            // If the maximum length of the column is shorter
                                            // than the value then truncate the value as it
                                            // is transferred  to the export row.
                                            if ((fieldLength != -1) && (fieldLength < itemStr.Length))
                                                exportRow[fieldMap[i][fieldIndex]] = itemStr.Substring(0, fieldLength);
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
                    if (!rowAdded && (!exportRow.IsNull(fieldMap[0][1])))
                    {
                        exportTable.Rows.Add(exportRow);
                        rowAdded = true;

                        // Increment the output row count.
                        outputRowCount += 1;
                    }
                }
                //---------------------------------------------------------------------

                // Commit any remaining outputs and update the export row count.
                exportRowCount += adapterOut.Update(datasetOut);

                // Turn notifications and index maintenance back on again.
                exportTable.EndLoadData();

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

        /// <summary>
        /// Gets the length of the original source field.
        /// </summary>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="columnOrdinal">The column ordinal.</param>
        /// <returns></returns>
        private int GetFieldLength(string tableName, int columnOrdinal)
        {
            int fieldLength = 0;
            var relations = _viewModelMain.HluDataRelations.Where(rel =>
                    rel.ParentTable.TableName == HluDatasetStatic.incid.TableName);

            // Get a list of all the incid related tables (including the
            // incid table itself.
            List<DataTable> tables;
            tables = _viewModelMain.HluDataset.incid.ChildRelations
                .Cast<DataRelation>().Select(r => r.ChildTable).ToList();
            tables.Add(_viewModelMain.HluDataset.incid);

            foreach (DataTable t in tables)
            {
                if (t.TableName == tableName)
                {
                    DataColumn[] columns = new DataColumn[t.Columns.Count];
                    t.Columns.CopyTo(columns, 0);

                    // Get the field length.
                    fieldLength = columns[columnOrdinal - 1].MaxLength;
                    break;
                }
            }

            return fieldLength;
        }

        //---------------------------------------------------------------------
        // FIX: 048 Enable fields to be exported using a different
        // data type.
        //
        /// <summary>
        /// Converts the input field into the output field, applying any
        /// required formatting as appropriate.
        /// </summary>
        /// <param name="inOrdinal">The input field ordinal.</param>
        /// <param name="inValue">The input field value.</param>
        /// <param name="inType">Data type of the input field.</param>
        /// <param name="outType">Date type of the output field.</param>
        /// <param name="outFormat">The required output field format.</param>
        /// <param name="sourceDateStart">The source date start.</param>
        /// <param name="sourceDateEnd">The source date end.</param>
        /// <param name="sourceDateType">The source date type.</param>
        /// <returns></returns>
        private object ConvertInput(int inOrdinal, object inValue, System.Type inType,
            System.Type outType, string outFormat, int sourceDateStart, int sourceDateEnd, string sourceDateType)
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
                //---------------------------------------------------------------------
                // CHANGED: CR17 (Exporting date fields)
                // Convert source dates into a text field with the required
                // date format.
                // If the input field is an integer and is part of
                // the source date.
                else if ((inType == System.Type.GetType("System.Int32")) &&
                    (_sourceDateStartOrdinals.Contains(inOrdinal) || _sourceDateEndOrdinals.Contains(inOrdinal)))
                {
                    // Convert the value to an integer.
                    int inInt = (int)inValue;

                    // Convert the value to a vague date instance.
                    Date.VagueDateInstance vd = new Date.VagueDateInstance(sourceDateStart, sourceDateEnd, sourceDateType);

                    // If the vague date is invalid then set the output
                    // field to null.
                    if ((vd == null) || (vd.IsBad))
                        return null;
                    else if (vd.IsUnknown)
                        return VagueDate.VagueDateTypes.Unknown.ToString();
                    else
                    {
                        // If the output format is blank then format the date according
                        // to the source date type.
                        if (String.IsNullOrEmpty(outFormat))
                        {
                            if (_sourceDateStartOrdinals.Contains(inOrdinal))
                            {
                                return Date.VagueDate.FromVagueDateInstance(vd, VagueDate.DateType.Start);
                            }
                            else if (_sourceDateEndOrdinals.Contains(inOrdinal))
                            {
                                return Date.VagueDate.FromVagueDateInstance(vd, VagueDate.DateType.End);
                            }
                            else
                                return null;
                        }
                        // If the output format is a vague date type then format it
                        // like that.
                        else if (VagueDate.FromCode(outFormat) != VagueDate.VagueDateTypes.Unknown)
                        {
                            // If the field is a start date then use the first character of
                            // the vague date type.
                            if (_sourceDateStartOrdinals.Contains(inOrdinal))
                            {
                                // Set the date type applicable to the source date.
                                string dateType = outFormat.Substring(0, 1);

                                return Date.VagueDate.FromVagueDateInstance(new Date.VagueDateInstance(sourceDateStart, sourceDateEnd, dateType),
                                    VagueDate.DateType.Start);
                            }
                            // If the field is an end date then use the last character of
                            // the vague date type.
                            else if (_sourceDateEndOrdinals.Contains(inOrdinal))
                            {
                                // Set the date type applicable to the source date.
                                string dateType = outFormat.Length == 1 ? outFormat + outFormat : outFormat;

                                return Date.VagueDate.FromVagueDateInstance(new Date.VagueDateInstance(sourceDateStart, sourceDateEnd, dateType),
                                    VagueDate.DateType.End);
                            }
                            else
                                return null;
                        }
                        // If the output format is not a vague date type then format it
                        // as if it is a standard date format.
                        else
                        {
                            // Parse the date into a date format using the output format.
                            string inStr = Date.VagueDate.FromVagueDateInstance(new Date.VagueDateInstance(inInt, inInt, "D"), VagueDate.DateType.Vague);
                            DateTime inDate;
                            if (!DateTime.TryParseExact(inStr, "dd/MM/yyyy", null, DateTimeStyles.None, out inDate))
                                return null;

                            // Convert the DateTime to a string of the output format.
                            string outDate;
                            outDate = inDate.ToString(outFormat);

                            // Parse the formatted date back into a date using the
                            // output format to check it is a valid date.
                            DateTime inDateAgain;
                            if (!DateTime.TryParseExact(outDate, outFormat, null, DateTimeStyles.None, out inDateAgain) ||
                                (inDate != inDateAgain))
                                return null;
                            else
                                return outDate;
                        }
                    }
                }
                //---------------------------------------------------------------------
                else
                {
                    // Otherwise, try and parse the input value as
                    // if it was a date string and return the value
                    // as a string if it is valid.
                    string inStr = inValue.ToString();

                    DateTime inDate;
                    if (DateTime.TryParse(inStr, null, DateTimeStyles.None, out inDate))
                        return inDate.ToString();
                    else
                        return inValue;
                }
            }
            else
                return inValue;

        }
        //---------------------------------------------------------------------

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

            // Increment each time a different table is referenced.
            if (tableName != _lastTableName)
                _tableCount += 1;

            //---------------------------------------------------------------------
            // FIX: 048 Enable fields to be exported using a different
            // data type.
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
                    if (maxLength > 0)
                    {
                        fieldLength = Math.Min(maxLength, 254);
                        attributeLength = fieldLength;
                    }
                    else
                    {
                        fieldLength = 254;
                        attributeLength = fieldLength;
                    }
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
            //---------------------------------------------------------------------

            // If this field has multiple occurrences.
            if (numFields > 0)
            {
                int fieldCount = exportFields.Count + 1;

                for (int i = 1; i <= numFields; i++)
                {
                    ExportField fld = new ExportField();

                    //---------------------------------------------------------------------
                    // FIX: 043 Enable new 'empty' fields to be included in exports.
                    if (tableName.ToLower() == "<none>")
                        fld.FieldOrdinal = -1;
                    else
                        fld.FieldOrdinal = _fieldCount;
                    //---------------------------------------------------------------------
                    fld.TableName = tableName;
                    fld.ColumnName = columnName;
                    //---------------------------------------------------------------------
                    // FIX: 049 Enable the multi-record counter to be inserted
                    // 'within' the export field name.
                    //
                    // Include the occurrence counter in the field name, either
                    // where the user chooses or at the end.
                    if (Regex.IsMatch(fieldName, @"(<no>)", RegexOptions.IgnoreCase))
                        fld.FieldName = fieldName.Replace("<no>", i.ToString());
                    else
                        fld.FieldName = String.Format("{0}_{1}", fieldName, i);
                    //---------------------------------------------------------------------
                    fld.FieldType = dataType;
                    //---------------------------------------------------------------------
                    // FIX: 045 Interweave multiple record fields from the same
                    // table together.
                    fld.FieldOrder = (_tableCount * 1000) + (i * 100) + fieldCount;
                    //---------------------------------------------------------------------
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

                //---------------------------------------------------------------------
                // FIX: 043 Enable new 'empty' fields to be included in exports.
                if (tableName.ToLower() == "<none>")
                    fld.FieldOrdinal = -1;
                else
                    fld.FieldOrdinal = _fieldCount;
                //---------------------------------------------------------------------
                fld.TableName = tableName;
                fld.ColumnName = columnName;
                fld.FieldName = fieldName;
                fld.FieldType = dataType;
                //---------------------------------------------------------------------
                // FIX: 045 Interweave multiple record fields from the same
                // table together.
                fld.FieldOrder = (_tableCount * 1000) + exportFields.Count + 1;
                //---------------------------------------------------------------------
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
