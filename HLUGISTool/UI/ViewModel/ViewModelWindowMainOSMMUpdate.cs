// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2019 London & South East Record Centres (LaSER)
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
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Input;
using HLU.Data;
using HLU.Data.Model;
using HLU.Data.Model.HluDataSetTableAdapters;
using HLU.GISApplication;
using HLU.Properties;

namespace HLU.UI.ViewModel
{
    class ViewModelWindowMainOSMMUpdate
    {
        private ViewModelWindowMain _viewModelMain;
        private HluDataSet.incid_ihs_matrixDataTable _ihsMatrixTable = new HluDataSet.incid_ihs_matrixDataTable();
        private HluDataSet.incid_ihs_formationDataTable _ihsFormationTable = new HluDataSet.incid_ihs_formationDataTable();
        private HluDataSet.incid_ihs_managementDataTable _ihsManagementTable = new HluDataSet.incid_ihs_managementDataTable();
        private HluDataSet.incid_ihs_complexDataTable _ihsComplexTable = new HluDataSet.incid_ihs_complexDataTable();

        public ViewModelWindowMainOSMMUpdate(ViewModelWindowMain viewModelMain)
        {
            _viewModelMain = viewModelMain;
        }

        public void StartOSMMUpdate()
        {
            _viewModelMain.OSMMUpdateMode = true;
            _viewModelMain.TabIhsControlsEnabled = false;
            _viewModelMain.TabDetailsControlsEnabled = false;
            _viewModelMain.TabSourcesControlsEnabled = false;

            // Reset the incid and map selections and move
            // to the first incid in the database.
            _viewModelMain.ClearFilter(true);

            // Count the incid_osmm_update rows for the selected flag
            _viewModelMain.CountOSMMUpdates();

            _viewModelMain.RefreshAll();
        }

        /// <summary>
        /// Writes changes made to current incid_osmm_updates back to database.
        /// </summary>
        internal bool OSMMUpdate(int updateStatus)
        {
            _viewModelMain.DataBase.BeginTransaction(true, IsolationLevel.ReadCommitted);

            try
            {
                _viewModelMain.ChangeCursor(Cursors.Wait, "Saving ...");

                // Only update DateTime fields to whole seconds.
                // Fractions of a second can cause rounding differences when
                // comparing DateTime fields later in some databases.
                DateTime currDtTm = DateTime.Now;
                DateTime nowDtTm = new DateTime(currDtTm.Year, currDtTm.Month, currDtTm.Day, currDtTm.Hour, currDtTm.Minute, currDtTm.Second, DateTimeKind.Local);
                //---------------------------------------------------------------------
                _viewModelMain.IncidOSMMUpdatesRows[0].last_modified_date = nowDtTm;
                _viewModelMain.IncidOSMMUpdatesRows[0].last_modified_user_id = _viewModelMain.UserID;

                // Determine the status flag as required
                int newStatus_Flag = _viewModelMain.IncidOSMMUpdatesRows[0].status_flag;
                switch (updateStatus)
                {
                    case 1:     // Skip update
                        newStatus_Flag = newStatus_Flag + 1;
                        break;
                    case -1:    // Accept update
                        newStatus_Flag = -1;
                        break;
                    case -99:   // Reject update
                        newStatus_Flag = -99;
                        break;
                }

                // Set the status flag
                _viewModelMain.IncidOSMMUpdatesRows[0].status_flag = newStatus_Flag;

                // Update the incid_osmm_updates table
                if (_viewModelMain.HluTableAdapterManager.incid_osmm_updatesTableAdapter.Update(
                    (HluDataSet.incid_osmm_updatesDataTable)_viewModelMain.HluDataset.incid_osmm_updates.GetChanges()) == -1)
                    throw new Exception(String.Format("Failed to update '{0}' table.",
                        _viewModelMain.HluDataset.incid_osmm_updates.TableName));

                // Commit the changes
                _viewModelMain.DataBase.CommitTransaction();
                _viewModelMain.HluDataset.AcceptChanges();
                _viewModelMain.Saved = true;

                return true;
            }
            catch (Exception ex)
            {
                _viewModelMain.DataBase.RollbackTransaction();
                if (_viewModelMain.HaveGisApp)
                {
                    _viewModelMain.Saved = false;
                    MessageBox.Show("Your changes could not be saved. The error message returned was:\n\n" +
                        ex.Message, "HLU: Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
                else
                {
                    return true;
                }
            }
            finally
            {
                _viewModelMain.SavingAttempted = true;
                _viewModelMain.Saving = false;
                _viewModelMain.ChangeCursor(Cursors.Arrow, null);
            }
        }

        private R[] FilterUpdateRows<T, R>(R[] rows)
            where T : DataTable
            where R : DataRow
        {
            if ((rows == null) || (rows.Length == 0)) return rows;

            T table = (T)rows[0].Table;

            var q = from rel in table.ParentRelations.Cast<DataRelation>()
                    where rel.ParentTable.TableName.ToLower().StartsWith("lut_") &&
                          rel.ParentColumns.Length == 1
                    select rel;

            var lookup = (from c in table.Columns.Cast<DataColumn>()
                          let p = from rel in q
                                  where rel.ChildColumns.Length == 1 && rel.ChildColumns[0].Ordinal == c.Ordinal
                                  select rel.ParentColumns[0]
                          select new
                          {
                              ChildColumnOrdinal = c.Ordinal,
                              ParentColumnOrdinal = p.Count() != 0 ? p.ElementAt(0).Ordinal : -1,
                              ParentTable = p.Count() != 0 ? p.ElementAt(0).Table.TableName : String.Empty
                          }).ToArray();

            List<R> newRows = new List<R>(rows.Length);

            for (int i = 0; i < rows.Length; i++)
            {
                if (rows[i] == null) continue;
                bool add = true;
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    if ((lookup[j].ParentColumnOrdinal != -1) && (_viewModelMain.HluDataset.Tables[lookup[j].ParentTable]
                        .AsEnumerable().Count(r => r[lookup[j].ParentColumnOrdinal].Equals(rows[i][j])) == 0))
                    {
                        add = false;
                        break;
                    }
                }
                if (add) newRows.Add(rows[i]);
            }

            return newRows.ToArray();
        }

        private T CloneUpdateRows<T, R>(R[] rows, string incid)
            where T : DataTable, new()
            where R : DataRow
        {
            T newRows = new T();
            int incidOrdinal = newRows.Columns[_viewModelMain.HluDataset.incid.incidColumn.ColumnName].Ordinal;

            for (int i = 0; i < rows.Length; i++)
            {
                object[] itemArray = new object[rows[i].ItemArray.Length];
                Array.Copy(rows[i].ItemArray, itemArray, itemArray.Length);
                itemArray[incidOrdinal] = incid;
                R rm = (R)newRows.NewRow();
                rm.ItemArray = itemArray;
                newRows.Rows.Add(rm);
            }

            return newRows;
        }

        /// <summary>
        /// Uses the bap_habitat code to identify existing records and distinguish them from new ones to be inserted.
        /// Therefore, a bap_habitat must always be entered, even for existing records for which only other attributes
        /// are meant to be updated. 
        /// </summary>
        /// <param name="currIncid"></param>
        /// <param name="ihsHabitat"></param>
        /// <param name="incidBapTable"></param>
        /// <param name="ihsMatrixRows"></param>
        /// <param name="ihsFormationRows"></param>
        /// <param name="ihsManagementRows"></param>
        /// <param name="ihsComplexRows"></param>
        /// <param name="deleteExtraRows"></param>
        private void BulkUpdateBap(string currIncid, string ihsHabitat,
            HluDataSet.incid_bapDataTable incidBapTable,
            HluDataSet.incid_ihs_matrixRow[] ihsMatrixRows,
            HluDataSet.incid_ihs_formationRow[] ihsFormationRows,
            HluDataSet.incid_ihs_managementRow[] ihsManagementRows,
            HluDataSet.incid_ihs_complexRow[] ihsComplexRows,
            bool deleteExtraRows)
        {
            var mx = ihsMatrixRows.Where(r => r.RowState != DataRowState.Deleted).Select(r => r.matrix);
            string[] ihsMatrixVals = mx.Concat(new string[3 - mx.Count()]).ToArray();
            var fo = ihsFormationRows.Where(r => r.RowState != DataRowState.Deleted).Select(r => r.formation);
            string[] ihsFormationVals = fo.Concat(new string[2 - fo.Count()]).ToArray();
            var mg = ihsManagementRows.Where(r => r.RowState != DataRowState.Deleted).Select(r => r.management);
            string[] ihsManagementVals = mg.Concat(new string[2 - mg.Count()]).ToArray();
            var cx = ihsComplexRows.Where(r => r.RowState != DataRowState.Deleted).Select(r => r.complex);
            string[] ihsComplexVals = cx.Concat(new string[2 - cx.Count()]).ToArray();

            IEnumerable<string> primaryBap = _viewModelMain.PrimaryBapEnvironments(ihsHabitat, 
                ihsMatrixVals[0], ihsMatrixVals[1], ihsMatrixVals[2], ihsFormationVals[0], ihsFormationVals[1],
                ihsManagementVals[0], ihsManagementVals[1], ihsComplexVals[0], ihsComplexVals[1]);

            int[] skipOrdinals = new int[3] { 
                _viewModelMain.HluDataset.incid_bap.bap_idColumn.Ordinal, 
                _viewModelMain.HluDataset.incid_bap.incidColumn.Ordinal, 
                _viewModelMain.HluDataset.incid_bap.bap_habitatColumn.Ordinal };

            List<HluDataSet.incid_bapRow> updateRows = new List<HluDataSet.incid_bapRow>();
            HluDataSet.incid_bapRow updateRow;

            // BAP environments from UI
            IEnumerable<BapEnvironment> beUI = from b in _viewModelMain.IncidBapRowsAuto.Concat(_viewModelMain.IncidBapRowsUser)
                                               group b by b.bap_habitat into habs
                                               select habs.First();

            foreach (BapEnvironment be in beUI)
            {
                // BAP environments from database
                IEnumerable<HluDataSet.incid_bapRow> dbRows =
                    incidBapTable.Where(r => r.bap_habitat == be.bap_habitat);

                bool isSecondary = !primaryBap.Contains(be.bap_habitat);
                if (isSecondary) be.MakeSecondary();

                switch (dbRows.Count())
                {
                    case 0: // insert newly added BAP environments
                        HluDataSet.incid_bapRow newRow = _viewModelMain.IncidBapTable.Newincid_bapRow();
                        newRow.ItemArray = be.ToItemArray(_viewModelMain.RecIDs.NextIncidBapId, currIncid);
                        if (be.IsValid(false, isSecondary, newRow)) // reset bulk update mode for full validation of a new row
                            _viewModelMain.HluTableAdapterManager.incid_bapTableAdapter.Insert(newRow);
                        break;
                    case 1: // update existing row
                        updateRow = dbRows.ElementAt(0);
                        object[] itemArray = be.ToItemArray();
                        for (int i = 0; i < itemArray.Length; i++)
                        {
                            if ((itemArray[i] != null) && (Array.IndexOf(skipOrdinals, i) == -1))
                                updateRow[i] = itemArray[i];
                        }
                        updateRows.Add(updateRow);
                        break;
                    default: // impossible if rules properly enforced
                        break;
                }
            }

            incidBapTable.Where(r => !primaryBap.Contains(r.bap_habitat)).ToList().ForEach(delegate(HluDataSet.incid_bapRow r)
            {
                updateRows.Add(BapEnvironment.MakeSecondary(r));
            });

            if (updateRows.Count > 0)
            {
                if (_viewModelMain.HluTableAdapterManager.incid_bapTableAdapter.Update(updateRows.ToArray()) == -1)
                    throw new Exception(String.Format("Failed to update {0} table.", _viewModelMain.HluDataset.incid_bap.TableName));
            }

            // delete non-primary BAP environments from DB that are not in _viewModelMain.IncidBapRowsUser
            if (deleteExtraRows)
            {
                var delRows = incidBapTable.Where(r => !primaryBap.Contains(r.bap_habitat) && 
                    _viewModelMain.IncidBapRowsUser.Count(be => be.bap_habitat == r.bap_habitat) == 0);
                foreach (HluDataSet.incid_bapRow r in delRows)
                    _viewModelMain.HluTableAdapterManager.incid_bapTableAdapter.Delete(r);
            }
        }

        public void CancelOSMMUpdate()
        {
            OSMMUpdateResetControls();
        }

        private void OSMMUpdateResetControls()
        {
            _viewModelMain.OSMMUpdateMode = null;
            _viewModelMain.IncidCurrentRowIndex = 1;
            _viewModelMain.OSMMUpdateMode = false;
            
            _viewModelMain.TabIhsControlsEnabled = true;
            _viewModelMain.TabDetailsControlsEnabled = true;
            _viewModelMain.TabSourcesControlsEnabled = true;

            _viewModelMain.RefreshAll();
            _viewModelMain.ChangeCursor(Cursors.Arrow, String.Empty);
        }
    }
}
