﻿// HLUTool is used to view and maintain habitat and land use GIS data.
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
    //---------------------------------------------------------------------
    // CHANGED: CR49 Process proposed OSMM Updates
    // Functionality to process proposed OSMM Updates.
    //    
    class ViewModelWindowMainOSMMUpdate
    {
        private ViewModelWindowMain _viewModelMain;

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
            _viewModelMain.ApplyOSMMUpdatesFilter(null, null, null, null);

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
                _viewModelMain.ChangeCursor(Cursors.Wait, "Updating ...");

                // Only update DateTime fields to whole seconds.
                // Fractions of a second can cause rounding differences when
                // comparing DateTime fields later in some databases.
                DateTime currDtTm = DateTime.Now;
                DateTime nowDtTm = new DateTime(currDtTm.Year, currDtTm.Month, currDtTm.Day, currDtTm.Hour, currDtTm.Minute, currDtTm.Second, DateTimeKind.Local);
                //---------------------------------------------------------------------
                _viewModelMain.IncidOSMMUpdatesRows[0].last_modified_date = nowDtTm;
                _viewModelMain.IncidOSMMUpdatesRows[0].last_modified_user_id = _viewModelMain.UserID;

                // Determine the status flag as required
                int newStatus = _viewModelMain.IncidOSMMUpdatesRows[0].status;
                switch (updateStatus)
                {
                    case 1:     // Skip update
                        if (newStatus > 0)
                            newStatus = newStatus + 1;
                        break;
                    case 0:    // Accept update
                        newStatus = 0;
                        break;
                    case -99:   // Reject update
                        newStatus = -99;
                        break;
                }

                // Set the status flag
                _viewModelMain.IncidOSMMUpdatesRows[0].status = newStatus;

                // Update the incid_osmm_updates table
                if (_viewModelMain.HluTableAdapterManager.incid_osmm_updatesTableAdapter.Update(
                    (HluDataSet.incid_osmm_updatesDataTable)_viewModelMain.HluDataset.incid_osmm_updates.GetChanges()) == -1)
                    throw new Exception(String.Format("Failed to update '{0}' table.",
                        _viewModelMain.HluDataset.incid_osmm_updates.TableName));

                // Commit the changes
                _viewModelMain.DataBase.CommitTransaction();
                _viewModelMain.HluDataset.AcceptChanges();
                _viewModelMain.Saved = true;

                // Move to the next Incid
                _viewModelMain.IncidCurrentRowIndex += 1;

                return true;
            }
            catch (Exception ex)
            {
                _viewModelMain.DataBase.RollbackTransaction();
                MessageBox.Show("OSMM Update failed. The error message returned was:\n\n" +
                    ex.Message, "HLU: OSMM Update", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
            finally
            {
                _viewModelMain.ChangeCursor(Cursors.Arrow, null);
            }
        }

        /// <summary>
        /// Writes changes made to all the remaining selected incid_osmm_updates
        /// back to database.
        /// </summary>
        internal void OSMMUpdateAll(int updateStatus)
        {
            _viewModelMain.ChangeCursor(Cursors.Wait, "Updating all ...");

            _viewModelMain.DataBase.BeginTransaction(true, IsolationLevel.ReadCommitted);

            try
            {
                // Get the incid number
                int currIncidNum = RecordIds.IncidNumber(_viewModelMain.IncidOSMMUpdatesRows[0].incid);
                
                // Apply the updates to the current incid and all following incids
                BulkIncidOSMMUpdates(updateStatus, currIncidNum);

                _viewModelMain.DataBase.CommitTransaction();
                _viewModelMain.HluDataset.AcceptChanges();

                MessageBox.Show("OSMM update succeeded.", "HLU: OSMM Update",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                // Move beyond the end of the Incids (to show they have
                // all been processed)
                _viewModelMain.IncidCurrentRowIndex = _viewModelMain.IncidSelection.Rows.Count + 1;
            }
            catch (Exception ex)
            {
                _viewModelMain.DataBase.RollbackTransaction();
                MessageBox.Show("OSMM Update Failed. The error message returned was:\n\n" +
                    ex.Message, "HLU: OSMM Update", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _viewModelMain.ChangeCursor(Cursors.Arrow, String.Empty);
            }
        }

        private void BulkIncidOSMMUpdates(int updateStatus, int fromIncidNum)
        {
            // Get the incid column number
            int incidOrdinal =
                _viewModelMain.IncidSelection.Columns[_viewModelMain.DataBase.ColumnAlias(_viewModelMain.HluDataset.incid.incidColumn)].Ordinal;

            // Get the column names
            string incidColumn = _viewModelMain.HluDataset.incid_osmm_updates.incidColumn.ColumnName;
            string statusColumn = _viewModelMain.HluDataset.incid_osmm_updates.statusColumn.ColumnName;
            string last_modified_dateColumn = _viewModelMain.HluDataset.incid_osmm_updates.last_modified_dateColumn.ColumnName;
            string last_modified_user_idColumn = _viewModelMain.HluDataset.incid_osmm_updates.last_modified_user_idColumn.ColumnName;

            // Get the current date/time
            // Only update DateTime fields to whole seconds.
            // Fractions of a second can cause rounding differences when
            // comparing DateTime fields later in some databases.
            DateTime currDtTm = DateTime.Now;
            DateTime nowDtTm = new DateTime(currDtTm.Year, currDtTm.Month, currDtTm.Day, currDtTm.Hour, currDtTm.Minute, currDtTm.Second, DateTimeKind.Local);

            // Get the current userid
            string user_id = _viewModelMain.UserID;

            // Loop through all rows in the selection
            foreach (DataRow r in _viewModelMain.IncidSelection.Rows)
            {
                // Get the incid of the current row
                string currIncid = r[incidOrdinal].ToString();

                // Get the incid number
                int currIncidNum = RecordIds.IncidNumber(currIncid);

                // Check the incid is to be updated
                if (currIncidNum >= fromIncidNum)
                {
                    // Update the incid for the current row
                    if (_viewModelMain.DataBase.ExecuteNonQuery(String.Format("UPDATE {0} SET {1} = {2}, {3} = {4}, {5} = {6} WHERE {7} = {8}",
                        _viewModelMain.DataBase.QualifyTableName(_viewModelMain.HluDataset.incid_osmm_updates.TableName),
                        _viewModelMain.DataBase.QuoteIdentifier(statusColumn),
                        updateStatus,
                        _viewModelMain.DataBase.QuoteIdentifier(last_modified_dateColumn),
                        _viewModelMain.DataBase.QuoteValue(nowDtTm),
                        _viewModelMain.DataBase.QuoteIdentifier(last_modified_user_idColumn),
                        _viewModelMain.DataBase.QuoteValue(user_id),
                        _viewModelMain.DataBase.QuoteIdentifier(incidColumn),
                        _viewModelMain.DataBase.QuoteValue(currIncid)),
                        _viewModelMain.DataBase.Connection.ConnectionTimeout, CommandType.Text) == -1)
                        throw new Exception("Failed to update incid_osmm_updates table.");
                }
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
            _viewModelMain.TabItemHistoryEnabled = true;

            _viewModelMain.RefreshAll();
            _viewModelMain.ChangeCursor(Cursors.Arrow, String.Empty);
        }
    }
    //---------------------------------------------------------------------
}