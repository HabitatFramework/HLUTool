// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2022 Greenspace Information for Greater London CIC
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
using System.Linq;
using System.Text;
using HLU.Data.Model;
using HLU.Properties;

namespace HLU.Data
{
    public class SecondaryHabitat : IDataErrorInfo
    {
        #region Fields

        private int _secondary_id;
        private string _incid;
        private bool _bulkUpdateMode;

        private string _secondary_habitat;
        private string _secondary_group;

        private string _incid_bak;
        private string _secondary_habitat_bak;
        private string _secondary_group_bak;

        string _error;
        private static IEnumerable<SecondaryHabitat> _secondaryHabitatList;

        #endregion

        #region ctor

        public SecondaryHabitat()
        {
            //TODO: Check
            _bulkUpdateMode = false;
            _secondary_id = -1; // arbitrary PK for a new row
        }

        public SecondaryHabitat(bool bulkUpdateMode, HluDataSet.incid_secondaryRow dataRow)
        {
            //TODO: Check
            _bulkUpdateMode = bulkUpdateMode;
            HluDataSet.incid_secondaryDataTable table = (HluDataSet.incid_secondaryDataTable)dataRow.Table;
            _secondary_id = dataRow.secondary_id;
            _incid = dataRow.incid;
            _secondary_habitat = dataRow.IsNull(table.secondaryColumn) ? null : dataRow.secondary;

            //TODO: Lookup code group???
            //_secondary_group = ???;
        }

        public SecondaryHabitat(bool bulkUpdateMode, HluDataSet.incid_secondaryRow dataRow, IEnumerable<SecondaryHabitat> shList)
        {
            _bulkUpdateMode = bulkUpdateMode;

            HluDataSet.incid_secondaryDataTable table = (HluDataSet.incid_secondaryDataTable)dataRow.Table;
            _secondary_id = dataRow.secondary_id;
            _incid = dataRow.incid;
            _secondary_habitat = dataRow.IsNull(table.secondaryColumn) ? null : dataRow.secondary;

            //_secondary_group = ???;
            _secondaryHabitatList = shList;
        }

        public SecondaryHabitat(bool bulkUpdateMode, object[] itemArray)
        {
            _bulkUpdateMode = bulkUpdateMode;

            Int32.TryParse(itemArray[0].ToString(), out _secondary_id);
            _incid = itemArray[1].ToString();
            _secondary_habitat = itemArray[2].ToString();

            //TODO: Lookup code group???
            //_secondary_group = ???;
        }

        public SecondaryHabitat(bool bulkUpdateMode, int secondary_id, string incid, string secondary_habitat,
            string secondary_group)
        {
            _bulkUpdateMode = bulkUpdateMode;

            _secondary_id = secondary_id;
            _incid = incid;
            _secondary_habitat = secondary_habitat;
            _secondary_group = secondary_group;
        }

        #endregion

        #region DataChanged

        //---------------------------------------------------------------------
        // CHANGED: CR2 (Apply button)
        // Create a handler so that updates to the secondary records can be picked
        // up back in the main window.
        //
        // declare the delegate since using the generic pattern
        public delegate void DataChangedEventHandler(bool Changed);

        // declare the event
        public event DataChangedEventHandler DataChanged;
        //---------------------------------------------------------------------

        #endregion

        #region Properties

        public static IEnumerable<SecondaryHabitat> SecondaryHabitatList
        {
            set { _secondaryHabitatList = value; }
        }

        public bool BulkUpdateMode
        {
            get { return _bulkUpdateMode; }
            set { _bulkUpdateMode = value; }
        }

        public bool IsAdded
        {
            get { return _secondary_id == -1; }
        }

        #region incid_secondaryRow

        public int secondary_id
        {
            get { return _secondary_id; }
            set { _secondary_id = value; }
        }

        public string incid
        {
            get { return _incid; }
            set { _incid = value; }
        }

        public string secondary_habitat
        {
            get { return _secondary_habitat; }
            set
            {
                _secondary_habitat = value;

                // Flag that the current record has changed so that the apply button
                // will appear.
                if (this.DataChanged != null)
                    this.DataChanged(true);
            }
        }

        //TODO: Needed???
        public string secondary_group
        {
            get { return _secondary_group; }
            set
            {
                _secondary_group = value;

                // Flag that the current record has changed so that the apply button
                // will appear.
                if (this.DataChanged != null)
                    this.DataChanged(true);
            }
        }

        #endregion

        #endregion

        #region Public methods

        public object[] ToItemArray()
        {
            return new object[] { _secondary_id, _incid, _secondary_habitat, _secondary_group };
        }

        public object[] ToItemArray(int secondaryID, string incid)
        {
            return new object[] { secondaryID, incid, _secondary_habitat, _secondary_group };
        }

        #endregion

        #region Validation

        public bool IsDuplicate
        {
            get
            {
                return _secondaryHabitatList != null && _secondaryHabitatList.Count(sh => sh.secondary_habitat == this.secondary_habitat) > 0;
            }
        }

        public bool IsValid()
        {
            return ValidateRow();
        }

        public bool IsValid(bool bulkUpdateMode)
        {
            return String.IsNullOrEmpty(ValidateRow(bulkUpdateMode, _secondary_id,
                _incid, _secondary_habitat));
        }

        public bool IsValid(bool bulkUpdateMode, HluDataSet.incid_secondaryRow r)
        {
            return String.IsNullOrEmpty(ValidateRow(bulkUpdateMode, r.secondary_id,
                r.incid, r.secondary));
        }

        public string ErrorMessages { get { return _error.ToString(); } }

        public static bool ValidateRow(bool bulkUpdateMode, HluDataSet.incid_secondaryRow r)
        {
            return ValidateRow(bulkUpdateMode, r.secondary_id, r.incid, r.secondary) == null;
        }

        public static bool ValidateRow(bool bulkUpdateMode, HluDataSet.incid_secondaryRow r,
            IEnumerable<SecondaryHabitat> secondaryHabitatList)
        {
            return ValidateRow(bulkUpdateMode, r.secondary_id, r.incid, r.secondary) == null;
        }

        private static string ValidateRow(bool _bulkUpdateMode, int secondary_id, string incid,
            string secondary_habitat)
        {
            StringBuilder sbError = new StringBuilder();

            //TODO: Check
            if ((secondary_id != -1) && String.IsNullOrEmpty(incid))
                sbError.Append(Environment.NewLine).Append("INCID is a mandatory field");

            if (String.IsNullOrEmpty(secondary_habitat))
                sbError.Append(Environment.NewLine).Append("Secondary habitat is a mandatory field");

            //TODO: Needed???
            ////if ((_bapEnvironmentList != null) && (_bapEnvironmentList.Count(b => b.bap_habitat == bap_habitat) > 1))
            ////    sbError.Append(Environment.NewLine).Append("Duplicate priority environment");

            return sbError.Length > 0 ? sbError.Remove(0, 1).ToString() : null;
        }

        private bool ValidateRow()
        {
            _error = ValidateRow(_bulkUpdateMode, secondary_id, incid, secondary_habitat);
            return _error == null;
        }

        #endregion

        #region IDataErrorInfo Members

        string IDataErrorInfo.Error
        {
            get
            {
                ValidateRow();
                return _error;
            }
        }

        string IDataErrorInfo.this[string columnName]
        {
            get
            {
                string error = null;

                switch (columnName)
                {
                    case "incid":
                        if ((secondary_id != -1) && String.IsNullOrEmpty(incid))
                        {
                            _incid = _incid_bak;
                            return "INCID is a mandatory field";
                        }
                        _incid_bak = _incid;
                        break;
                    //TODO: Check
                    case "secondary_habitat":
                        if (String.IsNullOrEmpty(secondary_habitat))
                        {
                            return "Secondary habitat is a mandatory field";
                        }
                        else if ((_secondaryHabitatList != null) && (_secondaryHabitatList.Count(b => b.secondary_habitat == secondary_habitat) > 1))
                        {
                            return "Duplicate secondary habitat";
                        }
                        _secondary_habitat_bak = _secondary_habitat;
                        break;
                    case "secondary_group":
                        if (!_bulkUpdateMode && String.IsNullOrEmpty(secondary_group))
                        {
                            return "Secondary group is a mandatory field";
                        }
                        _secondary_group_bak = _secondary_group;
                        break;
                }

                return error;
            }
        }

        #endregion
    }
}
