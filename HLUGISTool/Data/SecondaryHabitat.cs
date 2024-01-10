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
    public class SecondaryHabitat : IDataErrorInfo, ICloneable
    {
        #region Fields

        private int _secondary_id;
        private string _incid;
        private string _secondary_habitat;
        private int _secondary_habitat_int;
        private string _secondary_group;

        private bool _bulkUpdateMode;
        private string _incid_bak;
        private string _secondary_habitat_bak;

        string _error;
        private static IEnumerable<SecondaryHabitat> _secondaryHabitatList;
        private static IEnumerable<string> _validSecondaryCodes;
        private static Dictionary<string, String> _secondaryGroupCodes;

        public static int _secondaryCodeValidation;

        #endregion

        #region ctor

        public SecondaryHabitat()
        {
            _bulkUpdateMode = false;

            _secondary_id = -1; // arbitrary PK for a new row
        }

        public SecondaryHabitat(bool bulkUpdateMode, HluDataSet.incid_secondaryRow dataRow)
        {
            _bulkUpdateMode = bulkUpdateMode;

            HluDataSet.incid_secondaryDataTable table = (HluDataSet.incid_secondaryDataTable)dataRow.Table;
            _secondary_id = dataRow.secondary_id;
            _incid = dataRow.incid;
            _secondary_habitat = dataRow.IsNull(table.secondaryColumn) ? null : dataRow.secondary;
            int secondary_habitat_int;
            if (int.TryParse(_secondary_habitat, out secondary_habitat_int))
                _secondary_habitat_int = secondary_habitat_int;
            else
                _secondary_habitat_int = 0;
            _secondary_group = dataRow.secondary_group;
        }

        public SecondaryHabitat(bool bulkUpdateMode, HluDataSet.incid_secondaryRow dataRow, IEnumerable<SecondaryHabitat> shList)
        {
            _bulkUpdateMode = bulkUpdateMode;

            HluDataSet.incid_secondaryDataTable table = (HluDataSet.incid_secondaryDataTable)dataRow.Table;
            _secondary_id = dataRow.secondary_id;
            _incid = dataRow.incid;
            _secondary_habitat = dataRow.IsNull(table.secondaryColumn) ? null : dataRow.secondary;
            int secondary_habitat_int;
            if (int.TryParse(_secondary_habitat, out secondary_habitat_int))
                _secondary_habitat_int = secondary_habitat_int;
            else
                _secondary_habitat_int = 0;
            _secondary_group = dataRow.secondary_group;
        }

        public SecondaryHabitat(bool bulkUpdateMode, object[] itemArray)
        {
            _bulkUpdateMode = bulkUpdateMode;

            Int32.TryParse(itemArray[0].ToString(), out _secondary_id);
            _incid = itemArray[1].ToString();
            _secondary_habitat = itemArray[2].ToString();
            int secondary_habitat_int;
            if (int.TryParse(_secondary_habitat, out secondary_habitat_int))
                _secondary_habitat_int = secondary_habitat_int;
            else
                _secondary_habitat_int = 0;
            _secondary_group = itemArray[3].ToString();
        }

        public SecondaryHabitat(bool bulkUpdateMode, int secondary_id, string incid, string secondary_habitat, string secondary_group)
        {
            _bulkUpdateMode = bulkUpdateMode;

            _secondary_id = secondary_id;
            _incid = incid;
            _secondary_habitat = secondary_habitat;
            int secondary_habitat_int;
            if (int.TryParse(_secondary_habitat, out secondary_habitat_int))
                _secondary_habitat_int = secondary_habitat_int;
            else
                _secondary_habitat_int = 0;
            _secondary_group = secondary_group;
        }

        public SecondaryHabitat(SecondaryHabitat inSH)
        {
            _bulkUpdateMode = false;

            _secondary_id = -1; // arbitrary PK for a new row
            _incid = null;
            _secondary_habitat = inSH.secondary_habitat;
            _secondary_habitat_int = inSH.secondary_habitat_int;
            _secondary_group = inSH.secondary_group;
        }

        public object Clone()
        {
            return new SecondaryHabitat(this);
        }

        #endregion

        #region DataChanged

        // Create a handler so that updates to the secondary records can be picked
        // up back in the main window.
        //
        // declare the delegate since using the generic pattern
        public delegate void DataChangedEventHandler(bool Changed);

        // declare the event
        public event DataChangedEventHandler DataChanged;

        #endregion

        #region Properties

        public static IEnumerable<SecondaryHabitat> SecondaryHabitatList
        {
            get { return _secondaryHabitatList; }
            set {
                _secondaryHabitatList = value;

                //TODO: Secondaries - Needed?
                //for (int i = 0; i < _secondaryHabitatList.Count(); i++)
                //{
                //    SecondaryHabitat sh = new SecondaryHabitat();
                //    sh = _secondaryHabitatList.ElementAt(i);
                //}
            }
        }

        public static Dictionary<string, String> SecondaryGroupCodes
        {
            get { return _secondaryGroupCodes; }
            set { _secondaryGroupCodes = value; }
        }

        public static IEnumerable<string> ValidSecondaryCodes
        {
            get { return _validSecondaryCodes; }
            set { _validSecondaryCodes = value; }
        }

        public static int SecondaryCodeValidation
        {
            get { return _secondaryCodeValidation; }
            set { _secondaryCodeValidation = value; }
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

                //TODO: Secondaries - Needed?
                // Flag that the current record has changed so that the apply button
                // will appear.
                if (this.DataChanged != null)
                    this.DataChanged(true);
            }
        }

        public int secondary_habitat_int
        {
            get { return _secondary_habitat_int; }
            set { _secondary_habitat_int = value; }
        }

        public string secondary_group
        {
            get { return _secondary_group; }
            set { _secondary_group = value; }
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
                _incid, _secondary_habitat, _secondary_group));
        }

        public bool IsValid(bool bulkUpdateMode, HluDataSet.incid_secondaryRow r)
        {
            return String.IsNullOrEmpty(ValidateRow(bulkUpdateMode, r.secondary_id,
                r.incid, r.secondary, r.secondary_group));
        }

        public string ErrorMessages { get { return _error.ToString(); } }

        public static bool ValidateRow(bool bulkUpdateMode, HluDataSet.incid_secondaryRow r)
        {
            return ValidateRow(bulkUpdateMode, r.secondary_id, r.incid, r.secondary, r.secondary_group) == null;
        }

        public static bool ValidateRow(bool bulkUpdateMode, HluDataSet.incid_secondaryRow r,
            IEnumerable<SecondaryHabitat> secondaryHabitatList)
        {
            return ValidateRow(bulkUpdateMode, r.secondary_id, r.incid, r.secondary, r.secondary_group) == null;
        }

        private static string ValidateRow(bool _bulkUpdateMode, int secondary_id, string incid,
            string secondary_habitat, string secondary_group)
        {
            StringBuilder sbError = new StringBuilder();

            // Only validate if errors are to be shown
            if (SecondaryCodeValidation > 0)
            {
                if ((secondary_id != -1) && String.IsNullOrEmpty(incid))
                    sbError.Append(Environment.NewLine).Append("INCID is a mandatory field");

                if (String.IsNullOrEmpty(secondary_habitat))
                    sbError.Append(Environment.NewLine).Append("Secondary habitat is a mandatory field");

                if (_validSecondaryCodes == null)
                    sbError.Append(Environment.NewLine).Append("Secondary habitat is not valid without primary habitat");

                if ((_validSecondaryCodes != null) && (!_validSecondaryCodes.Contains(secondary_habitat)))
                    sbError.Append(Environment.NewLine).Append("Secondary habitat is not valid for primary habitat");

                if ((_secondaryHabitatList != null) && (_secondaryHabitatList.Count(b => b.secondary_habitat == secondary_habitat) > 1))
                    sbError.Append(Environment.NewLine).Append("Duplicate secondary habitat");
            }

            return sbError.Length > 0 ? sbError.Remove(0, 1).ToString() : null;
        }

        private bool ValidateRow()
        {
            _error = ValidateRow(_bulkUpdateMode, secondary_id, incid, secondary_habitat, secondary_group);
            return _error == null;
            //return true;
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
                            return "Error: INCID is a mandatory field";
                        }
                        _incid_bak = _incid;
                        break;
                    case "secondary_habitat":
                        // Only validate if errors are to be shown
                        if (SecondaryCodeValidation == 1)
                        {
                            if (String.IsNullOrEmpty(secondary_habitat))
                            {
                                return "Error: Secondary habitat is a mandatory field";
                            }
                            else if (_validSecondaryCodes == null)
                            {
                                return "Warning: Secondary habitat is not valid without primary habitat";
                            }
                            else if ((_validSecondaryCodes != null) && (!_validSecondaryCodes.Contains(secondary_habitat)))
                            {
                                return "Warning: Secondary habitat is not valid for primary habitat";
                            }
                            else if ((_secondaryHabitatList != null) && (_secondaryHabitatList.Count(b => b.secondary_habitat == secondary_habitat) > 1))
                            {
                                return "Error: Duplicate secondary habitat";
                            }

                            _secondary_habitat_bak = _secondary_habitat;
                        }
                        break;
                }

                return error;
            }
        }

        #endregion
    }
}
