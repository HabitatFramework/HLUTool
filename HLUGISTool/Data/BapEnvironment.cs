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
using System.ComponentModel;
using System.Linq;
using System.Text;
using HLU.Data.Model;
using HLU.Properties;

namespace HLU.Data
{
    public class BapEnvironment : IDataErrorInfo
    {
        #region Fields

        private int _bap_id;
        private string _incid;
        private bool _bulkUpdateMode;
        private bool _secondaryPriorityHabitat;
        private string _bap_habitat;
        private string _quality_determination;
        private string _quality_interpretation;
        private string _interpretation_comments;
        private string _incid_bak;
        private string _bap_habitat_bak;
        private string _quality_determination_bak;
        private string _quality_interpretation_bak;
        string _error;
        private static IEnumerable<BapEnvironment> _bapEnvironmentList;
        public readonly static string BAPDetQltyUserAdded = Settings.Default.BAPDeterminationQualiltyUserAdded;
        public readonly static string BAPHabitatIgnore = Settings.Default.BAPHabitatIgnore;

        #endregion

        #region ctor

        public BapEnvironment()
        {
            _bulkUpdateMode = false;
            _secondaryPriorityHabitat = true; // new rows default to secondary as that is what UI needs to create
            _bap_id = -1; // arbitrary PK for a new row
        }

        public BapEnvironment(bool bulkUpdateMode, bool isSecondary)
        {
            _bulkUpdateMode = bulkUpdateMode;
            _secondaryPriorityHabitat = isSecondary;
            _bap_id = -1; // arbitrary PK for a new row
        }

        public BapEnvironment(bool bulkUpdateMode, bool isSecondary, HluDataSet.incid_bapRow dataRow)
        {
            _bulkUpdateMode = bulkUpdateMode;
            _secondaryPriorityHabitat = isSecondary;
            HluDataSet.incid_bapDataTable table = (HluDataSet.incid_bapDataTable)dataRow.Table;
            _bap_id = dataRow.bap_id;
            _incid = dataRow.incid;
            _bap_habitat = dataRow.IsNull(table.bap_habitatColumn) ? null : dataRow.bap_habitat;
            _quality_determination = dataRow.IsNull(table.quality_determinationColumn) ? null : dataRow.quality_determination;
            _quality_interpretation = dataRow.IsNull(table.quality_interpretationColumn) ? null : dataRow.quality_interpretation;
            //---------------------------------------------------------------------
            // CHANGED: CR2 (Apply button)
            // Update the _interpretation_comments string directly, rather than via the property,
            // so that the Changed flag is not set.
            //
            //this.interpretation_comments = dataRow.IsNull(table.interpretation_commentsColumn) ?
            //    null : dataRow.interpretation_comments;
            if (dataRow.IsNull(table.interpretation_commentsColumn))
                _interpretation_comments = null;
            else
                _interpretation_comments = dataRow.interpretation_comments.Length < 255 ? dataRow.interpretation_comments : dataRow.interpretation_comments.Substring(0, 254);
            //---------------------------------------------------------------------
        }

        public BapEnvironment(bool bulkUpdateMode, bool isSecondary, HluDataSet.incid_bapRow dataRow, IEnumerable<BapEnvironment> beList)
        {
            _bulkUpdateMode = bulkUpdateMode;
            _secondaryPriorityHabitat = isSecondary;
            HluDataSet.incid_bapDataTable table = (HluDataSet.incid_bapDataTable)dataRow.Table;
            _bap_id = dataRow.bap_id;
            _incid = dataRow.incid;
            _bap_habitat = dataRow.IsNull(table.bap_habitatColumn) ? null : dataRow.bap_habitat;
            _quality_determination = dataRow.IsNull(table.quality_determinationColumn) ? null : dataRow.quality_determination;
            _quality_interpretation = dataRow.IsNull(table.quality_interpretationColumn) ? null : dataRow.quality_interpretation;
            //---------------------------------------------------------------------
            // CHANGED: CR2 (Apply button)
            // Update the _interpretation_comments string directly, rather than via the property,
            // so that the Changed flag is not set.
            //
            //this.interpretation_comments = dataRow.IsNull(table.interpretation_commentsColumn) ?
            //    null : dataRow.interpretation_comments;
            if (dataRow.IsNull(table.interpretation_commentsColumn))
                _interpretation_comments = null;
            else
                _interpretation_comments = dataRow.interpretation_comments.Length < 255 ? dataRow.interpretation_comments : dataRow.interpretation_comments.Substring(0, 254);
            //---------------------------------------------------------------------
            _bapEnvironmentList = beList;
        }

        public BapEnvironment(bool bulkUpdateMode, bool isSecondary, object[] itemArray)
        {
            _bulkUpdateMode = bulkUpdateMode;
            _secondaryPriorityHabitat = isSecondary;
            Int32.TryParse(itemArray[0].ToString(), out _bap_id);
            _incid = itemArray[1].ToString();
            _bap_habitat = itemArray[2].ToString();
            _quality_determination = itemArray[3].ToString();
            _quality_interpretation = itemArray[4].ToString();
            //---------------------------------------------------------------------
            // CHANGED: CR2 (Apply button)
            // Update the _interpretation_comments string directly, rather than via the property,
            // so that the Changed flag is not set.
            //
            //this.interpretation_comments = itemArray[5].ToString();
            if (itemArray[5].ToString() == null)
                _interpretation_comments = null;
            else
                _interpretation_comments = itemArray[5].ToString().Length < 255 ? itemArray[5].ToString() : itemArray[5].ToString().Substring(0, 254);
            //---------------------------------------------------------------------
        }

        public BapEnvironment(bool bulkUpdateMode, bool isSecondary, int bap_id, string incid, string bap_habitat, 
            string quality_determination, string quality_interpretation, string interpretation_comments)
        {
            _bulkUpdateMode = bulkUpdateMode;
            _secondaryPriorityHabitat = isSecondary;
            _bap_id = bap_id;
            _incid = incid;
            _bap_habitat = bap_habitat;
            _quality_determination = quality_determination;
            _quality_interpretation = quality_interpretation;
            //---------------------------------------------------------------------
            // CHANGED: CR2 (Apply button)
            // Update the _interpretation_comments string directly, rather than via the property,
            // so that the Changed flag is not set.
            //
            //this.interpretation_comments = interpretation_comments;
            if (interpretation_comments == null)
                _interpretation_comments = null;
            else
                _interpretation_comments = interpretation_comments.Length < 255 ? interpretation_comments : interpretation_comments.Substring(0, 254);
            //---------------------------------------------------------------------
        }

        #endregion

        #region DataChanged

        //---------------------------------------------------------------------
        // CHANGED: CR2 (Apply button)
        // Create a handler so that updates to the BAP records can be picked
        // up back in the main window.
        //
        // declare the delegate since using the generic pattern
        public delegate void DataChangedEventHandler(bool Changed);

        // declare the event
        public event DataChangedEventHandler DataChanged;
        //---------------------------------------------------------------------

        #endregion

        #region Properties

        public static IEnumerable<BapEnvironment> BapEnvironmentList
        {
            set { _bapEnvironmentList = value; }
        }

        public bool BulkUpdateMode
        {
            get { return _bulkUpdateMode; }
            set { _bulkUpdateMode = value; }
        }

        public bool IsAdded
        {
            get { return _bap_id == -1; }
        }

        #region incid_bapRow

        public int bap_id
        {
            get { return _bap_id; }
            set { _bap_id = value; }
        }

        public string incid
        {
            get { return _incid; }
            set { _incid = value; }
        }

        public string bap_habitat
        {
            get { return _bap_habitat; }
            set 
            {
                _bap_habitat = value;
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                if (this.DataChanged != null)
                    this.DataChanged(true);
                //---------------------------------------------------------------------
            }
        }

        public string quality_determination
        {
            get { return _quality_determination; }
            set 
            {
                _quality_determination = value;
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                if (this.DataChanged != null)
                    this.DataChanged(true);
                //---------------------------------------------------------------------
            }
        }

        public string quality_interpretation
        {
            get { return _quality_interpretation; }
            set 
            { 
                _quality_interpretation = value;
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                if (this.DataChanged != null)
                    this.DataChanged(true);
                //---------------------------------------------------------------------
            }
        }

        public string interpretation_comments
        {
            get { return _interpretation_comments; }
            set 
            { 
                _interpretation_comments = value == null || value.Length < 255 ? value : value.Substring(0, 254);
                //---------------------------------------------------------------------
                // CHANGED: CR2 (Apply button)
                // Flag that the current record has changed so that the apply button
                // will appear.
                if (this.DataChanged != null)
                    this.DataChanged(true);
                //---------------------------------------------------------------------
            }
        }
        
        #endregion

        #endregion

        #region Public methods

        public object[] ToItemArray()
        {
            return new object[] { _bap_id, _incid, _bap_habitat, _quality_determination, 
                _quality_interpretation, _interpretation_comments };
        }

        public object[] ToItemArray(int bapID, string incid)
        {
            return new object[] { bapID, incid, _bap_habitat, _quality_determination, 
                _quality_interpretation, _interpretation_comments };
        }

        public object[] ToItemArray(int bapID, string incid, bool isSecondary)
        {
            if (isSecondary) MakeSecondary();
            return new object[] { bapID, incid, _bap_habitat, _quality_determination, 
                _quality_interpretation, _interpretation_comments };
        }

        public void MakeSecondary()
        {
            _secondaryPriorityHabitat = true;
        }

        public static HluDataSet.incid_bapRow MakeSecondary(HluDataSet.incid_bapRow r)
        {
            return r;
        }

        #endregion

        #region Validation

        public bool IsDuplicate
        {
            get
            {
                return _bapEnvironmentList != null && _bapEnvironmentList.Count(be => be.bap_habitat == this.bap_habitat) > 0;
            }
        }

        public bool IsValid()
        {
            return ValidateRow();
        }

        public bool IsValid(bool bulkUpdateMode, bool isSecondary)
        {
            return String.IsNullOrEmpty(ValidateRow(bulkUpdateMode, isSecondary, _bap_id, 
                _incid, _bap_habitat, _quality_determination, _quality_interpretation));
        }

        public bool IsValid(bool bulkUpdateMode, bool isSecondary, HluDataSet.incid_bapRow r)
        {
            return String.IsNullOrEmpty(ValidateRow(bulkUpdateMode, isSecondary, r.bap_id,
                r.incid, r.bap_habitat, r.quality_determination, r.quality_interpretation));
        }

        public string ErrorMessages { get { return _error.ToString(); } }

        public static bool ValidateRow(bool bulkUpdateMode, bool isSecondary, HluDataSet.incid_bapRow r)
        {
            return ValidateRow(bulkUpdateMode, isSecondary, r.bap_id, r.incid, r.bap_habitat, 
                r.quality_determination, r.quality_interpretation) == null;
        }

        public static bool ValidateRow(bool bulkUpdateMode, bool isSecondary, HluDataSet.incid_bapRow r, 
            IEnumerable<BapEnvironment> bapEnvironmentList)
        {
            return ValidateRow(bulkUpdateMode, isSecondary, r.bap_id, r.incid, r.bap_habitat, 
                r.quality_determination, r.quality_interpretation) == null;
        }

        private static string ValidateRow(bool _bulkUpdateMode, bool isSecondary, int bap_id, string incid, 
            string bap_habitat, string quality_determination, string quality_interpretation)
        {
            StringBuilder sbError = new StringBuilder();

            if ((bap_id != -1) && String.IsNullOrEmpty(incid))
                sbError.Append(Environment.NewLine).Append("INCID is a mandatory field");

            if (String.IsNullOrEmpty(bap_habitat))
                sbError.Append(Environment.NewLine).Append("Priority habitat is a mandatory field");

            //if ((_bapEnvironmentList != null) && (_bapEnvironmentList.Count(b => b.bap_habitat == bap_habitat) > 1))
            //    sbError.Append(Environment.NewLine).Append("Duplicate priority environment");

            if (String.IsNullOrEmpty(quality_determination))
            {
                if (!_bulkUpdateMode)
                    sbError.Append(Environment.NewLine).Append("Determination quality is a mandatory field");
            }
            else
            {
                //---------------------------------------------------------------------
                // FIX: 015 Allow 'None' habitats to be managed
                // Validate that if this is a secondary priority habitat (i.e. in
                // the secondary list) and the habitat is to be ignored (i.e. it equals
                // 'None') then the determination quality can be anything except
                // 'Not present but close to definition').
                if (isSecondary)
                {
                    if ((bap_habitat == BAPHabitatIgnore) && (quality_determination == BAPDetQltyUserAdded))
                    {
                        sbError.Append(Environment.NewLine)
                            .Append("Determination quality for 'None' priority habitats cannot be " +
                            "'Not present but close to definition'");
                    }
                    // Validate that if this is a secondary priority habitat (i.e. in
                    // the secondary list) and the habitat is not to be ignored (i.e. it
                    // does not equal 'None') then the determination quality can only be
                    // 'Not present but close to definition').
                    else if ((bap_habitat != BAPHabitatIgnore) && (quality_determination != BAPDetQltyUserAdded))
                    {
                        sbError.Append(Environment.NewLine)
                            .Append("Determination quality for potential priority habitats can only be " +
                            "'Not present but close to definition'");
                    }
                }
                // Validate that if this is not a secondary priority habitat (i.e. in
                // the primary list) then the determination quality can be anything except
                // 'Not present but close to definition').
                else
                {
                    if (!isSecondary && (quality_determination == BAPDetQltyUserAdded))
                    {
                        sbError.Append(Environment.NewLine)
                            .Append("Determination quality cannot be " +
                            "'Not present but close to definition' for 'primary' priority habitats");
                    }
                }
                //---------------------------------------------------------------------
            }

            if (!_bulkUpdateMode && String.IsNullOrEmpty(quality_interpretation))
                sbError.Append(Environment.NewLine).Append("Interpretation quality is a mandatory field");

            return sbError.Length > 0 ? sbError.Remove(0, 1).ToString() : null;
        }

        private bool ValidateRow()
        {
            _error = ValidateRow(_bulkUpdateMode, _secondaryPriorityHabitat, bap_id, incid, bap_habitat, quality_determination, quality_interpretation);
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
                        if ((bap_id != -1) && String.IsNullOrEmpty(incid))
                        {
                            _incid = _incid_bak;
                            return "INCID is a mandatory field";
                        }
                        _incid_bak = _incid;
                        break;
                    case "bap_habitat":
                        if (String.IsNullOrEmpty(bap_habitat))
                        {
                            //_bap_habitat = _bap_habitat_bak;
                            return "Priority habitat is a mandatory field";
                        }
                        else if ((_bapEnvironmentList != null) && (_bapEnvironmentList.Count(b => b.bap_habitat == bap_habitat) > 1))
                        {
                            return "Duplicate priority habitat";
                        }
                        _bap_habitat_bak = _bap_habitat;
                        break;
                    case "quality_determination":
                        if (String.IsNullOrEmpty(quality_determination))
                        {
                            if (!_bulkUpdateMode)
                            {
                                //_quality_determination = _quality_determination_bak;
                                return "Determination quality is a mandatory field";
                            }
                        }
                        else
                        {
                            //---------------------------------------------------------------------
                            // FIX: 015 Allow 'None' habitats to be managed
                            // Validate that if this is a secondary priority habitat (i.e. in
                            // the secondary list) and the habitat is to be ignored (i.e. it equals
                            // 'None') then the determination quality can be anything except
                            // 'Not present but close to definition').
                            if (_secondaryPriorityHabitat)
                            {
                                if ((bap_habitat == BAPHabitatIgnore) && (quality_determination == BAPDetQltyUserAdded))
                                {
                                    //_quality_determination = _quality_determination_bak;
                                    return "Determination quality for 'None' priority habitats cannot be " +
                                        "'Not present but close to definition'";
                                }
                                // Validate that if this is a secondary priority habitat (i.e. in
                                // the secondary list) and the habitat is not to be ignored (i.e. it
                                // does not equal 'None') then the determination quality can only be
                                // 'Not present but close to definition').
                                else if ((bap_habitat != BAPHabitatIgnore) && (quality_determination != BAPDetQltyUserAdded))
                                {
                                    //_quality_determination = _quality_determination_bak;
                                    return "Determination quality for potential priority habitats can only be " +
                                        "'Not present but close to definition'";
                                }
                            }
                            // Validate that if this is not a secondary priority habitat (i.e. in
                            // the primary list) then the determination quality can be anything except
                            // 'Not present but close to definition').
                            else
                            {
                                if ((quality_determination == BAPDetQltyUserAdded))
                                {
                                    //_quality_determination = _quality_determination_bak;
                                    return "Determination quality cannot be " +
                                        "'Not present but close to definition' for 'primary' priority habitats";
                                }
                            }
                            //---------------------------------------------------------------------
                        }

                        _quality_determination_bak = _quality_determination;
                        break;
                    case "quality_interpretation":
                        if (!_bulkUpdateMode && String.IsNullOrEmpty(quality_interpretation))
                        {
                            //_quality_interpretation = _quality_interpretation_bak;
                            return "Interpretation quality is a mandatory field";
                        }
                        _quality_interpretation_bak = _quality_interpretation;
                        break;
                    case "interpretation_comments":
                        break;
                }

                return error;
            }
        }

        #endregion
    }
}
