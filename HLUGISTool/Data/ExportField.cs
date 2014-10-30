// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2014 Thames Valley Environmental Records Centre
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
using System.Linq;
using System.Text;

namespace HLU.Data
{
    public class ExportField
    {
        #region Fields

        private int _fieldOrdinal;
        private string _tableName;
        private string _columnName;
        private string _fieldName;
        private Type _fieldType;
        private int _fieldLength;
        private int _fieldsCount;
        private string _fieldFormat;
        private bool _autoNum;
        private int _fieldOrder;

        #endregion

        #region Constructor

        public ExportField()
        {
        }

        public ExportField(int fieldOrdinal, string tableName, string columnName, string fieldName, Type fieldType,
            int fieldLength, int fieldsCount, string fieldFormat, bool autoNum, int fieldOrder)
        {
            _fieldOrdinal = fieldOrdinal;
            _tableName = tableName;
            _columnName = columnName;
            _fieldName = fieldName;
            _fieldType = fieldType;
            _fieldLength = fieldLength;
            _fieldsCount = fieldsCount;
            _fieldFormat = fieldFormat;
            _autoNum = autoNum;
            _fieldOrder = fieldOrder;
        }

        #endregion // Constructor

        #region Properties

        public int FieldOrdinal
        {
            get { return _fieldOrdinal; }
            set { _fieldOrdinal = value; }
        }

        public string TableName
        {
            get { return _tableName; }
            set { _tableName = value; }
        }

        public string ColumnName
        {
            get { return _columnName; }
            set { _columnName = value; }
        }

        public string FieldName
        {
            get { return _fieldName; }
            set { _fieldName = value; }
        }

        public Type FieldType
        {
            get { return _fieldType; }
            set { _fieldType = value; }
        }

        public int FieldLength
        {
            get { return _fieldLength; }
            set { _fieldLength = value; }
        }

        public int FieldsCount
        {
            get { return _fieldsCount; }
            set { _fieldsCount = value; }
        }

        public string FieldFormat
        {
            get { return _fieldFormat; }
            set { _fieldFormat = value; }
        }

        public bool AutoNum
        {
            get { return _autoNum; }
            set { _autoNum = value; }
        }

        public int FieldOrder
        {
            get { return _fieldOrder; }
            set { _fieldOrder = value; }
        }

        #endregion // Properties

        #region Methods

        public override string ToString()
        {
            return _fieldName;
        }

        public override int GetHashCode()
        {
            return (this.TableName.GetHashCode() * 10000)
                + (this.ColumnName.GetHashCode() * 1000)
                + (this.FieldName.GetHashCode() * 100)
                + (this.FieldType.GetHashCode() * 10)
                + (this.FieldLength.GetHashCode() * 1);
        }

        public virtual bool Equals(ExportField other)
        {
            if (other == null) return false;

            return ((this._tableName == other._tableName)
                && (this._columnName == other._columnName)
                && (this._fieldName == other._fieldName));
        }

        public override bool Equals(object obj)
        {
            if (this.GetType() != obj.GetType()) return false;

            return Equals(obj as ExportField);
        }

        #endregion // Methods

    }
}