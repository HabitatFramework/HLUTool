// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2013 Andy Foy
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

using System.ComponentModel;
using System.Globalization;
using System.Windows.Controls;
using System.Windows.Data;

namespace HLU.UI.UserControls
{
    public class CellDataInfoValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            // obtain the bound business object
            BindingExpression expression = value as BindingExpression;
            IDataErrorInfo info = expression.DataItem as IDataErrorInfo;

            // determine the binding path
            string boundProperty = expression.ParentBinding.Path.Path;

            // obtain any errors relating to this bound property
            string error = info[boundProperty];
            if (!string.IsNullOrEmpty(error))
            {
                return new ValidationResult(false, error);
            }

            return ValidationResult.ValidResult;
        }
    }
}
