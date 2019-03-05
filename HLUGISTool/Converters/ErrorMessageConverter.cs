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
using System.Globalization;
using System.Windows.Data;

namespace HLU.Converters
{
    //---------------------------------------------------------------------
    // FIX: 080 Functionality to display warning level messages.
    //    
    /// <summary>
    /// Converter to split leading text (before first colon)
    /// from an error message to determine the error level
    /// (e.g. "Error" or "Warning") so they can be displayed
    /// differently in the interface.
    /// </summary>
    /// <seealso cref="System.Windows.Data.IValueConverter" />
    class ErrorMessageConverter : IValueConverter
    {
         #region IValueConverter Members
  
         public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
         {
             // Input as String
             var val = (string)value;

             // Attempt to split input message by colon
             string[] parts = ((string)val).Split(':');

             // Invalid input message format, return full string
             if (parts == null || parts.Length == 0)
                 return null;

             // Return input message before first colon
             if (parts.Length == 1)
                 return val;
             else
             {
                 //int colonPos = int.Parse(parts[0]);
                 //return val.Substring(0, colonPos - 1).Trim();
                 return parts[0].Trim();
             }
         }
  
         public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
         {
             return null;
         }
  
         #endregion
    }
    //---------------------------------------------------------------------
}
