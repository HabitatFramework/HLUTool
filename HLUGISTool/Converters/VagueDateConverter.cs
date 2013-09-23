using System;
using System.Globalization;
using System.Windows.Data;
using HLU.Date;

namespace HLU.Converters
{
    class VagueDateConverter : IValueConverter
    {
        #region IValueConverter Members

        /// <summary>
        /// from DB to UI : VagueDateInstance -> string
        /// </summary>
        /// <param name="value"></param>
        /// <param name="targetType"></param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            try
            {
                Date.VagueDateInstance vd = value as Date.VagueDateInstance;
                if (vd == null)
                    return null;
                else if (vd.IsBad)
                    return vd.UserEntry;
                else if (vd.IsUnknown)
                    return VagueDate.VagueDateTypes.Unknown.ToString();
                else
                    return Date.VagueDate.FromVagueDateInstance(vd, VagueDate.DateType.Vague);
            }
            catch { }

            return value;
        }

        /// <summary>
        /// from UI to DB : string -> VagueDateInstance
        /// </summary>
        /// <param name="value"></param>
        /// <param name="targetType"></param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            try
            {
                string vagueDateString = value as string;
                if (vagueDateString == null) return value;

                string formattedDateString;
                string dateType = VagueDate.GetType(vagueDateString, out formattedDateString);
                int startDate = Date.VagueDate.ToTimeSpanDays(formattedDateString, dateType, VagueDate.DateType.Start);
                int endDate = Date.VagueDate.ToTimeSpanDays(formattedDateString, dateType, VagueDate.DateType.End);

                if ((startDate < endDate) || (endDate == VagueDate.DateUnknown))
                    return new Date.VagueDateInstance(startDate, endDate, dateType, vagueDateString);
                else
                    return new Date.VagueDateInstance(startDate, endDate, dateType, vagueDateString);
            }
            catch { }
            
            return value;
        }

        #endregion
    }
}
