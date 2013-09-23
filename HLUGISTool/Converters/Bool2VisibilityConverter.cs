using System;
using System.Globalization;
using System.Windows.Data;

namespace HLU.Converters
{
    class Bool2VisibilityConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool val = (bool)value;
            if (val) 
                return System.Windows.Visibility.Visible;
            else 
                return System.Windows.Visibility.Hidden;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }

        #endregion
    }
}
