using System;
using System.Globalization;
using System.Windows.Data;

namespace BulkEditor.UI.Converters
{
    /// <summary>
    /// Converter that converts string values to boolean based on a parameter
    /// </summary>
    public class StringToBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string stringValue && parameter is string parameterValue)
            {
                return string.Equals(stringValue, parameterValue, StringComparison.OrdinalIgnoreCase);
            }
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue && parameter is string parameterValue)
            {
                return boolValue ? parameterValue : null;
            }
            return null;
        }
    }
}