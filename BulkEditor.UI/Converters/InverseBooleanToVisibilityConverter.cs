using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace BulkEditor.UI.Converters
{
    /// <summary>
    /// Converts a boolean value to a Visibility value, inverting the logic
    /// True -> Collapsed, False -> Visible
    /// </summary>
    public class InverseBooleanToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return boolValue ? Visibility.Collapsed : Visibility.Visible;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Visibility visibility)
            {
                return visibility != Visibility.Visible;
            }
            return true;
        }
    }
}