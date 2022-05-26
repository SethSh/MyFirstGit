using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace SubmissionCollector.View.Converters
{
    public class FontBoldConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null && (bool)value)
            {
                {
                    return FontWeights.Heavy;
                }
            }
            return FontWeights.Normal;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
