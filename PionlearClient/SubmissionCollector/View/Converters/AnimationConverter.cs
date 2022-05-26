using System;
using System.Globalization;
using System.Windows.Data;
using SubmissionCollector.Properties;

namespace SubmissionCollector.View.Converters
{
    internal class AnimationConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null && (bool)value) return Resources.AnimatedEllipse;

            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
