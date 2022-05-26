using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using SubmissionCollector.Models;
using SubmissionCollector.Properties;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View.Converters
{
    public class SynchronizationIconConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null) return null;

            switch ((SynchronizationCode) value)
            {
                case SynchronizationCode.New: return Resources.New.ToBitmapSource();
                case SynchronizationCode.Deleted: return Resources.DeleteX.ToBitmapSource();
                case SynchronizationCode.InSynchronization: return Resources.CloudChecked.ToBitmapSource();
                case SynchronizationCode.NotInSynchronization: return Resources.CloudRed.ToBitmapSource();

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class SynchronizationForegroundConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null) return null;

            switch ((SynchronizationCode)value)
            {
                case SynchronizationCode.New: return new SolidColorBrush(Colors.Black);
                case SynchronizationCode.Deleted: return new SolidColorBrush(Colors.Gray); 
                case SynchronizationCode.InSynchronization: return new SolidColorBrush(Colors.Black); 
                case SynchronizationCode.NotInSynchronization: return new SolidColorBrush(Colors.Black);

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class SynchronizationStrikethroughConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null) return null;

            switch ((SynchronizationCode)value)
            {
                case SynchronizationCode.Deleted: return TextDecorations.Strikethrough;
                    
                default:
                    return null;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}