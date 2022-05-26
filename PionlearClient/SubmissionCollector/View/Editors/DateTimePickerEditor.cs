using System.Globalization;
using System.Windows;
using System.Windows.Data;
using Xceed.Wpf.Toolkit;
using Xceed.Wpf.Toolkit.PropertyGrid;
using Xceed.Wpf.Toolkit.PropertyGrid.Editors;

namespace SubmissionCollector.View.Editors
{
    public class DateTimePickerEditor : DateTimePicker, ITypeEditor
    {
        public DateTimePickerEditor()
        {
            Format = DateTimeFormat.Custom;
            FormatString = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
            
            TimePickerVisibility = Visibility.Collapsed;
            ShowButtonSpinner = false;
            AutoCloseCalendar = true;
            ShowDropDownButton = true;
            HorizontalAlignment = HorizontalAlignment.Stretch;
            HorizontalContentAlignment = HorizontalAlignment.Left;
        }

        public FrameworkElement ResolveEditor(PropertyItem propertyItem)
        {
            var binding = new Binding("Value")
            {
                Source = propertyItem,
                Mode = propertyItem.IsReadOnly ? BindingMode.OneWay : BindingMode.TwoWay
            };

            BindingOperations.SetBinding(this, ValueProperty, binding);
            return this;
        }
    }
}
