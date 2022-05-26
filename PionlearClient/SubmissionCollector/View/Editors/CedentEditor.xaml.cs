using System.Windows;
using System.Windows.Data;
using SubmissionCollector.ExcelWorkspaceFolder;
using Xceed.Wpf.Toolkit.PropertyGrid.Editors;

namespace SubmissionCollector.View.Editors
{
    /// <summary>
    /// Interaction logic for CedentEditor.xaml
    /// </summary>
    public partial class CedentEditor : ITypeEditor
    {
        public CedentEditor()
        {
            InitializeComponent();
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            var manager = new CedentSelectorManager();
            manager.GetCedent(new StackTraceLogger());
        }

        public static readonly DependencyProperty ValueProperty =
            DependencyProperty.Register("Value", typeof(string), typeof(CedentEditor),
                new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));

        public string Value
        {
            get => (string)GetValue(ValueProperty);
            set => SetValue(ValueProperty, value);
        }

        public FrameworkElement ResolveEditor(Xceed.Wpf.Toolkit.PropertyGrid.PropertyItem propertyItem)
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