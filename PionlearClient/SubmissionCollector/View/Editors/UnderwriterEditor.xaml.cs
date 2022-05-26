using System.Windows;
using System.Windows.Data;
using SubmissionCollector.ExcelWorkspaceFolder;
using Xceed.Wpf.Toolkit.PropertyGrid.Editors;

namespace SubmissionCollector.View.Editors
{
    /// <summary>
    /// Interaction logic for UnderwriterEditor.xaml
    /// </summary>
    public partial class UnderwriterEditor :  ITypeEditor
    {
        public UnderwriterEditor()
        {
            InitializeComponent();
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            var manager = new UnderwriterSelectorManager();
            manager.GetUnderwriter();
        }

        public static readonly DependencyProperty ValueProperty =
            DependencyProperty.Register("Value", typeof(string), typeof(UnderwriterEditor),
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
