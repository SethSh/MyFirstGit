using System.Windows;
using System.Windows.Controls;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for SynchronizationSelector.xaml
    /// </summary>
    public partial class SynchronizationPreviewer
    {
        private readonly BaseSynchronizationViewModel _viewModel;

        public SynchronizationPreviewer(BaseSynchronizationViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = _viewModel;
            OkButton.Focus();

            FontSizeSlider.Value = _viewModel.DefaultTreeFontSize;
        }

        private void OkButton_OnClick(object sender, RoutedEventArgs e)
        {
            var isDirty = !_viewModel.IsEntirePackageSynchronized;
            Globals.Ribbons.SubmissionRibbon.SetSynchronizationButtonImage(isDirty);
        }
        
        private void FontSizeSlider_OnValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            var slider = (Slider) sender;
            var value = slider.Value;
            if (_viewModel != null) _viewModel.TreeFontSize = value;
        }
    }
}
