using System.Linq;
using System.Windows;
using System.Windows.Controls;
using PionlearClient;
using SubmissionCollector.View.Enums;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for UmbrellaSelectorWizard.xaml
    /// </summary>
    public partial class UmbrellaTypeAllocator
    {
        private readonly UmbrellaTypeAllocatorViewModel _viewModel;
        public FormResponse Response { get; set; }

        public UmbrellaTypeAllocator(UmbrellaTypeAllocatorViewModel viewModel)
        {
            InitializeComponent();

            DataContext = viewModel;
            _viewModel = viewModel;
            Response = FormResponse.Cancel;
        }

        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            Response = FormResponse.Cancel;
        }

        private void OkButton_OnClick(object sender, RoutedEventArgs e)
        {
            Response = FormResponse.Ok;
        }

        private void UmbrellaTypeListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
            
        }

        private void ToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            RedrawOkButton();
        }

        private void ToggleButton_OnUnchecked(object sender, RoutedEventArgs e)
        {
            RedrawOkButton();
        }

        private void RedrawOkButton()
        {
            if (_viewModel.UmbrellaItems.Any(item => item.IsSelected))
            {
                _viewModel.OkButtonEnabled = true;
                _viewModel.OkButtonToolTip = null;
            }
            else
            {
                _viewModel.OkButtonEnabled = false;
                _viewModel.OkButtonToolTip = $"At least one {BexConstants.UmbrellaTypeName.ToLower()} must be selected";
            }
        }
    }
}
