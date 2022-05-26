using System.Windows;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for MessageBoxForFactors.xaml
    /// </summary>
    public partial class MessageBoxForFactors
    {
        private readonly IMessageBoxForFactorsViewModel _viewModel;

        public MessageBoxForFactors(IMessageBoxForFactorsViewModel viewModel)
        {
            InitializeComponent();

            _viewModel = viewModel;
            
            DataContext = viewModel;
            CancelButton.Focus();
        }

        private void RenameButton_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.UpdateFactorOption = UpdateFactorOption.Rename;
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.UpdateFactorOption = UpdateFactorOption.Delete;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.UpdateFactorOption = UpdateFactorOption.Cancel;
        }
    }
}
