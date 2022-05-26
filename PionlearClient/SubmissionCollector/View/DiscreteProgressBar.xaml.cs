using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for DiscreteProgressBar.xaml
    /// </summary>
    public partial class DiscreteProgressBar
    {
        public DiscreteProgressBar(DiscreteProgressBarViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}
