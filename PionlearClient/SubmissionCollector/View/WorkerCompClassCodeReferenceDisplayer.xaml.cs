using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using PionlearClient.BexReferenceData;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for WorkerCompClassCodeDisplayer.xaml
    /// </summary>
    public partial class WorkerCompClassCodeReferenceDisplayer
    {
        private readonly WorkersCompClassCodeViewModel _viewModel;
        internal WorkerCompClassCodeReferenceDisplayer(WorkersCompClassCodeViewModel viewModel)
        {
            InitializeComponent();

            DataContext = viewModel;
            _viewModel = viewModel;
        }

        private void StateComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _viewModel.WorkersCompClassCodeView = _viewModel.WorkersCompClassCodeViews.Single(vw => vw.State.Abbreviation == _viewModel.StateAbbreviationSelected);
        }

        private void ExportButton_OnClick(object sender, RoutedEventArgs e)
        {
            var filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.LogFileName);
            
            var sb = new StringBuilder();
            sb.AppendLine(_viewModel.StateAbbreviationSelected);

            const int length = 12;
            var classCodeString = "Class Code".PadRight(length);
            var stateString = "State".PadRight(length);

            var header = $"{classCodeString}" +
                         $"\t{stateString}" +
                         "\tDescription";
            sb.AppendLine(header);
            sb.AppendLine(string.Empty);

            foreach (var item in _viewModel.WorkersCompClassCodeView.ClassCodeModels)
            {
                var line = $"{item.StateClassCodeAsString.PadRight(length)}" +
                           $"\t{item.HazardGroupName.PadRight(length)}" +
                           $"\t{item.StateDescription}";
                sb.AppendLine(line);
            }

            File.WriteAllText(filename, sb.ToString());
            Process.Start(filename);
        }
    }
}
