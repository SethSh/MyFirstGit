using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using PionlearClient.BexReferenceData;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for WorkerCompClassCodeReferenceQueryDisplayer.xaml
    /// </summary>
    public partial class WorkerCompClassCodeReferenceSearchDisplayer
    {
        private readonly WorkersCompClassCodeSearchViewModel _viewModel;

        internal WorkerCompClassCodeReferenceSearchDisplayer(WorkersCompClassCodeSearchViewModel viewModel)
        {
            InitializeComponent();

            DataContext = viewModel;
            _viewModel = viewModel;

            CriteriaTextBox.Text = string.Empty;
            CriteriaTextBox.Focus();
        }



        private void ExportButton_OnClick(object sender, RoutedEventArgs e)
        {
            var filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.LogFileName);

            var sb = new StringBuilder();
            sb.AppendLine(_viewModel.SearchCriteria);

            const int length = 12;
            var stateString = "State".PadRight(length);
            var classCodeString = "Class Code".PadRight(length);
            var hazardGroupString = "Hazard".PadRight(length);
            var descriptionString = "Description";
            var header = $"{stateString}" +
                         $"\t{classCodeString}" +
                         $"\t{hazardGroupString}" +
                         $"\t{descriptionString}";

            sb.AppendLine(header);
            sb.AppendLine(string.Empty);

            foreach (var item in _viewModel.FilteredClassCodeViewItems)
            {
                var line = $"{item.StateAbbreviation.PadRight(length)}" +
                           $"\t{item.StateClassCodeAsString.PadRight(length)}" +
                           $"\t{item.HazardGroupName.PadRight(length)}" +
                           $"\t{item.StateDescription}";
                sb.AppendLine(line);
            }

            File.WriteAllText(filename, sb.ToString());
            Process.Start(filename);
        }

    }
}