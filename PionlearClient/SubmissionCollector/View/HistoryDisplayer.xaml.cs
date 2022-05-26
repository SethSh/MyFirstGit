using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using PionlearClient;
using PionlearClient.BexReferenceData;
using SubmissionCollector.Models.Package;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.View
{
    /// <summary>
    /// Interaction logic for HistoryDisplay.xaml
    /// </summary>
    public partial class HistoryDisplayer
    {
        private readonly HistoryViewModel _viewModel;

        public HistoryDisplayer(HistoryViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = viewModel;
        }

        private void Close_OnClick(object sender, RoutedEventArgs e)
        {
            //do nothing
        }
        
        private void ExportButton_OnClick(object sender, RoutedEventArgs e)
        {
            var items = _viewModel.Items.OrderByDescending(comm => comm.Timestamp).ToList();

            var sb = new StringBuilder();
            sb.AppendLine($"History of loading this workbook content to the {BexConstants.ServerDatabaseName.ToLower()}");
            sb.AppendLine();
            
            if (items.Count > 0)
            {
                sb.AppendLine("User Name".PadRight(BexCommunicationEntry.Padding)
                              + "Timestamp".PadRight(BexCommunicationEntry.Padding)
                              + "Activity".PadRight(BexCommunicationEntry.Padding));
                sb.AppendLine("---- ----".PadRight(BexCommunicationEntry.Padding)
                              + "---------".PadRight(BexCommunicationEntry.Padding)
                              + "--------".PadRight(BexCommunicationEntry.Padding));
                
                items.ForEach(item =>
                {
                    var singleRow = item.UserName.PadRight(BexCommunicationEntry.Padding)
                                  + item.Timestamp.ToString(CultureInfo.CurrentCulture).PadRight(BexCommunicationEntry.Padding)
                                  + item.Activity?.PadRight(BexCommunicationEntry.Padding);
                    sb.AppendLine(singleRow);
                });
            }
            else
            {
                sb.AppendLine("No history entries in the log");
            }

            var filename = Path.Combine(ConfigurationHelper.AppDataFolder, BexFileNames.LogFileName);
            File.WriteAllText(filename, sb.ToString());
            Process.Start(filename);
        }
    }
}
