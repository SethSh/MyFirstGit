using System;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PionlearClient;
using PionlearClient.Model;
using SubmissionCollector.BexCommunication;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.View;
using SubmissionCollector.View.Enums;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class SynchronizationManager
    {
        private SynchronizationViewModel _viewModel;
        private const string NotApplicable = "NA";
        private const int NameLength = 65;
        private const int TimeStampLength = 28;
        private const int OutOfSyncLength = 10;
        private const int IdLength = 10;

        public bool RenderHeadless()
        {
            var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
            
            _viewModel = new SynchronizationViewModel(package, new BexCommunicationManager());
            _viewModel.CreateViews();

            if (!string.IsNullOrEmpty(_viewModel.ValidationMessage) || !string.IsNullOrEmpty(_viewModel.ErrorMessage)) return true;

            return _viewModel.IsEntirePackageSynchronized;
        }

        public void RenderInBackground(IWorkbookLogger logger)
        {
            try
            {
                var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;

                const int height = (int)FormSizeHeight.ExtraLarge;
                const int width = (int)FormSizeWidth.Medium;
                _viewModel = new SynchronizationViewModel(package, new BexCommunicationManager());
                var wizard = new SynchronizationPreviewer(_viewModel);
                var uf = new SynchronizationPreviewerForm(wizard)
                {
                    Text = BexConstants.ApplicationName,
                    Height = height,
                    Width = width,
                    StartPosition = FormStartPosition.CenterScreen,
                    ControlBox = false
                };

                var backgroundWorker = new BackgroundWorker { WorkerReportsProgress = true };
                backgroundWorker.DoWork += CompareTwoSources;

                backgroundWorker.RunWorkerCompleted += (sender, e) =>
                {
                    uf.ControlBox = true;
                    if (!string.IsNullOrEmpty(_viewModel.ValidationMessage) || !string.IsNullOrEmpty(_viewModel.ErrorMessage)) return;

                    _viewModel.AddPackage();
                };

                backgroundWorker.RunWorkerAsync();
                uf.ShowDialog();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
            }
        }

        public void ShowIndicators()
        {
            var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;

            var sb = new StringBuilder();
            sb.AppendLine(WriteLine(string.Empty, string.Empty, "Source", "Out"));
            sb.AppendLine(WriteLine("Name",  "Timestamp", "ID", "of Sync"));
            sb.AppendLine(WriteLine("----",  "---------", "--", "-------"));
            sb.AppendLine();

            sb.AppendLine(WriteLine(package));
            sb.AppendLine();

            foreach (var segment in package.Segments.OrderBy(seg => seg.DisplayOrder))
            {
                sb.AppendLine(WriteLine(segment));
                sb.AppendLine();

                foreach (var excelComponent in segment.ExcelComponents.Where(ec => ec.CommonExcelMatrix is SingleOccurrenceProfileExcelMatrix)
                    .OrderBy(ec => ec.InterDisplayOrder))
                {
                    sb.AppendLine(WriteLine(excelComponent.CommonExcelMatrix.FullName, excelComponent));
                }

                sb.AppendLine();

                foreach (var excelComponent in segment.ExcelComponents.Where(ec => ec.CommonExcelMatrix is MultipleOccurrenceSegmentExcelMatrix)
                    .OrderBy(ec => ec.InterDisplayOrder).ThenBy(ec => ec.IntraDisplayOrder))
                {
                    sb.AppendLine(WriteLine(excelComponent.CommonExcelMatrix.FullName, excelComponent));
                }
                
                sb.AppendLine();
            }

            var title = $"{BexConstants.ApplicationName} - Synchronization Indicators";
            
            var isDirtyFlagDisplay = new IsDirtyFlagDisplay(sb.ToString());
            var form = new IsDirtyFlagsForm(isDirtyFlagDisplay)
            {
                Height = (int)FormSizeHeight.Large,
                Width = (int)FormSizeWidth.Large,
                Text = title,
                StartPosition = FormStartPosition.CenterScreen
            };
            form.ShowDialog();
        }
        
        private void CompareTwoSources(object sender, DoWorkEventArgs e)
        {
            _viewModel.CreateViews();
        }
        
        private static string WriteLine(string nameHeading,  string sourceTimestampHeading, string sourceIdHeading, string isDirtyHeading)
        {
            return $"{nameHeading.PadRight(NameLength)}" +
                   $"{sourceTimestampHeading.PadRight(TimeStampLength)}" +
                   $"{sourceIdHeading.PadRight(IdLength)}" +
                   $"{isDirtyHeading.PadRight(OutOfSyncLength)}";
        }

        private static string WriteLine(string name, IModel model)
        {
            var line = $"{name.PadRight(NameLength)}";

            var onServer = model.SourceId.HasValue && model.SourceTimestamp.HasValue;
            var dateString = onServer ? model.SourceTimestamp.Value.ToString("G") : NotApplicable;
            var idString = onServer ? model.SourceId.Value.ToString() : NotApplicable;

            line += dateString.PadRight(TimeStampLength);
            line += idString.PadRight(IdLength);
            
            line += $"{model.IsDirty.ToString().PadRight(OutOfSyncLength)}";

            return line;
        }

        private static string WriteLine(IModel model)
        {
            return WriteLine(model.Name, model);
        }
    }
}
