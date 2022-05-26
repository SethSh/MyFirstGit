using System;
using System.Text;
using System.Windows.Forms;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.View;
using SubmissionCollector.View.Enums;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class GetStartedDocumentationManager
    {
        public static void GetDocumentation(IWorkbookLogger logger)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine($"Minimum requirements to upload to the {BexConstants.ServerDatabaseName.ToLower()}");
                sb.AppendLine();
                sb.AppendLine($"{BexConstants.HalfTab}Rename {BexConstants.PackageName}");
                sb.AppendLine($"{BexConstants.HalfTab}Select {BexConstants.AnalystName}");
                sb.AppendLine($"{BexConstants.HalfTab}Enter {BexConstants.UnderwritingYearName}");
                sb.AppendLine($"{BexConstants.HalfTab}Save workbook");

                Show(sb.ToString());
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Get started documentation failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        private static void Show(string documentation)
        {
            var viewModel = new MessageBoxControlViewModel {Message = documentation, MessageType = MessageType.Documentation};
            var messageBoxControl = new MessageBoxControl(viewModel);
            var form = new MessageForm(messageBoxControl)
            {
                Height = (int)FormSizeHeight.Medium,
                Width = (int)FormSizeWidth.Small,
                Text = BexConstants.ApplicationName,
                StartPosition = FormStartPosition.CenterScreen,
            };
            form.ShowDialog();
        }
    }
}
