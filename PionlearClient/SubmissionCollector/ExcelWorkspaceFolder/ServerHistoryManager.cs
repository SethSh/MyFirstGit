using System;
using System.Windows.Forms;
using SubmissionCollector.Enums;
using SubmissionCollector.Models.Package;
using SubmissionCollector.View;
using SubmissionCollector.View.Enums;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class ServerHistoryManager
    {
        public static void Show(IPackage package, IWorkbookLogger logger)
        {
            try
            {
                var historyViewModel = new HistoryViewModel(package);
                var historyDisplayer = new HistoryDisplayer(historyViewModel);
                var historyForm = new HistoryForm(historyDisplayer)
                {
                    Text = @"Server History",
                    Height = (int)FormSizeHeight.Medium,
                    Width = (int)FormSizeWidth.Large,
                    StartPosition = FormStartPosition.CenterScreen
                };
                historyForm.ShowDialog();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Show server history failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }
    }
}
