using System;
using System.Windows.Forms;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.View;
using SubmissionCollector.View.Enums;
using SubmissionCollector.View.Forms;
using SubmissionCollector.ViewModel;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class WorkersCompClassCodeReferenceManager
    {
        public static void Render(IWorkbookLogger logger)
        {
            try
            {
                var classCodeViewModel = new WorkersCompClassCodeViewModel();
                var classCodeWizard = new WorkerCompClassCodeReferenceDisplayer(classCodeViewModel);

                var form = new WorkersCompClassCodeForm(classCodeWizard)
                {
                    Text = BexConstants.ApplicationName,
                    Height = (int)FormSizeHeight.ExtraLarge,
                    Width = (int)FormSizeWidth.ExtraLarge,
                    StartPosition = FormStartPosition.CenterScreen
                };
                form.ShowDialog();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Workers Comp Class Code rendering failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }
    }
}
