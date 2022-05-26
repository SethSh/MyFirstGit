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
    internal class CedentSelectorManager
    {
        public void GetCedent(IWorkbookLogger logger)
        {
            try
            {
                GetCedent();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Get Cedent failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        private static void GetCedent()
        {
            var viewModel = new CedentSelectorViewModel();
            var cedentSelector = new CedentSelector(viewModel);
            var form = new CedentSelectorForm(cedentSelector)
            {
                Text = BexConstants.ApplicationName,
                Height = (int) FormSizeHeight.Large,
                Width = (int) FormSizeWidth.Medium,
                StartPosition = FormStartPosition.CenterScreen
            };
            form.ShowDialog();
        }
    }
}