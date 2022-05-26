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
    internal class UnderwriterSelectorManager
    {
        public void GetUnderwriter()
        {
            try
            {
                var underwriterSelectorViewModel = new UnderwriterSelectorViewModel();
                var underwriterSelector = new UnderwriterSelector(underwriterSelectorViewModel);
                var form = new UnderwriterSelectorForm(underwriterSelector)
                {
                    Text = BexConstants.ApplicationName,
                    Height = (int)FormSizeHeight.Large,
                    Width = (int)FormSizeWidth.Medium,
                    StartPosition = FormStartPosition.CenterScreen
                };
                form.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageHelper.Show($"Underwriter selector failed: {ex.Message}", MessageType.Stop);
            }
        }
    }
}
