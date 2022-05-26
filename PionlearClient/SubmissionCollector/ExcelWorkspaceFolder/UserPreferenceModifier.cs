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
    internal class UserPreferenceModifier
    {
        public void Modify(IWorkbookLogger logger)
        {
            try
            {
                var preferences = UserPreferences.ReadFromFile();
                var viewModel = new UserPreferencesViewModel(preferences);
                var selector = new UserPreferencesDisplayer(viewModel);
                var form = new UserPreferencesForm(selector)
                {
                    Height = (int)FormSizeHeight.ExtraLarge,
                    Width = (int)FormSizeWidth.Medium,
                    Text = BexConstants.ApplicationName,
                    StartPosition = FormStartPosition.CenterScreen
                };
                form.ShowDialog();

                if (selector.DialogResult == DialogResult.OK) selector.UserPreferences.WriteToFile();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"Modify {BexConstants.UserPreferencesName.ToLower()} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }
    }
}
