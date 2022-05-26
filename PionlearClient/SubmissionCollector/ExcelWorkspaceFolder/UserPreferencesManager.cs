using System;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class UserPreferencesManager
    {
        public static void ResetUserPreferences(IWorkbookLogger logger)
        {
            try
            {
                var userPreferences = new UserPreferences();
                userPreferences.CreateNew();
                userPreferences.WriteToFile();

                MessageHelper.Show("User preferences reset to defaults");
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"Reset {BexConstants.UserPreferencesName.ToLower()} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        public static void ShowJson(IWorkbookLogger logger)
        {
            try
            {
                var json = GetUserPreferencesAsJson();
                MessageHelper.Show(json, MessageType.Information);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                MessageHelper.Show("Show json failed", MessageType.Stop);
            }
        }

        private static string GetUserPreferencesAsJson()
        {
            var readFromFile = UserPreferences.ReadFromFile();
            return SerializeManager.ConvertToJson(readFromFile);
        }
    }
}
