using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient;
using PionlearClient.TokenAuthentication;
using SubmissionCollector.BexCommunication;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.Models.Package;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class ServerPackageDeleteManager
    {
        public void DeleteWrapper(IPackage package, IWorkbookLogger logger)
        {
            try
            {
                if (WorkbookSaveManager.IsReadOnly)
                {
                    var message = "Can't delete a read-only workbook";
                    MessageHelper.Show(message, MessageType.Stop);
                    return;
                }
                
                if (!package.SourceId.HasValue)
                {
                    MessageHelper.Show($"Delete Failed: this {BexConstants.PackageName.ToLower()} hasn't been uploaded to " +
                                       $"the {BexConstants.ServerDatabaseName.ToLower()}", MessageType.Warning);
                    return;
                }

                var sourceId = package.SourceId.Value;
                var bexCommunication = new BexCommunicationManager();
                IList<SubmissionPackageAttachedTo> ratingAnalyses;
                try
                {
                    ratingAnalyses = bexCommunication.GetRatingAnalysisIdsUsingThisPackage(sourceId).ToList();
                }
#pragma warning disable 168
                catch (FileNotFoundException fileNotFoundException)
#pragma warning restore 168
                {
                    var sb = new StringBuilder();
                    sb.AppendLine("Deletes are blocked");
                    sb.AppendLine(string.Empty);
                    sb.AppendLine($"Either 1) Can't find the {BexConstants.PackageName.ToLower()} or " +
                                  $"2) can't connect to the {BexConstants.BexName} API service.");
                    MessageHelper.Show(sb.ToString(), MessageType.Stop);
                    return;
                }

                if (ratingAnalyses.Any())
                {
                    var message = $"Deletes are blocked because this {BexConstants.PackageName.ToLower()} is {BexConstants.AttachName.ToLower()}ed to " +
                                  $"a {BexConstants.BexName} {BexConstants.RatingAnalysisName.ToLower()}.";

                    MessageHelper.Show(message, MessageType.Stop);
                    return;
                }

                if (Confirm(package.Name) != DialogResult.Yes) return;
                using (new CursorToWait())
                {
                    using (new StatusBarUpdater($"Deleting {BexConstants.PackageName.ToLower()} <{package.Name}> ..."))
                    {
                        Delete(package);
                        WorkbookSaveManager.Save();
                    }
                }

                MessageHelper.Show("Delete Successful", MessageType.Success);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"Delete {BexConstants.PackageName.ToLower()} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        private static void Delete(IPackage package)
        {
            Debug.Assert(package.SourceId.HasValue,"Package has no source id");

            var client = BexCollectorClientFactory.CreateBexCollectorClient(ConfigurationHelper.SecretWord, ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);
            var collectorClient = client.SubmissionPackagesClient;
            var task = collectorClient.DeleteAsync(package.SourceId.Value);
            task.Wait();

            package.BexCommunications.Clear();
            ((Package)package).DecoupleFromServer();
        }

        private static DialogResult Confirm(string name)
        {
            var confirmationMessage = $"Are you sure you want to delete {BexConstants.PackageName.ToLower()} <{name}> from the {BexConstants.ServerDatabaseName.ToLower()}?";
            return MessageHelper.ShowWithYesNo(confirmationMessage);
        }
    }
}

