using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient;
using PionlearClient.Extensions;
using SubmissionCollector.BexCommunication;
using SubmissionCollector.Enums;
using SubmissionCollector.Models.Package;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class RatingAnalysisManager
    {
        public static void ShowRatingAnalyses(IPackage package, IWorkbookLogger logger)
        {
            try
            {
                var packageName = package.Name;
                string message;
                if (!package.SourceId.HasValue)
                {
                    message = $"This {BexConstants.PackageName.ToLower()} isn't currently " +
                              $"{BexConstants.AttachName.ToLower()}ed to any {BexConstants.BexName} " +
                              $"{BexConstants.RatingAnalysisName.ToLower()}.";
                }
                else
                {
                    var sourceId = package.SourceId.Value;
                    var bexCommunicationManager = new BexCommunicationManager();
                    IList<SubmissionPackageAttachedTo> ratingAnalyses;
                    try
                    {
                        ratingAnalyses = bexCommunicationManager.GetRatingAnalysisIdsUsingThisPackage(sourceId).OrderBy(ra => ra.Name).ToList();
                    }
                    catch (FileNotFoundException fileNotFoundException)
                    {
                        Console.WriteLine(fileNotFoundException);
                        MessageHelper.Show(fileNotFoundException.Message, MessageType.Warning);
                        return;
                    }

                    if (ratingAnalyses.Any())
                    {
                        var ratingAnalysesAsString = string.Join(Environment.NewLine, ratingAnalyses.Select(ra => ra.ToFriendlyString()));

                        message = $"{BexConstants.PackageName.ToStartOfSentence()} <{packageName}> is currently " +
                                  $"{BexConstants.AttachName.ToLower()}ed to the following " +
                                  $"{BexConstants.BexName} {BexConstants.RatingAnalysisName.ToLower()}:" +
                                  Environment.NewLine + Environment.NewLine +
                                  ratingAnalysesAsString;
                    }
                    else
                    {
                        message = $"{BexConstants.PackageName.ToStartOfSentence()} <{packageName}> isn't currently " +
                                  $"{BexConstants.AttachName.ToLower()}ed to a " +
                                  $"{BexConstants.BexName} {BexConstants.RatingAnalysisName.ToLower()}";
                    }
                }
                MessageHelper.Show(message, MessageType.Information);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"Show {BexConstants.RatingAnalysesName.ToLower()} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }
    }

    internal static class RatingAnalysisExtensions
    {
        internal static string ToFriendlyString(this SubmissionPackageAttachedTo ra)
        {
            Debug.Assert(ra.PeriodBegin.HasValue, "Period Begin can't be blank");
            var commonString = ra.Name + Environment.NewLine +
                               ra.CedentName + Environment.NewLine +
                               ra.PeriodBegin.Value.Year + Environment.NewLine;
            
            Debug.Assert(ra.IsLocked != null, "ra.IsLocked != null");
            var extraString = ra.IsLocked.Value ? "Locked" + Environment.NewLine : string.Empty;
            
            return commonString + extraString;
        }
    }
}
