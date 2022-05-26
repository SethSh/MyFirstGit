using System;
using System.Text;
using PionlearClient;
using PionlearClient.Extensions;
using SubmissionCollector.Enums;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class QualityControlDocumentationManager
    {
        public static void Document(IWorkbookLogger logger)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("Quality Control Features");
                sb.AppendLine("========================");
                sb.AppendLine();

                sb.AppendLine(BexConstants.PackageName);
                sb.AppendLine("\t No checks");
                sb.AppendLine();

                sb.AppendLine(BexConstants.SegmentName);
                sb.AppendLine($"\t Identifies when there\'re gaps between consecutive {BexConstants.PeriodName.ToLower()}s");
                sb.AppendLine($"\t Identifies when {BexConstants.EvaluationDateName.ToLower()}s in the future");
                sb.AppendLine();

                sb.AppendLine(BexConstants.SublineAllocationName);
                sb.AppendLine(BexConstants.UmbrellaAllocationName);
                sb.AppendLine(BexConstants.PolicyProfileName);
                sb.AppendLine(BexConstants.StateProfileName);
                sb.AppendLine(BexConstants.HazardProfileName);
                sb.AppendLine($"\t Identifies when sum of percents are not close enough to {1d:P0}");
                sb.AppendLine();

                sb.AppendLine(BexConstants.UmbrellaAllocationName);
                sb.AppendLine("\t No feature checks");
                sb.AppendLine();

                sb.AppendLine(BexConstants.ExposureSetName);
                sb.AppendLine(BexConstants.AggregateLossSetName);
                sb.AppendLine("\t No feature checks");
                sb.AppendLine();
            
                sb.AppendLine(BexConstants.IndividualLossSetName);
                sb.AppendLine($"\t Identifies when {BexConstants.OccurrenceIdName} equals zero");
                sb.AppendLine("\t Identifies when loss is less than zero");
                sb.AppendLine("\t Identifies when loss (and ALAE if appropriate) is greater than policy limit");
                sb.AppendLine("\t Identifies when paid is greater reported");
                sb.AppendLine();

                sb.AppendLine($"{BexConstants.AggregateLossSetName.ConnectWithDash(BexConstants.IndividualLossSetName)} Consistency");
                sb.AppendLine($"\t Identifies {BexConstants.SegmentName.ToLower()} {BexConstants.PeriodName.ToLower()}s " +
                              $"when the sum of {BexConstants.IndividualLossSetName.ToLower()} is greater than {BexConstants.AggregateLossSetName.ToLower()} ");
            
                MessageHelper.Show(sb.ToString(), MessageType.Documentation);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"Show {BexConstants.DataQualityControlTitle.ToLower()} documentation failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }
    }
}
