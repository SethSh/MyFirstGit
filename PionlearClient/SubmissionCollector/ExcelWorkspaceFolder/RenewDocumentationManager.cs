using System;
using System.Text;
using PionlearClient;
using PionlearClient.Extensions;
using SubmissionCollector.Enums;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class RenewDocumentationManager
    {
        public static void Document(IWorkbookLogger logger)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("Renew Features");
                sb.AppendLine("==============");
                sb.AppendLine();

                sb.AppendLine("Validates workbook");
                sb.AppendLine("Advances package underwriting year by one year");
                sb.AppendLine();

                sb.AppendLine($"For each {BexConstants.SegmentName.ToLower()}:");
                sb.AppendLine($"\t inserts one additional {BexConstants.PeriodName.ToLower()} row");
                sb.AppendLine($"\t advances {BexConstants.PeriodName.ToLower()} by one year");
                sb.AppendLine($"\t advances {BexConstants.PeriodName.ToLower()} {BexConstants.EvaluationDateName.ToLower()} by one year");

                sb.AppendLine();
                sb.AppendLine("Copies source IDs into predecessor IDs");
                sb.AppendLine($"{BexConstants.DecoupleName.ToStartOfSentence()}s workbook from {BexConstants.ServerDatabaseName}");
                sb.AppendLine("Prompts for Save As:");
                sb.AppendLine("\t records workbook name of renewed source");
                sb.AppendLine("\t records date and time of renewal");
                sb.AppendLine();
                sb.AppendLine();
                sb.AppendLine("Note: renew advancing will override cell formulas with values");
                sb.AppendLine("Note: Save As is a default assumption and optional");

                MessageHelper.Show(sb.ToString());
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Show renew documentation failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }
    }
}
