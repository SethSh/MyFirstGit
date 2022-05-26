using System;
using System.Text;
using System.Windows.Forms;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.View;
using SubmissionCollector.View.Enums;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class PolicyProfileDimensionHelpManager
    {
        private const string LimitAbbreviation = "Lim";
        private const string SirAttachmentAbbreviation = "SIR"; 
        
        public void Help(IWorkbookLogger logger)
        {
            try
            {
                var sb = new StringBuilder();
                AddFlat(sb);
                AddLimitBySir(sb);
                AddSirByLimit(sb);

                ShowPolicyProfileDimensionAlternatives(sb.ToString());
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"Show {BexConstants.PolicyProfileName.ToLower()} dimension help failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        private static void AddFlat(StringBuilder sb)
        {
            sb.AppendLine("Option #1: Flat");
            sb.AppendLine("================================================");
            sb.AppendLine($"{BexConstants.LimitName.PadRight(20)} " +
                          $"{BexConstants.SirAttachmentName.PadRight(20)} " +
                          $"{BexConstants.PercentName.PadRight(20)}");
            sb.AppendLine($"{"1,000,000",-20} {"50,000",-20} {"20 %",-20}");
            sb.AppendLine($"{"1,000,000",-20} {"75,000",-20} {"30 %",-20}");
            sb.AppendLine($"{"2,000,000",-20} {"50,000",-20} {"35 %",-20}");
            sb.AppendLine($"{"2,000,000",-20} {"75,000",-20} {"15 %",-20}");
            sb.AppendLine();
            sb.AppendLine();
            sb.AppendLine();
        }

        private static void AddLimitBySir(StringBuilder sb)
        {
            var title = $"{LimitAbbreviation} {RangeExtensions.DownArrow} {SirAttachmentAbbreviation} {RangeExtensions.RightArrow}".PadRight(20);
            sb.AppendLine($"Option #2: {BexConstants.LimitName} by {BexConstants.SirAttachmentName}");
            sb.AppendLine("================================================");
            sb.AppendLine($"{title} {"50,000",-20} {"75,000",-20}");
            sb.AppendLine($"{"1,000,000",-20} {"20 %",-20} {"30 %",-20}");
            sb.AppendLine($"{"2,000,000",-20} {"35 %",-20} {"15 %",-20}");
            sb.AppendLine();
            sb.AppendLine();
            sb.AppendLine();
        }

        private static void AddSirByLimit(StringBuilder sb)
        {
            sb.AppendLine($"Option #3: {BexConstants.SirAttachmentName} by {BexConstants.LimitName}");
            sb.AppendLine("================================================");

            var title = $"{SirAttachmentAbbreviation} {RangeExtensions.DownArrow} {LimitAbbreviation} {RangeExtensions.RightArrow}".PadRight(20);
            sb.AppendLine($"{title} {"1,000,000",-20} {"2,000,000",-20}");
            sb.AppendLine($"{"50,000",-20} {"20 %",-20} {"30 %",-20}");
            sb.AppendLine($"{"75,000",-20} {"35 %",-20} {"15 %",-20}");
        }

        private static void ShowPolicyProfileDimensionAlternatives(string message)
        {
            var alternatives = new PolicyProfileDimensionAlternatives(message);
            var form = new PolicyProfileDimensionsForm(alternatives)
            {
                Height = (int)FormSizeHeight.ExtraLarge,
                Width = (int)FormSizeWidth.Medium,
                Text = BexConstants.ApplicationName,
                StartPosition = FormStartPosition.CenterScreen
            };
            form.ShowDialog();
        }
    }
}
