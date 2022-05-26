using System;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.PolicyProfileDimensionConverter;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class PolicyProfileDimensionManager
    {
        public void ChangeToSirByLimit(IWorkbookLogger logger)
        {
            try
            {
                var validator = new PolicyProfileDimensionValidator();
                if (!validator.Validate()) return;

                var excelMatrix = validator.PolicyExcelMatrix;
                if (excelMatrix.Dimension is SirByLimitPolicyProfileDimension)
                {
                    MessageHelper.Show($"{BexConstants.PolicyProfileName} dimension is already " +
                                       $"{BexConstants.SirAttachmentName} By {BexConstants.LimitName}", MessageType.Stop);
                    return;
                }

                using (new ExcelScreenUpdateDisabler())
                {
                    using (new ExcelEventDisabler())
                    {
                        excelMatrix.Dimension.ConvertToSirByLimit();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"Change to {BexConstants.SirAttachmentName} by {BexConstants.LimitName} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        public void ChangeToLimitBySir(IWorkbookLogger logger)
        {
            try
            {
                var validator = new PolicyProfileDimensionValidator();
                if (!validator.Validate()) return;

                var excelMatrix = validator.PolicyExcelMatrix;
                if (excelMatrix.Dimension is LimitBySirPolicyProfileDimension)
                {
                    MessageHelper.Show($"{BexConstants.PolicyProfileName} dimension is already Limit By SIR/Attachment", MessageType.Stop);
                    return;
                }

                using (new ExcelScreenUpdateDisabler())
                {
                    using (new ExcelEventDisabler())
                    {
                        excelMatrix.Dimension.ConvertToLimitBySir();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                var message = $"Change to {BexConstants.LimitName} by {BexConstants.SirAttachmentName} failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        public void ChangeToFlat(IWorkbookLogger logger)
        {
            try
            {
                var validator = new PolicyProfileDimensionValidator();
                if (!validator.Validate()) return;

                var excelMatrix = validator.PolicyExcelMatrix;
                if (excelMatrix.Dimension is FlatPolicyProfileDimension)
                {
                    MessageHelper.Show($"{BexConstants.PolicyProfileName} dimension is already Flat", MessageType.Stop);
                    return;
                }

                using (new ExcelScreenUpdateDisabler())
                {
                    using (new ExcelEventDisabler())
                    {
                        excelMatrix.Dimension.ConvertToFlat();
                    }
                }

            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Change to Flat failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }
    }
}
