using System;
using System.Linq;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities;
using SubmissionCollector.Models;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.View;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class CheckboxSynchronizationFixer
    {
        internal static void Fix(IWorkbookLogger logger)
        {
            try
            {
                var identifier = new SegmentWorksheetValidator();
                if (!identifier.Validate()) return;

                var segment = identifier.Segment;

                Segment.IsCurrentlyFixing = true;
                SetTotalInsuredValueCheckbox(segment);
                SetAggregateCheckboxes(segment);
                SetIndividualCheckboxes(segment);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Checkbox synchronization failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
            finally
            {
                Segment.IsCurrentlyFixing = false;
            }
        }

        private static void SetAggregateCheckboxes(ISegment segment)
        {
            var labelsRange = segment.AggregateLossSets.First().ExcelMatrix.GetInputLabelRange();
            var labels = labelsRange.GetContent().ForceContentToStrings().GetRow(0).ToList();

            var descriptor = segment.AggregateLossSetDescriptor;
            descriptor.IsPaidAvailable = labels.Contains(BexConstants.PaidLossName) || labels.Contains(BexConstants.PaidLossAndAlaeName);
            descriptor.IsLossAndAlaeCombined = labels.Contains(BexConstants.ReportedLossAndAlaeName);
        }

        private static void SetIndividualCheckboxes(ISegment segment)
        {
            var labelsRange = segment.IndividualLossSets.First().ExcelMatrix.GetInputLabelRange();
            var labels = labelsRange.GetContent().ForceContentToStrings().GetRow(0).ToList();

            var descriptor = segment.IndividualLossSetDescriptor;
            descriptor.IsEventCodeAvailable = labels.Contains(BexConstants.EventCodeName);
            descriptor.IsPaidAvailable = labels.Contains(BexConstants.PaidLossName) || labels.Contains(BexConstants.PaidLossAndAlaeName);
            descriptor.IsLossAndAlaeCombined = labels.Contains(BexConstants.ReportedLossAndAlaeName);
            descriptor.IsPolicyLimitAvailable = labels.Contains(BexConstants.LimitName);
            descriptor.IsPolicyAttachmentAvailable = labels.Contains(BexConstants.AttachName);
            descriptor.IsAccidentDateAvailable = labels.Contains(BexConstants.AccidentDateName);
            descriptor.IsReportDateAvailable = labels.Contains(BexConstants.ReportDateName);
            descriptor.IsPolicyDateAvailable = labels.Contains(BexConstants.PolicyDateName);
        }

        private static void SetTotalInsuredValueCheckbox(ISegment segment)
        {
            if (!segment.ContainsProperty()) return;

            const int unExpandedCount = 2;
            foreach (var profile in segment.TotalInsuredValueProfiles)
            {
                var labelsRange = profile.ExcelMatrix.GetInputLabelRange();
                profile.IsExpanded = labelsRange.Columns.Count > unExpandedCount;
            }
        }
    }
}
