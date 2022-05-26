using System;
using System.Linq;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class TotalInsuredValueExpandManager
    {
        public static void ChangeExpand(IWorkbookLogger logger)
        {
            try
            {
                ChangeExpand();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                MessageHelper.Show("TIV Range Expand Failed", MessageType.Stop);
            }
        }

        private static void ChangeExpand()
        {
            var rangeValidator = new SegmentWorksheetValidator();
            if (!rangeValidator.Validate()) return;

            var selectedRange = rangeValidator.SelectedRange;

            var rangeNames = rangeValidator.Segment.TotalInsuredValueProfiles.Select(x => x.CommonExcelMatrix.RangeName);
            var selectedRangeName = string.Empty;
            foreach (var item in rangeNames)
            {
                if (!item.ContainsRange(selectedRange)) continue;
                selectedRangeName = item;
                break;
            }

            if (string.IsNullOrEmpty(selectedRangeName))
            {
                MessageHelper.Show($"The selection must be within a {BexConstants.TotalInsuredValueProfileName.ToLower()} range",
                    MessageType.Stop);
                return;
            }

            var profile = rangeValidator.Segment.TotalInsuredValueProfiles.Single(x => x.CommonExcelMatrix.RangeName == selectedRangeName);
            profile.IsExpanded = !profile.IsExpanded;
            profile.ExcelMatrix.SynchronizeExpansion();

            //only do when removing previously visible columns
            if (!profile.IsExpanded)
            {
                profile.IsDirty = true;
            }
        }
    }
}