using System;
using System.Collections.Generic;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class RangeNameDisplayer
    {
        internal static void SetVisibility(bool showRangeNames, IWorkbookLogger logger)
        {
            try
            {
                var filters = new List<string>
                {
                    $"{ExcelConstants.SegmentTemplateRangeName}.",
                };

                var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;
                filters.Add($"{ExcelConstants.SubmissionRangeName}.");
                foreach (var segment in package.Segments)
                {
                    filters.Add($"{ExcelConstants.SegmentRangeName}{segment.Id}.");
                }

                RangeExtensions.SetWorkbookRangeNamesVisibility(filters, showRangeNames);
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Show Range Names failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }
    }
}
