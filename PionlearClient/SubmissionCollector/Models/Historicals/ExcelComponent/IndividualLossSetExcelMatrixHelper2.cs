using System.Linq;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Historicals.ExcelComponent
{
    internal static class IndividualLossSetExcelMatrixHelper2
    {
        public static void ModifyRangesToReflectChangeToAlaeFormat(ISegment segment)
        {
            var lossSets = segment.IndividualLossSets.Where(x => x.ExcelMatrix.RangeName.ExistsInWorkbook()).ToList();
            var isLossAndAlaeCombined = segment.IndividualLossSetDescriptor.IsLossAndAlaeCombined;

            using (new ExcelEventDisabler())
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    foreach (var set in lossSets)
                    {
                        var excelMatrix = set.ExcelMatrix;
                        excelMatrix.ModifyRangeToReflectChangeToAlaeFormat(isLossAndAlaeCombined);
                        excelMatrix.Reformat();
                    }
                }
            }
        }

        public static void ModifyRangesToReflectChangeToLimit(ISegment segment)
        {
            var lossSets = segment.IndividualLossSets.Where(x => x.ExcelMatrix.RangeName.ExistsInWorkbook()).ToList();
            var isLimitAvailable = segment.IndividualLossSetDescriptor.IsPolicyLimitAvailable;

            using (new ExcelEventDisabler())
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    foreach (var set in lossSets)
                    {
                        var excelMatrix = set.ExcelMatrix;
                        excelMatrix.ModifyRangeToReflectChangeToLimit(isLimitAvailable);
                        excelMatrix.Reformat();
                    }
                }
            }
        }

        public static void ModifyRangesToReflectChangeToAttachment(ISegment segment)
        {
            var lossSets = segment.IndividualLossSets.Where(x => x.ExcelMatrix.RangeName.ExistsInWorkbook()).ToList();
            var isPolicyAttachmentAvailable = segment.IndividualLossSetDescriptor.IsPolicyAttachmentAvailable;

            using (new ExcelEventDisabler())
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    foreach (var set in lossSets)
                    {
                        var excelMatrix = set.ExcelMatrix;
                        excelMatrix.ModifyRangeToReflectChangeToAttachment(isPolicyAttachmentAvailable);
                        excelMatrix.Reformat();
                    }
                }
            }
        }

        public static void ModifyRangesToReflectChangeToPaid(ISegment segment)
        {
            var lossSets = segment.IndividualLossSets.Where(x => x.ExcelMatrix.RangeName.ExistsInWorkbook()).ToList();
            var isPaidAvailable = segment.IndividualLossSetDescriptor.IsPaidAvailable;

            using (new ExcelEventDisabler())
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    foreach (var set in lossSets)
                    {
                        var excelMatrix = set.ExcelMatrix;
                        excelMatrix.ModifyRangeToReflectChangeToPaid(isPaidAvailable);
                        excelMatrix.Reformat();
                    }
                }
            }
        }

        public static void ModifyRangesToReflectChangeToAccidentDate(ISegment segment)
        {
            var lossSets = segment.IndividualLossSets.Where(x => x.ExcelMatrix.RangeName.ExistsInWorkbook()).ToList();
            var isAccidentDateAvailable = segment.IndividualLossSetDescriptor.IsAccidentDateAvailable;

            using (new ExcelEventDisabler())
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    foreach (var set in lossSets)
                    {
                        var excelMatrix = set.ExcelMatrix;
                        excelMatrix.ModifyRangeToReflectChangeToAccidentDate(isAccidentDateAvailable);
                        excelMatrix.Reformat();
                    }
                }
            }
        }

        public static void ModifyRangesToReflectChangeToPolicyDate(ISegment segment)
        {
            var lossSets = segment.IndividualLossSets.Where(x => x.ExcelMatrix.RangeName.ExistsInWorkbook()).ToList();
            var isPolicyDateAvailable = segment.IndividualLossSetDescriptor.IsPolicyDateAvailable;

            using (new ExcelEventDisabler())
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    foreach (var set in lossSets)
                    {
                        var excelMatrix = set.ExcelMatrix;
                        excelMatrix.ModifyRangeToReflectChangeToPolicyDate(isPolicyDateAvailable);
                        excelMatrix.Reformat();
                    }
                }
            }
        }

        public static void ModifyRangesToReflectChangeToReportDate(ISegment segment)
        {
            var lossSets = segment.IndividualLossSets.Where(x => x.ExcelMatrix.RangeName.ExistsInWorkbook()).ToList();
            var isReportDateAvailable = segment.IndividualLossSetDescriptor.IsReportDateAvailable;

            using (new ExcelEventDisabler())
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    foreach (var set in lossSets)
                    {
                        var excelMatrix = set.ExcelMatrix;
                        excelMatrix.ModifyRangeToReflectChangeToReportDate(isReportDateAvailable);
                        excelMatrix.Reformat();
                    }
                }
            }
        }

        public static void ModifyRangesToReflectChangeToEventCode(ISegment segment)
        {
            var lossSets = segment.IndividualLossSets.Where(x => x.ExcelMatrix.RangeName.ExistsInWorkbook()).ToList();
            var isEventCodeAvailable = segment.IndividualLossSetDescriptor.IsEventCodeAvailable;

            using (new ExcelEventDisabler())
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    foreach (var set in lossSets)
                    {
                        var excelMatrix = set.ExcelMatrix;
                        excelMatrix.ModifyRangeToReflectChangeToEventCode(isEventCodeAvailable);
                        excelMatrix.Reformat();
                    }
                }
            }
        }

    }
}
