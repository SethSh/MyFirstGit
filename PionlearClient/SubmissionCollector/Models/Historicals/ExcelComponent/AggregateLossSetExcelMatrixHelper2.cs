using System.Linq;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Historicals.ExcelComponent
{
    internal static class AggregateLossSetExcelMatrixHelper2
    {
        public static void ModifyRangesToReflectChangeToAlaeFormat(ISegment segment)
        {
            var lossSets = segment.AggregateLossSets.Where(x => x.ExcelMatrix.RangeName.ExistsInWorkbook()).ToList();
            var isLossAndAlaeCombined = segment.AggregateLossSetDescriptor.IsLossAndAlaeCombined;

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

        public static void ModifyRangesToReflectPaids(ISegment segment)
        {
            var lossSets = segment.AggregateLossSets.Where(x => x.ExcelMatrix.RangeName.ExistsInWorkbook()).ToList();
            var isPaidAvailable = segment.AggregateLossSetDescriptor.IsPaidAvailable;

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

    }
}
