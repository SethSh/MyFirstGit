using System.Linq;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using PionlearClient.BexReferenceData;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Profiles.ExcelComponent.Helpers
{
    internal class MinnesotaRetentionExcelMatrixHelper : BaseSingleOccurenceExcelMatrixHelper<MinnesotaRetention>
    {
        public MinnesotaRetentionExcelMatrixHelper(ISegment segment) : base(segment)
        {
            ColumnCount = 2;
            IsPreviousRangeAdjacent = false;
            ComponentName = BexConstants.MinnesotaRetentionName;
            ExcelRangeName = ExcelConstants.MinnesotaRetentionRangeName;
        }

        public override void InsertRange(Range anchorRange, SingleOccurrenceProfileExcelMatrix excelMatrix)
        {
            var rangeName = MinnesotaRetentionExcelMatrix.GetRangeName(Segment.Id);
            var headerRangeName = MinnesotaRetentionExcelMatrix.GetHeaderRangeName(Segment.Id);
            var basisRangeName = MinnesotaRetentionExcelMatrix.GetBasisRangeName(Segment.Id);

            anchorRange.Resize[1, ColumnCountPlusOne].InsertColumnsToRight();
            anchorRange = anchorRange.Offset[0, -ColumnCountPlusOne].GetTopLeftCell();
            var topLeftRange = anchorRange;

            var range = topLeftRange.Resize[2, ColumnCount];
            range.GetFirstRow().SetInvisibleRangeName(headerRangeName);
            range.GetRangeSubset(1, 0).SetInvisibleRangeName(rangeName);
            range.GetFirstRow().GetTopRightCell().SetInvisibleRangeName(basisRangeName);  //ignore this

            var retentions = MinnesotaRetentionsFromBex.ReferenceData.OrderBy(retention => retention.RetentionAmount).Select(retention => retention.RetentionAmount).ToList();
            range.GetRangeSubset(1, 0).GetTopRightCell().Value2 = retentions.First();
            
            var retentionsInDropdown = string.Join(", ", retentions);
            range.Validation.Delete();
            range.Validation.Add(XlDVType.xlValidateList, Formula1: retentionsInDropdown);

            excelMatrix.Reformat();
        }
    }
}
