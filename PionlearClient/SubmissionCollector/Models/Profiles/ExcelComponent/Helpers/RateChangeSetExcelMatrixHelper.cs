using Microsoft.Office.Interop.Excel;
using PionlearClient;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Historicals.ExcelComponent;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Profiles.ExcelComponent.Helpers
{
    internal class RateChangeSetExcelMatrixHelper : BaseMultiOccurenceExcelMatrixHelper<RateChangeSet>
    {
        public RateChangeSetExcelMatrixHelper(ISegment segment) : base(segment)
        {
            ColumnCount = 2;
            IsPreviousRangeAdjacent = false;
            ExcelRangeName = ExcelConstants.RateChangeSetRangeName;
            ComponentName = BexConstants.RateChangeSetName;
        }

        public override void InsertRanges(Range anchorRange, MultipleOccurrenceSegmentExcelMatrix excelMatrix)
        {
            var worksheetSublineRowCount = GetWorksheetSublineRowCount();

            var componentIndex = excelMatrix.ComponentId;
            var sublinesHeaderRangeName = RateChangeExcelMatrix.GetSublinesHeaderRangeName(Segment.Id, componentIndex);
            var sublinesRangeName = RateChangeExcelMatrix.GetSublinesRangeName(Segment.Id, componentIndex);
            var headerRangeName = RateChangeExcelMatrix.GetHeaderRangeName(Segment.Id, componentIndex);
            var rangeName = RateChangeExcelMatrix.GetRangeName(Segment.Id, componentIndex);
            
            anchorRange.Resize[1, ColumnCountPlusOne].InsertColumnsToRight();
            anchorRange = anchorRange.Offset[0, -ColumnCountPlusOne].GetTopLeftCell();

            var sublinesRange = anchorRange.Offset[-(worksheetSublineRowCount + ExcelConstants.InBetweenRowCount), 0].Resize[worksheetSublineRowCount, ColumnCount];
            sublinesRange.SetInvisibleRangeName(sublinesRangeName);
            sublinesRange.GetFirstRow().SetInvisibleRangeName(sublinesHeaderRangeName);
            sublinesRange.GetRangeSubset(1, 0).SetInvisibleRangeName(sublinesRangeName);

            var range = anchorRange.Resize[UserPrefs.RateChangeCount + 2, ColumnCount];
            range.SetInvisibleRangeName(rangeName);
            range.GetFirstRow().SetInvisibleRangeName(headerRangeName);
            range.GetRangeSubset(1, 0).SetInvisibleRangeName(rangeName);

            excelMatrix.Reformat();
        }
    }
}

