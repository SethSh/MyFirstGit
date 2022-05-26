using Microsoft.Office.Interop.Excel;
using PionlearClient;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Historicals.ExcelComponent;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Profiles.ExcelComponent.Helpers
{
    internal class IndividualLossSetExcelMatrixHelper : BaseMultiOccurenceExcelMatrixHelper<IndividualLossSet>
    {
        public IndividualLossSetExcelMatrixHelper(ISegment segment) : base(segment)
        {
            ColumnCount = WorksheetManager.GetIndividualColumnCount(segment.IndividualLossSetDescriptor);
            IsPreviousRangeAdjacent = false;
            ExcelRangeName = ExcelConstants.IndividualLossSetRangeName;
            ComponentName = BexConstants.IndividualLossSetName;
        }
        
        public override void InsertRanges(Range anchorRange, MultipleOccurrenceSegmentExcelMatrix excelMatrix)
        {
            var worksheetSublineRowCount = GetWorksheetSublineRowCount();

            var componentIndex = excelMatrix.ComponentId;
            var sublinesHeaderRangeName = ExcelMatrix.GetSublinesHeaderRangeName(Segment.Id, componentIndex);
            var sublinesRangeName = ExcelMatrix.GetSublinesRangeName(Segment.Id, componentIndex);
            var headerRangeName = ExcelMatrix.GetHeaderRangeName(Segment.Id, componentIndex);
            var rangeName = ExcelMatrix.GetRangeName(Segment.Id, componentIndex);
            var thresholdRangeName = ExcelMatrix.GetThresholdRangeName(Segment.Id, componentIndex);

            anchorRange.Resize[1, ColumnCountPlusOne].InsertColumnsToRight();
            anchorRange = anchorRange.Offset[0, -ColumnCountPlusOne].GetTopLeftCell();

            var sublinesRange = anchorRange.Offset[-(worksheetSublineRowCount + ExcelConstants.InBetweenRowCount), 0]
                .Resize[worksheetSublineRowCount, ColumnCount];
            sublinesRange.GetFirstRow().SetInvisibleRangeName(sublinesHeaderRangeName);
            sublinesRange.GetRangeSubset(1, 0).SetInvisibleRangeName(sublinesRangeName);

            var range = anchorRange.Resize[UserPrefs.IndividualLossCount + 3, ColumnCount];
            range.GetFirstRow().Offset[0, 2].Resize[1, ColumnCount - 2].SetInvisibleRangeName(headerRangeName);
            range.GetFirstRow().Resize[1, 2].SetInvisibleRangeName(thresholdRangeName);
            range.GetRangeSubset(1, 0).SetInvisibleRangeName(rangeName);

            excelMatrix.Reformat();
        }
    }
}

