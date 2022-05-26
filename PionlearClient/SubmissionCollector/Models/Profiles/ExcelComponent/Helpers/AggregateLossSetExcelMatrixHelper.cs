using Microsoft.Office.Interop.Excel;
using PionlearClient;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Historicals.ExcelComponent;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Profiles.ExcelComponent.Helpers
{
    internal class AggregateLossSetExcelMatrixHelper : BaseMultiOccurenceExcelMatrixHelper<AggregateLossSet>
    {
        public AggregateLossSetExcelMatrixHelper(ISegment segment): base(segment)
        {
            ColumnCount = WorksheetManager.GetAggregateColumnCount(segment.AggregateLossSetDescriptor);
            IsPreviousRangeAdjacent = true;
            ExcelRangeName = ExcelConstants.AggregateLossSetRangeName;
            ComponentName = BexConstants.AggregateLossSetName;
        }
        
        public override void InsertRanges(Range anchorRange, MultipleOccurrenceSegmentExcelMatrix excelMatrix)
        {
            var worksheetSublineRowCount = GetWorksheetSublineRowCount();
            var periodCount = WorksheetManager.GetPeriodCount(Segment);

            var componentIndex = excelMatrix.ComponentId;
            var sublinesHeaderRangeName = AggregateLossSetExcelMatrix.GetSublinesHeaderRangeName(Segment.Id, componentIndex);
            var sublinesRangeName = AggregateLossSetExcelMatrix.GetSublinesRangeName(Segment.Id, componentIndex);
            var headerRangeName = AggregateLossSetExcelMatrix.GetHeaderRangeName(Segment.Id, componentIndex);
            var rangeName = AggregateLossSetExcelMatrix.GetRangeName(Segment.Id, componentIndex);

            anchorRange.Resize[1, ColumnCount].InsertColumnsToRight();
            anchorRange = anchorRange.Offset[0, -ColumnCount].GetTopLeftCell();
            anchorRange.EntireColumn.Clear();
            anchorRange.EntireColumn.Locked = false;

            var sublinesRange = anchorRange.Offset[-(worksheetSublineRowCount + ExcelConstants.InBetweenRowCount), 0].Resize[worksheetSublineRowCount, ColumnCount];
            sublinesRange.GetFirstRow().SetInvisibleRangeName(sublinesHeaderRangeName);
            sublinesRange.GetRangeSubset(1, 0).SetInvisibleRangeName(sublinesRangeName);

            var range = anchorRange.Resize[periodCount + 3, ColumnCount];
            range.GetFirstRow().SetInvisibleRangeName(headerRangeName);
            range.GetRangeSubset(1, 0).RemoveLastRow().SetInvisibleRangeName(rangeName);

            excelMatrix.Reformat();
        }
    }
}
