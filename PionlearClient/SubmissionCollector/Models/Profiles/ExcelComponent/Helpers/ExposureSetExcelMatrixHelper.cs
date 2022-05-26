using System;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using PionlearClient.BexReferenceData;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Historicals.ExcelComponent;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Profiles.ExcelComponent.Helpers
{
    internal class ExposureSetExcelMatrixHelper : BaseMultiOccurenceExcelMatrixHelper<ExposureSet>
    {
        public ExposureSetExcelMatrixHelper(ISegment segment): base(segment)
        {
            ColumnCount = 1;
            IsPreviousRangeAdjacent = true;
            ExcelRangeName = ExcelConstants.ExposureSetRangeName;
            ComponentName = BexConstants.ExposureSetName;
        }

        public override void InsertRanges(Range anchorRange, MultipleOccurrenceSegmentExcelMatrix excelMatrix)
        {
            var worksheetSublineRowCount = GetWorksheetSublineRowCount();
            var periodCount = WorksheetManager.GetPeriodCount(Segment);

            var componentIndex = excelMatrix.ComponentId;
            var sublinesHeaderRangeName = ExposureSetExcelMatrix.GetSublinesHeaderRangeName(Segment.Id, componentIndex);
            var sublinesRangeName = ExposureSetExcelMatrix.GetSublinesRangeName(Segment.Id, componentIndex);
            var headerRangeName = ExposureSetExcelMatrix.GetHeaderRangeName(Segment.Id, componentIndex);
            var rangeName = ExposureSetExcelMatrix.GetRangeName(Segment.Id, componentIndex);

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

            var basisRange = excelMatrix.GetInputLabelRange();
            basisRange.Value2 = ExposureBasisFromBex.GetExposureBasisName(Convert.ToInt16(Segment.HistoricalExposureBasis));

            excelMatrix.Reformat();
        }
    }
}
