using System.Linq;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using PionlearClient.BexReferenceData;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Profiles.ExcelComponent.Helpers
{
    internal class WorkersCompClassCodeExcelMatrixHelper : BaseSingleOccurenceExcelMatrixHelper<WorkersCompClassCodeProfile>
    {
        public WorkersCompClassCodeExcelMatrixHelper(ISegment segment) : base(segment)
        {
            ColumnCount = 5;
            IsPreviousRangeAdjacent = false;
            ComponentName = BexConstants.WorkersCompClassCodeProfileName;
            ExcelRangeName = ExcelConstants.WorkersCompClassCodeProfileRangeName;
        }

        public override void InsertRange(Range anchorRange, SingleOccurrenceProfileExcelMatrix excelMatrix)
        {
            const int metaHeaderRows = 1;
            const int topBufferRows = 1;
            const int bufferRows = metaHeaderRows + topBufferRows + WorkersCompClassCodeExcelMatrix.BufferRowCount;

            var rangeName = WorkersCompClassCodeExcelMatrix.GetRangeName(Segment.Id);
            var headerRangeName = WorkersCompClassCodeExcelMatrix.GetHeaderRangeName(Segment.Id);
            var basisRangeName = WorkersCompClassCodeExcelMatrix.GetBasisRangeName(Segment.Id);

            anchorRange.Resize[1, ColumnCountPlusOne].InsertColumnsToRight();
            anchorRange = anchorRange.Offset[0, -ColumnCountPlusOne].GetTopLeftCell();
            var topLeftRange = anchorRange;

            var range = topLeftRange.Resize[UserPrefs.WorkersCompClassCodeCount + bufferRows, ColumnCount];
            range.GetFirstRow().SetInvisibleRangeName(headerRangeName);
            range.GetRangeSubset(1, 0).SetInvisibleRangeName(rangeName);

            var stateColumn = excelMatrix.GetInputRange().GetFirstColumn().RemoveLastRows(2);
            stateColumn.Validation.Delete(); 
            
            var states = WorkersCompClassCodesAndHazardsFromBex.StateClassCodes
                .Select(cc => cc.State)
                .Select(state => state.Abbreviation)
                .Distinct()
                .OrderBy(state => state);
            var statesInDropdown = string.Join(", ", states);

            var worksheet = Segment.WorksheetManager.Worksheet;
            if (worksheet.ProtectContents)
            {
                worksheet.UnprotectInterface();
                stateColumn.Validation.Add(XlDVType.xlValidateList, Formula1: statesInDropdown);
                worksheet.ProtectInterface();
            }
            else
            {
                stateColumn.Validation.Add(XlDVType.xlValidateList, Formula1: statesInDropdown);
            }
            
            
            excelMatrix.GetBodyHeaderRange().GetTopLeftCell().Offset[0, 2].SetInvisibleRangeName(basisRangeName);
            excelMatrix.Reformat();
        }
    }
}
