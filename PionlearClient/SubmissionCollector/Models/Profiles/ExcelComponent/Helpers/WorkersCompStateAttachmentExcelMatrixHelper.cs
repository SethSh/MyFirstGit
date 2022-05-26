using System.Linq;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using PionlearClient.BexReferenceData;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Profiles.ExcelComponent.Helpers
{
    internal class WorkersCompStateAttachmentExcelMatrixHelper : BaseSingleOccurenceExcelMatrixHelper<WorkersCompStateAttachmentProfile>
    {
        public WorkersCompStateAttachmentExcelMatrixHelper(ISegment segment) : base(segment)
        {
            ColumnCount = UserPrefs.WorkersCompStateAttachmentCount + 3;
            IsPreviousRangeAdjacent = false;
            ComponentName = BexConstants.WorkersCompStateAttachmentProfileName;
            ExcelRangeName = ExcelConstants.WorkersCompStateAttachmentProfileRangeName;
        }

        public override void InsertRange(Range anchorRange, SingleOccurrenceProfileExcelMatrix excelMatrix)
        {
            var states = StateCodesFromBex.GetWorkersCompStates().ToList();
            var stateAbbreviations = states.Select(x => x.Abbreviation).ToList();
            var stateNames = states.Select(x => x.Name).ToList();

            var rangeName = WorkersCompStateAttachmentExcelMatrix.GetRangeName(Segment.Id);
            var headerRangeName = WorkersCompStateAttachmentExcelMatrix.GetHeaderRangeName(Segment.Id);
            var basisRangeName = WorkersCompStateAttachmentExcelMatrix.GetBasisRangeName(Segment.Id);

            anchorRange.Resize[1, ColumnCountPlusOne].InsertColumnsToRight();
            anchorRange = anchorRange.Offset[0, -ColumnCountPlusOne ].GetTopLeftCell();
            var topLeftRange = anchorRange;

            var rowCount = stateAbbreviations.Count + 2;
            var range = topLeftRange.Resize[rowCount, ColumnCount];
            range.GetFirstRow().SetInvisibleRangeName(headerRangeName);
            range.GetRangeSubset(1, 0).SetInvisibleRangeName(rangeName);
            range.GetTopRightCell().SetInvisibleRangeName(basisRangeName);

            excelMatrix.GetInputLabelRange().GetFirstColumn().Value = stateAbbreviations.ToNByOneArray();
            excelMatrix.GetInputLabelRange().GetColumn(1).Value = stateNames.ToNByOneArray();
            
            excelMatrix.Reformat();
        }
    }
}
