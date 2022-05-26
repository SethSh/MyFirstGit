using System.Linq;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using PionlearClient.BexReferenceData;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Profiles.ExcelComponent.Helpers
{
    internal class WorkersCompStateHazardGroupExcelMatrixHelper : BaseSingleOccurenceExcelMatrixHelper<WorkersCompStateHazardGroupProfile>
    {
        public WorkersCompStateHazardGroupExcelMatrixHelper(ISegment segment) : base(segment)
        {
            var hazardGroupCount = WorkersCompClassCodesAndHazardsFromBex.HazardGroups.Count();
            
            ColumnCount = hazardGroupCount + WorkersCompStateHazardGroupExcelMatrix.ColumnLabelCount;
            IsPreviousRangeAdjacent = false;
            ComponentName = BexConstants.WorkersCompStateHazardGroupProfileName;
            ExcelRangeName = ExcelConstants.WorkersCompStateHazardProfileRangeName;
        }

        public override void InsertRange(Range anchorRange, SingleOccurrenceProfileExcelMatrix excelMatrix)
        {
            const int stateAbbreviationIndex = 0;
            const int stateNameIndex = 1;
            
            var states = StateCodesFromBex.GetWorkersCompStates().ToList();
            var stateAbbreviations = states.Select(x => x.Abbreviation).ToList();
            var stateNames = states.Select(x => x.Name).ToList();

            var hazards = WorkersCompClassCodesAndHazardsFromBex.HazardGroups;
            var hazardNames = hazards.OrderBy(hazard => hazard.DisplayOrder).Select(hazard => hazard.Name).ToList();
            
            var rangeName = WorkersCompStateHazardGroupExcelMatrix.GetRangeName(Segment.Id);
            var headerRangeName = WorkersCompStateHazardGroupExcelMatrix.GetHeaderRangeName(Segment.Id);
            var basisRangeName = WorkersCompStateHazardGroupExcelMatrix.GetBasisRangeName(Segment.Id);

            anchorRange.Resize[1, ColumnCountPlusOne].InsertColumnsToRight();
            anchorRange = anchorRange.Offset[0, -ColumnCountPlusOne].GetTopLeftCell();
            var topLeftRange = anchorRange;
            
            var range = topLeftRange.Resize[stateAbbreviations.Count + 1 + WorkersCompStateHazardGroupExcelMatrix.RowLabelCount, ColumnCount];
            range.GetFirstRow().SetInvisibleRangeName(headerRangeName);
            range.GetRangeSubset(1, 0).SetInvisibleRangeName(rangeName);
            range.GetTopRightCell().SetInvisibleRangeName(basisRangeName);

            var em = (WorkersCompStateHazardGroupExcelMatrix) excelMatrix;
            excelMatrix.GetInputLabelRange().GetColumn(stateAbbreviationIndex).Value = stateAbbreviations.ToNByOneArray();
            excelMatrix.GetInputLabelRange().GetColumn(stateNameIndex).Value = stateNames.ToNByOneArray();
            em.GetHazardGroupRange().Value = hazardNames.ToArray();

            em.GetHazardGroupRange().Offset[-1, 0].Value2 = 0;
            em.GetStatePremiumsRange().Value2 = 0;

            excelMatrix.Reformat();
        }
    }
}
