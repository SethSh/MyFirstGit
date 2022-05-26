using System;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using PionlearClient.BexReferenceData;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Profiles.ExcelComponent.Helpers
{
    internal class StateExcelMatrixHelper : BaseMultiOccurenceExcelMatrixHelper<StateProfile>
    {
        public StateExcelMatrixHelper(ISegment segment): base(segment)
        {
            ColumnCount = 3;
            IsPreviousRangeAdjacent = false;
            ExcelRangeName = ExcelConstants.StateProfileRangeName;
            ComponentName = BexConstants.StateProfileName;
        }
        
        public override void InsertRanges(Range anchorRange, MultipleOccurrenceSegmentExcelMatrix excelMatrix)
        {
            var worksheetSublineRowCount = GetWorksheetSublineRowCount();

            var states = StateCodesFromBex.GetLiabilityStates().ToList();
            var abbreviation = states.Select(state => state.Abbreviation).ToList();
            var names = states.Select(state => state.Name).ToList();

            var componentIndex = excelMatrix.ComponentId;
            var sublinesHeaderRangeName = StateExcelMatrix.GetSublinesHeaderRangeName(Segment.Id, componentIndex);
            var sublinesRangeName = StateExcelMatrix.GetSublinesRangeName(Segment.Id, componentIndex);
            var headerRangeName = StateExcelMatrix.GetHeaderRangeName(Segment.Id, componentIndex);
            var rangeName = StateExcelMatrix.GetRangeName(Segment.Id, componentIndex);
            var basisRangeName = StateExcelMatrix.GetBasisRangeName(Segment.Id, componentIndex);

            anchorRange.Resize[1, ColumnCountPlusOne].InsertColumnsToRight();
            anchorRange = anchorRange.Offset[0, -ColumnCountPlusOne].GetTopLeftCell();
            var topLeftRange = anchorRange;

            var sublinesRange = topLeftRange.Offset[-(worksheetSublineRowCount + ExcelConstants.InBetweenRowCount), 0].Resize[worksheetSublineRowCount, ColumnCount];
            sublinesRange.GetFirstRow().SetInvisibleRangeName(sublinesHeaderRangeName);
            sublinesRange.GetRangeSubset(1, 0).SetInvisibleRangeName(sublinesRangeName);

            var range = topLeftRange.Resize[states.Count + 2, ColumnCount];

            range.GetFirstRow().SetInvisibleRangeName(headerRangeName);
            range.GetRangeSubset(1, 0).SetInvisibleRangeName(rangeName);
            range.GetRangeSubset(1, 0).GetTopRightCell().SetInvisibleRangeName(basisRangeName);

            excelMatrix.GetInputLabelRange().GetColumn(0).Value = abbreviation.ToNByOneArray();
            excelMatrix.GetInputLabelRange().GetColumn(1).Value = names.ToNByOneArray();
            
            if (!(excelMatrix is MultipleOccurrenceProfileExcelMatrix em)) throw new InvalidCastException($"Can't insert {ComponentName.ToLower()} profile"); 
            
            em.ProfileFormatter = ProfileFormatterFactory.Create(UserPrefs.ProfileBasisId);
            em.SetProfileBasisInWorksheet();

            excelMatrix.Reformat();
        }
    }
}
