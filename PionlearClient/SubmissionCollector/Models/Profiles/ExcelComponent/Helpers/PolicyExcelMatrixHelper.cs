using System;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Profiles.ExcelComponent.Helpers
{
    internal class PolicyExcelMatrixHelper : BaseMultiOccurenceExcelMatrixHelper<PolicyProfile>
    {
        public PolicyExcelMatrixHelper(ISegment segment) : base(segment)
        {
            ColumnCount = 3;
            IsPreviousRangeAdjacent = false;
            ComponentName = BexConstants.PolicyProfileName;
            ExcelRangeName = ExcelConstants.PolicyProfileRangeName;
        }

        public override void InsertRanges(Range anchorRange, MultipleOccurrenceSegmentExcelMatrix excelMatrix)
        {
            var worksheetSublineRowCount = GetWorksheetSublineRowCount();

            var componentIndex = excelMatrix.ComponentId;
            var sublinesHeaderRangeName = PolicyExcelMatrix.GetSublinesHeaderRangeName(Segment.Id, componentIndex);
            var sublinesRangeName = PolicyExcelMatrix.GetSublinesRangeName(Segment.Id, componentIndex);
            var headerRangeName = PolicyExcelMatrix.GetHeaderRangeName(Segment.Id, componentIndex);
            var rangeName = PolicyExcelMatrix.GetRangeName(Segment.Id, componentIndex);
            var basisRangeName = PolicyExcelMatrix.GetBasisRangeName(Segment.Id, componentIndex);
            
            anchorRange.Resize[1, ColumnCountPlusOne].InsertColumnsToRight();
            anchorRange = anchorRange.Offset[0, -ColumnCountPlusOne].GetTopLeftCell();

            var sublinesRange = anchorRange.Offset[-(worksheetSublineRowCount + ExcelConstants.InBetweenRowCount), 0].Resize[worksheetSublineRowCount, ColumnCount];
            sublinesRange.GetFirstRow().SetInvisibleRangeName(sublinesHeaderRangeName);
            sublinesRange.GetRangeSubset(1, 0).SetInvisibleRangeName(sublinesRangeName);

            var range = anchorRange.Resize[UserPrefs.PolicyProfileRowCount + 3, ColumnCount];
            range.GetFirstRow().SetInvisibleRangeName(headerRangeName);
            range.GetRangeSubset(1, 0).SetInvisibleRangeName(rangeName);
            range.GetRangeSubset(1, 0).GetTopRightCell().SetInvisibleRangeName(basisRangeName);

            if (!(excelMatrix is MultipleOccurrenceProfileExcelMatrix em)) throw new InvalidCastException($"Can't insert {ComponentName.ToLower()} profile");

            em.ProfileFormatter = ProfileFormatterFactory.Create(UserPrefs.ProfileBasisId);
            em.SetProfileBasisInWorksheet();

            excelMatrix.Reformat();
        }
    }
}
