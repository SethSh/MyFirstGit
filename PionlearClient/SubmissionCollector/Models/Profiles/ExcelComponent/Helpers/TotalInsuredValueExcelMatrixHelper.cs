using System;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Profiles.ExcelComponent.Helpers
{
    internal class TotalInsuredValueExcelMatrixHelper : BaseMultiOccurenceExcelMatrixHelper<TotalInsuredValueProfile>
    {
        public TotalInsuredValueExcelMatrixHelper(ISegment segment): base(segment)
        {
            ColumnCount = UserPrefs.IsTotalInsuredValueProfileExpanded ? 5 : 2;
            ComponentName = BexConstants.TotalInsuredValueProfileName;
            ExcelRangeName = ExcelConstants.TotalInsuredValueProfileRangeName;
        }

        public override void InsertRanges(Range anchorRange, MultipleOccurrenceSegmentExcelMatrix excelMatrix)
        {
            //componentIndex is subline code
            var componentIndex = excelMatrix.ComponentId;
            var rangeName = TotalInsuredValueExcelMatrix.GetRangeName(Segment.Id, componentIndex);
            var headerRangeName = TotalInsuredValueExcelMatrix.GetHeaderRangeName(Segment.Id, componentIndex);
            var basisRangeName = TotalInsuredValueExcelMatrix.GetBasisRangeName(Segment.Id, componentIndex);
            
            anchorRange.Resize[1, ColumnCountPlusOne].InsertColumnsToRight();
            anchorRange = anchorRange.Offset[0, -ColumnCountPlusOne].GetTopLeftCell();
            var topLeftRange = anchorRange;

            var range = topLeftRange.Resize[UserPrefs.TotalInsuredValueProfileRowCount + 3, ColumnCount];
            range.GetFirstRow().SetInvisibleRangeName(headerRangeName);
            range.GetRangeSubset(1, 0).SetInvisibleRangeName(rangeName);
            range.GetRangeSubset(1, 0).GetTopRightCell().SetInvisibleRangeName(basisRangeName);

            ExcelComponents.Single(ec => ec.ExcelMatrix.RangeName == rangeName).IsExpanded = UserPrefs.IsTotalInsuredValueProfileExpanded;
            
            if (!(excelMatrix is MultipleOccurrenceProfileExcelMatrix em)) throw new InvalidCastException($"Can't insert {ComponentName.ToLower()} profile");

            em.ProfileFormatter = ProfileFormatterFactory.Create(UserPrefs.ProfileBasisId);
            em.SetProfileBasisInWorksheet();
            
            excelMatrix.Reformat();
        }
    }
}
