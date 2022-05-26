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
    internal class HazardExcelMatrixHelper : BaseMultiOccurenceExcelMatrixHelper<HazardProfile>
    {
        public HazardExcelMatrixHelper(ISegment segment) : base(segment)
        {
            ColumnCount = 2;
            IsPreviousRangeAdjacent = false;
            ComponentName = BexConstants.HazardProfileName;
            ExcelRangeName = ExcelConstants.HazardProfileRangeName;
        }

        public override void InsertRanges(Range anchorRange, MultipleOccurrenceSegmentExcelMatrix excelMatrix)
        {
            //componentIndex is subline code
            var componentIndex = excelMatrix.ComponentId;
            var headerRangeName = HazardExcelMatrix.GetHeaderRangeName(Segment.Id, componentIndex);
            var rangeName = HazardExcelMatrix.GetRangeName(Segment.Id, componentIndex);
            var basisRangeName = HazardExcelMatrix.GetBasisRangeName(Segment.Id, componentIndex);

            anchorRange.Resize[1, ColumnCountPlusOne].InsertColumnsToRight();
            anchorRange = anchorRange.Offset[0, -ColumnCountPlusOne].GetTopLeftCell();
            var topLeftRange = anchorRange;

            var typeNames = HazardCodesFromBex.ReferenceData
                .Where(x => x.SubLineOfBusinessCode == componentIndex)
                .OrderBy(x => x.DisplayOrder).Select(data => data.Name).ToList();

            var range = topLeftRange.Resize[typeNames.Count + 2, ColumnCount];
            range.GetFirstRow().SetInvisibleRangeName(headerRangeName);
            range.GetRangeSubset(1, 0).SetInvisibleRangeName(rangeName);
            range.GetRangeSubset(1, 0).GetTopRightCell().SetInvisibleRangeName(basisRangeName);

            excelMatrix.GetInputLabelRange().Value = typeNames.ToNByOneArray();

            if (!(excelMatrix is MultipleOccurrenceProfileExcelMatrix em)) throw new InvalidCastException($"Can't insert {ComponentName.ToLower()} profile");

            em.ProfileFormatter = ProfileFormatterFactory.Create(UserPrefs.ProfileBasisId);
            em.SetProfileBasisInWorksheet();

            excelMatrix.Reformat();
        }
    }
}
