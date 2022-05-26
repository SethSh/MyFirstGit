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
    internal class OccupancyTypeExcelMatrixHelper : BaseMultiOccurenceExcelMatrixHelper<OccupancyTypeProfile>
    {
        public OccupancyTypeExcelMatrixHelper(ISegment segment) : base(segment)
        {
            ColumnCount = 2;
            IsPreviousRangeAdjacent = false;
            ComponentName = BexConstants.OccupancyTypeProfileName;
            ExcelRangeName = ExcelConstants.OccupancyTypeProfileRangeName;
        }
        
        public override void InsertRanges(Range anchorRange, MultipleOccurrenceSegmentExcelMatrix excelMatrix)
        {
            var typeNames = OccupancyTypeCodesFromBex.ReferenceData.OrderBy(data => data.DisplayOrder).Select(data => data.Name).ToList();

            //componentIndex is subline code
            var componentIndex = excelMatrix.ComponentId;
            var rangeName = OccupancyTypeExcelMatrix.GetRangeName(Segment.Id, componentIndex);
            var headerRangeName = OccupancyTypeExcelMatrix.GetHeaderRangeName(Segment.Id, componentIndex);
            var basisRangeName = OccupancyTypeExcelMatrix.GetBasisRangeName(Segment.Id, componentIndex);
            
            anchorRange.Resize[1, ColumnCountPlusOne].InsertColumnsToRight();
            anchorRange = anchorRange.Offset[0, -ColumnCountPlusOne].GetTopLeftCell();
            var topLeftRange = anchorRange;

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
