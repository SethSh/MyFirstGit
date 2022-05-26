using System.Linq;
using Microsoft.Office.Interop.Excel;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.Models.Profiles.ExcelComponent.Helpers
{
    internal abstract class BaseSingleOccurenceExcelMatrixHelper<T> : BaseExcelMatrixHelper<T> where T : IExcelComponent
    {
        protected BaseSingleOccurenceExcelMatrixHelper(ISegment segment) : base(segment)
        {
        
        }

        public void CreateRange()
        {
            if (!ExcelComponents.Any()) return;

            var adjacentRange = GetAdjacentRange();

            var excelMatrix = (SingleOccurrenceProfileExcelMatrix)ExcelComponents.Single().CommonExcelMatrix;
            var anchorRange = adjacentRange.Offset[0, 1];
            InsertRange(anchorRange, excelMatrix);
        }

        public void ModifyRange()
        {
            var rangeName = $"segment{Segment.Id}.{ExcelRangeName}";
            if (ExcelComponents.Any())
            {
                if (!rangeName.ExistsInWorkbook()) CreateRange();
            }
            else
            {
                if (rangeName.ExistsInWorkbook()) DeleteOrphanRanges(rangeName);
            }
        }
        
        public abstract void InsertRange(Range anchorRange, SingleOccurrenceProfileExcelMatrix excelMatrix);

        private static void DeleteOrphanRanges(string rangeName)
        {
            rangeName.GetRange().AppendColumn().EntireColumn.Delete();
            rangeName.DeleteRangeName();
            
            WorkbookExtensions.RemoveErrorRanges();
        }
    }
}