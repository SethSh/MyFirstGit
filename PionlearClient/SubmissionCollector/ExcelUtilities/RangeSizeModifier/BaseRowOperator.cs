using Microsoft.Office.Interop.Excel;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Segment.DataComponents;

namespace SubmissionCollector.ExcelUtilities.RangeSizeModifier
{
    internal abstract class BaseRowOperator : BaseRangeResizer
    {
        public bool IsSelectionOnFirstRow { get; set; }
        public bool IsSelectionOnSecondRow { get; set; }
        public int StartRow { get; set; }
        public int RowCount { get; set; }

        public override void SetCommonProperties(ISegmentExcelMatrix excelMatrix, Range range)
        {
            base.SetCommonProperties(excelMatrix, range);

            IsSelectionOnFirstRow = ExcelRange.IsTopRowEqual(range);
            IsSelectionOnSecondRow = ExcelRange.GetRangeSubset(1,0).IsTopRowEqual(range);
            StartRow = range.Row - ExcelRange.GetTopLeftCell().Row;
            RowCount = range.Rows.Count;
        }
    }
}
