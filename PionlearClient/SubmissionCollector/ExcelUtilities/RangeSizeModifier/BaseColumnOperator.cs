using Microsoft.Office.Interop.Excel;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Segment.DataComponents;

namespace SubmissionCollector.ExcelUtilities.RangeSizeModifier
{
    internal abstract class BaseColumnOperator : BaseRangeResizer
    {
        public bool IsSelectionOnFirstColumn { get; set; }
        public bool IsSelectionOnSecondColumn { get; set; }
        public int StartColumn { get; set; }
        public int SelectedColumnCount { get; set; }
        protected abstract int FrozenLeftColumnCount { get; }
        protected abstract int FrozenRightColumnCount { get; }

        public override void SetCommonProperties(ISegmentExcelMatrix excelMatrix, Range range)
        {
            base.SetCommonProperties(excelMatrix, range);
            StartColumn = range.Column - ExcelRange.GetTopLeftCell().Column;
            SelectedColumnCount = range.Columns.Count;
            IsSelectionOnFirstColumn = StartColumn == 0;
            IsSelectionOnSecondColumn = StartColumn == 1;
        }
    }
}
