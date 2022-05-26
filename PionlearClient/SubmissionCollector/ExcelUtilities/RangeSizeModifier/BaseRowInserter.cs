using Microsoft.Office.Interop.Excel;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.RangeSizeModifier
{
    internal abstract class BaseRowInserter : BaseRowOperator
    {
        private const int MaximumRowCount = 10000;

        public override void ModifyRange()
        {
            ExcelRange
                .Offset[StartRow, 0]
                .Resize[RowCount, ExcelRange.Columns.Count]
                .InsertRangeDown();
        }

        public override bool Validate(ISegmentExcelMatrix excelMatrix, Range range)
        {
            SetCommonProperties(excelMatrix, range);

            if (RowCount <= MaximumRowCount) return true;

            var message = $"Select fewer rows.  Row inserting has a practical max of {MaximumRowCount:N0}";
            MessageHelper.Show(message, MessageType.Stop);
            return false;
        }

        public override bool IsOkToReformat()
        {
            if (!IsSelectionOnSecondRow) return true;
            
            const string message = "When the first row is selected, row inserting will reformat entire range.";
            return IsOkToReformatCommon(message);
        }

        public bool IgnoreLedger { get; set; }
    }
}
