using Microsoft.Office.Interop.Excel;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.RangeSizeModifier
{
    internal abstract class BaseRowDeleter : BaseRowOperator
    {
        protected const int MinimumRowCount = 5;
        protected const string FirstRowWarning = "When the first row is selected, row deleting will reformat entire range.";
        public bool IsSelectionOnLastRow { get; set; }

        public override void ModifyRange()
        {
            ExcelRange
                .Offset[StartRow, 0]
                .Resize[RowCount, ExcelRange.Columns.Count]
                .DeleteRangeUp();
        }

        public override bool Validate(ISegmentExcelMatrix excelMatrix, Range range)
        {
            SetCommonProperties(excelMatrix, range);
            IsSelectionOnLastRow = range.DoesRangeIntersect(ExcelRange.GetLastRow());

            if (!IsSelectionOnLastRow) return true;
            if (range.GetBottomLeftCell().Row <= ExcelRange.GetBottomLeftCell().Row) return true;

            const string message = "Selected cells below the data range can't be deleted";
            MessageHelper.Show(message, MessageType.Stop);
            return false;
        }

        public override bool IsOkToReformat()
        {
            if (!IsSelectionOnLastRow) return true;
            
            const string message = "Row deleting, when the last row is selected, will reformat entire matrix.";
            return IsOkToReformatCommon(message);
        }
    }
}
