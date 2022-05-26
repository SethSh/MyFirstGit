using Microsoft.Office.Interop.Excel;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.RangeSizeModifier
{
    internal abstract class BaseColumnDeleter : BaseColumnOperator
    {
        public bool IsSelectionOnLastColumn { get; set; }

        protected abstract int MinimumInputColumnCount { get; }
        protected abstract int LabelColumnCount { get; }
        

        public override bool Validate(ISegmentExcelMatrix excelMatrix, Range range)
        {
            SetCommonProperties(excelMatrix, range);
            IsSelectionOnLastColumn = ExcelRange.IsSelectionIntersectingLastColumn();

            var firstSelectedColumnIndex = range.GetTopLeftCell().Column;
            var maximumFrozenLeftColumnIndex = ExcelRange.GetTopLeftCell().Column + FrozenLeftColumnCount - 1;
            if (firstSelectedColumnIndex <= maximumFrozenLeftColumnIndex)
            {
                var message = $"Column deleting not allowed in first {FrozenLeftColumnCount} column(s)";
                MessageHelper.Show(message, MessageType.Stop);
                return false;
            }

            var minimumFrozenRightColumnIndex = ExcelRange.GetTopRightCell().Column - FrozenRightColumnCount + 1;
            if (firstSelectedColumnIndex >= minimumFrozenRightColumnIndex)
            {
                var message = $"Column deleting not allowed in last {FrozenRightColumnCount} column(s)";
                MessageHelper.Show(message, MessageType.Stop);
                return false;
            }

            if (range.GetTopRightCell().Column > ExcelRange.GetTopRightCell().Column)
            {
                const string message = "Columns to the right of data range can't be deleted";
                MessageHelper.Show(message, MessageType.Stop);
                return false;
            }

            var candidateColumnCount = ExcelRange.Columns.Count - SelectedColumnCount;
            var overallMinimumColumnCount = MinimumInputColumnCount + LabelColumnCount;
            if (candidateColumnCount < overallMinimumColumnCount)
            {
                var message = $"Column deleting is not allowed when resulting range would have less than {MinimumInputColumnCount} data columns";
                MessageHelper.Show(message, MessageType.Stop);
                return false;
            }

            return true;
        }

        public override bool IsOkToReformat()
        {
            if (IsSelectionOnSecondColumn)
            {
                const string message = "Column deleting, when the second column is selected, will reformat entire matrix.";
                return IsOkToReformatCommon(message);
            }

            if (IsSelectionOnLastColumn)
            {
                const string message = "Column deleting, when the last column is selected, will reformat entire matrix.";
                return IsOkToReformatCommon(message);
            }

            return true;
        }

        public override void ModifyRange()
        {
            using (new ExcelScreenUpdateDisabler())
            {
                ExcelRange.Resize[1, SelectedColumnCount].Offset[0, StartColumn].EntireColumn.Delete();
                if (IsSelectionOnSecondColumn || IsSelectionOnLastColumn) ExcelMatrix.Reformat();
            }
        }

    }
}