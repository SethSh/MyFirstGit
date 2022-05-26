using Microsoft.Office.Interop.Excel;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.RangeSizeModifier
{
    internal abstract class BaseColumnInserter : BaseColumnOperator
    {
        private const int MaximumSelectionColumnCount = 20;

        public override bool Validate(ISegmentExcelMatrix excelMatrix, Range range)
        {
            SetCommonProperties(excelMatrix, range);

            var firstSelectedColumnIndex = range.GetTopLeftCell().Column;
            
            var minimumFrozenRightColumnIndex = ExcelRange.GetTopRightCell().Column - FrozenRightColumnCount + 1;
            if (firstSelectedColumnIndex >= minimumFrozenRightColumnIndex)
            {
                var message = $"Column inserting not allowed in last {FrozenRightColumnCount} column(s)";
                MessageHelper.Show(message, MessageType.Stop);
                return false;
            }
            
            if (SelectedColumnCount > MaximumSelectionColumnCount)
            {
                var message = $"Column inserting has a practical maximum of {MaximumSelectionColumnCount:N0}";
                MessageHelper.Show(message, MessageType.Stop);
                return false;
            }

            var maximumFrozenLeftColumnIndex = ExcelRange.GetTopLeftCell().Column + FrozenLeftColumnCount - 1;
            if (firstSelectedColumnIndex <= maximumFrozenLeftColumnIndex)
            {
                var message = $"Column inserting not allowed in first {FrozenLeftColumnCount} column(s)";
                MessageHelper.Show(message, MessageType.Stop);
                return false;
            }

            return true;
        }
        
        
        public override void ModifyRange()
        {
            using (new ExcelScreenUpdateDisabler())
            {
                ExcelRange.Resize[1, SelectedColumnCount].Offset[0, StartColumn].InsertColumnsToRight();
                if (IsSelectionOnSecondColumn) ExcelMatrix.Reformat();
            }
        }

        public override bool IsOkToReformat()
        {
            if (!IsSelectionOnSecondColumn) return true;

            const string message = "Column inserting, when the second column is selected, will reformat entire matrix.";
            return IsOkToReformatCommon(message);
        }
    }
}