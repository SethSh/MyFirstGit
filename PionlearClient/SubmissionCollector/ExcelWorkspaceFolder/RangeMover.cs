using System;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class RangeMover
    {
        public static void MoveRight(IWorkbookLogger logger)
        {
            try
            {
                var validator = new RangeMoverValidator();
                if (!validator.Validate()) return;

                var excelMatrix = (MultipleOccurrenceSegmentExcelMatrix)validator.ExcelMatrix;
                if (!excelMatrix.IsOkToMoveRight) return;

                var originalSelection = Globals.ThisWorkbook.GetSelectedRange();
                using (new ExcelEventDisabler())
                {
                    using (new ExcelScreenUpdateDisabler())
                    {
                        excelMatrix.MoveRight();
                    }
                }
                originalSelection.Select();
                excelMatrix.SwitchDisplayOrderAfterMoveRight();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Move right failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        public static void MoveLeft(IWorkbookLogger logger)
        {
            try
            {
                var validator = new RangeMoverValidator();
                if (!validator.Validate()) return;

                var excelMatrix = (MultipleOccurrenceSegmentExcelMatrix)validator.ExcelMatrix;
                if (!excelMatrix.IsOkToMoveLeft) return;

                var originalSelection = Globals.ThisWorkbook.GetSelectedRange();
                using (new ExcelEventDisabler())
                {
                    using (new ExcelScreenUpdateDisabler())
                    {
                        excelMatrix.MoveLeft();
                    }
                }
                originalSelection.Select();
                excelMatrix.SwitchDisplayOrderAfterMoveLeft();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Move left failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }
    }
}
