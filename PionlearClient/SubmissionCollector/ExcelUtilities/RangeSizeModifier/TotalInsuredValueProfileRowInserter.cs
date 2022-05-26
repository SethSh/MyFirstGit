using Microsoft.Office.Interop.Excel;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.RangeSizeModifier
{
    internal class TotalInsuredValueProfileRowInserter : BaseRowInserter
    {
        public override bool Validate(ISegmentExcelMatrix excelMatrix, Range range)
        {
            if (!base.Validate(excelMatrix, range)) return false;

            if (range.GetTopRow() == excelMatrix.GetInputRange().GetBottomRow())
            {
                MessageHelper.Show($"Top selected row can't be on {BexConstants.TotalInsuredValueProfileName.ToLower()} bottom row", MessageType.Stop);
                return false;
            }

            if (!IsSelectionOnFirstRow) return true;
            const string message = "The first selected cell can't be above the data range";
            MessageHelper.Show(message, MessageType.Stop);
            return false;
        }

        public override void ModifyRange()
        {
            using (new ExcelScreenUpdateDisabler())
            {
                base.ModifyRange();
                if (!IsSelectionOnSecondRow) return;

                ExcelMatrix.Reformat();
            }
        }
    }
}
