using System.Linq;
using Microsoft.Office.Interop.Excel;
using PionlearClient;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Historicals.ExcelComponent;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.RangeSizeModifier
{
    internal class IndividualLossSetRowInserter : BaseRowInserter
    {
        public override bool Validate(ISegmentExcelMatrix excelMatrix, Range range)
        {
            if (!base.Validate(excelMatrix, range)) return false;

            if (range.GetTopRow() == excelMatrix.GetInputRange().GetBottomRow())
            {
                MessageHelper.Show($"Top selected row can't be on {BexConstants.IndividualLossSetName.ToLower()} totals row", MessageType.Stop);
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
                
                var lossSet = (IndividualLossSet) ((ExcelMatrix) ExcelMatrix).GetParent();

                if (!IgnoreLedger && lossSet.Ledger.Any())
                {
                    var startRowId = StartRow - 1;
                    lossSet.Ledger.InsertItems(startRowId, RowCount);
                }

                if (!IsSelectionOnSecondRow) return;
                lossSet.ExcelMatrix.Reformat();
            }
        }
    }
}