using System.Linq;
using Microsoft.Office.Interop.Excel;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.Models.Historicals;
using SubmissionCollector.Models.Historicals.ExcelComponent;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.RangeSizeModifier
{
    internal class IndividualLossSetRowDeleter : BaseRowDeleter
    {
        public override bool Validate(ISegmentExcelMatrix excelMatrix, Range range)
        {
            if (!base.Validate(excelMatrix, range)) return false;
            
            if (IsSelectionOnFirstRow)
            {
                const string message = "The first selected cell can't be above the data range";
                MessageHelper.Show(message, MessageType.Stop);
                return false;
            }

            if (IsSelectionOnLastRow)
            {
                const string message = "The first selected cell can't be on the bottom row";
                MessageHelper.Show(message, MessageType.Stop);
                return false;
            }


            if (ExcelRange.Rows.Count - RowCount > MinimumRowCount) return true;

            var sizeMessage = $"Row deleting not allowed when resulting data range would have less than or equal to {MinimumRowCount} rows";
            MessageHelper.Show(sizeMessage, MessageType.Stop);
            return false;
        }

        public override void ModifyRange()
        {
            using (new ExcelScreenUpdateDisabler())
            {
                base.ModifyRange();

                var lossSet = (IndividualLossSet) ((ExcelMatrix) ExcelMatrix).GetParent();
                lossSet.IsDirty = true;

                if (lossSet.Ledger.Any())
                {
                    var startRowId = StartRow - 1;
                    lossSet.Ledger.RemoveItems(startRowId, RowCount);
                }
                

                if (!IsSelectionOnSecondRow && !IsSelectionOnLastRow) return;
                ExcelMatrix.Reformat();
            }
        }

        public override bool IsOkToReformat()
        {
            return !IsSelectionOnSecondRow 
                ? base.IsOkToReformat() 
                : IsOkToReformatCommon(FirstRowWarning);
        }
    }
}
