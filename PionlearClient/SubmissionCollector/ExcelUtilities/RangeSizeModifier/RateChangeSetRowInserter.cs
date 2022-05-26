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
    internal class RateChangeSetRowInserter : BaseRowInserter
    {
        public override bool Validate(ISegmentExcelMatrix excelMatrix, Range range)
        {
            if (!base.Validate(excelMatrix, range)) return false;
            
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

                var rateChangeSet = (RateChangeSet)((RateChangeExcelMatrix)ExcelMatrix).GetParent();

                if (!IgnoreLedger && rateChangeSet.Ledger.Any())
                {
                    var startRowId = StartRow - 1;
                    rateChangeSet.Ledger.InsertItems(startRowId, RowCount);
                }

                if (!IsSelectionOnSecondRow) return;
                rateChangeSet.ExcelMatrix.Reformat();
            }
        }
    }
}
