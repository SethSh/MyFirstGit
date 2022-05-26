using System.Linq;
using Microsoft.Office.Interop.Excel;
using PionlearClient.Extensions;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.RangeSizeModifier
{
    internal class PeriodSetRowInserter : BaseRowInserter
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
            var segment = ExcelMatrix.GetSegment();
            
            using (new ExcelScreenUpdateDisabler())
            {
                base.ModifyRange();

                segment.ExposureSets.ForEach(exposureSet =>
                {
                    ExcelRange = exposureSet.ExcelMatrix.RangeName.GetRange();
                    base.ModifyRange();
                    if (!exposureSet.Ledger.Any()) return;

                    var startRowId = StartRow - 1;
                    exposureSet.Ledger.InsertItems(startRowId, RowCount);
                });

                segment.AggregateLossSets.ForEach(lossSet =>
                {
                    ExcelRange = lossSet.ExcelMatrix.RangeName.GetRange();
                    base.ModifyRange();
                    if (!lossSet.Ledger.Any()) return;

                    var startRowId = StartRow - 1;
                    lossSet.Ledger.InsertItems(startRowId, RowCount);
                });
            }

            if (!IsSelectionOnSecondRow) return;
            
            using (new ExcelScreenUpdateDisabler())
            {
                ExcelMatrix.Reformat();

                segment.ExposureSets.ForEach(exposureSet =>
                {
                    var range = exposureSet.CommonExcelMatrix.RangeName.GetRange();
                    range.SetInvisibleRangeName(exposureSet.CommonExcelMatrix.RangeName);
                    exposureSet.CommonExcelMatrix.Reformat();
                });

                segment.AggregateLossSets.ForEach(lossSet =>
                {
                    var range = lossSet.CommonExcelMatrix.RangeName.GetRange();
                    range.SetInvisibleRangeName(lossSet.CommonExcelMatrix.RangeName);
                    lossSet.CommonExcelMatrix.Reformat();
                    
                });

            }
        }
    }
}