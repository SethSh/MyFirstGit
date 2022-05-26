using System.Linq;
using Microsoft.Office.Interop.Excel;
using PionlearClient.Extensions;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Historicals.ExcelComponent;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelUtilities.RangeSizeModifier
{
    internal class PeriodSetRowDeleter : BaseRowDeleter
    {
        public override bool Validate(ISegmentExcelMatrix excelMatrix, Range range)
        {
            if (!base.Validate(excelMatrix, range)) return false;

            if (ExcelRange.Rows.Count - RowCount > MinimumRowCount) return true;

            var message = $"Row deleting is not allowed when resulting range would have less than or equal to {MinimumRowCount} rows";
            MessageHelper.Show(message, MessageType.Stop);
            return false;
        }

        public override void ModifyRange()
        {
            var segment = ExcelMatrix.GetSegment();
            using (new ExcelScreenUpdateDisabler())
            {
                base.ModifyRange();

                segment.IsDirty = true;
                segment.PeriodSet.IsDirty = true;

                segment.ExposureSets.ForEach(expoSet =>
                {
                    ExcelRange = expoSet.ExcelMatrix.RangeName.GetRange();
                    base.ModifyRange();

                    if (!expoSet.Ledger.Any()) return;

                    var startRowId = StartRow - 1;
                    expoSet.Ledger.RemoveItems(startRowId, RowCount);
                });

                segment.AggregateLossSets.ForEach(lossSet =>
                {
                    ExcelRange = lossSet.ExcelMatrix.RangeName.GetRange();
                    base.ModifyRange();

                    if (!lossSet.Ledger.Any()) return;

                    var startRowId = StartRow - 1;
                    lossSet.Ledger.RemoveItems(startRowId, RowCount);
                });


                if (IsSelectionOnSecondRow)
                {
                    ((PeriodSetExcelMatrix)ExcelMatrix).ReformatBorderTop();
                    segment.ExposureSets.ForEach(exposureSet => exposureSet.ExcelMatrix.ReformatBorderTop());
                    segment.AggregateLossSets.ForEach(lossSet => lossSet.ExcelMatrix.ReformatBorderTop());
                }

                if (IsSelectionOnLastRow)
                {
                    using (new ExcelScreenUpdateDisabler())
                    {
                        ((PeriodSetExcelMatrix) ExcelMatrix).ReformatBorderBottom();
                        segment.ExposureSets.ForEach(exposureSet => exposureSet.ExcelMatrix.ReformatBorderBottom());
                        segment.AggregateLossSets.ForEach(lossSet => lossSet.ExcelMatrix.ReformatBorderBottom());
                    }
                }
            }
        }

        public override bool IsOkToReformat()
        {
            return !IsSelectionOnFirstRow 
                ? base.IsOkToReformat()
                : IsOkToReformatCommon(FirstRowWarning);
        }
    }
}