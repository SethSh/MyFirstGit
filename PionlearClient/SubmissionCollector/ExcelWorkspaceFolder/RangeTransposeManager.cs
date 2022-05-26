using System;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelUtilities.RangeTransposer;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment.DataComponents;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class RangeTransposeManager
    {
        public void Transpose(IWorkbookLogger logger)
        {
            try
            {
                Transpose();
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                const string message = "Range Transpose failed";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        private static void Transpose()
        {
            var identifier = new SegmentExcelComponentIdentifier();
            if (!identifier.Validate()) return;

            if (!(identifier.ExcelMatrix is IRangeTransposable))
            {
                MessageHelper.Show(@"The selection must be within a transpose-able range", MessageType.Stop);
                return;
            }

            var segment = identifier.Segment;
            var excelMatrix = identifier.ExcelMatrix;
            var isTransposed = ((IRangeTransposable) excelMatrix).IsTransposed;

            var rangeTransposer = isTransposed
                ? new RangeReverseTransposer(segment, excelMatrix)
                : (IRangeTransposer) new RangeTransposer(segment, excelMatrix);

            rangeTransposer.TransposeWrapper();
        }
    }
}