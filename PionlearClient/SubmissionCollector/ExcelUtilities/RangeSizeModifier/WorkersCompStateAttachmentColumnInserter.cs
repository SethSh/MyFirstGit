using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Profiles.ExcelComponent;

namespace SubmissionCollector.ExcelUtilities.RangeSizeModifier
{
    internal class WorkersCompStateAttachmentColumnInserter : BaseColumnInserter
    {
        protected override int FrozenLeftColumnCount => 3;
        protected override int FrozenRightColumnCount => 0;
    }

    internal class WorkersCompStateAttachmentColumnDeleter : BaseColumnDeleter
    {
        protected override int MinimumInputColumnCount => 3;
        protected override int LabelColumnCount => 3;
        protected override int FrozenLeftColumnCount => 3;
        protected override int FrozenRightColumnCount => 1;

        public override void ModifyRange()
        {
            using (new ExcelEventDisabler())
            {
                using (new ExcelScreenUpdateDisabler())
                {
                    if (IsSelectionOnLastColumn)
                    {
                        var basisRangeName = WorkersCompStateAttachmentExcelMatrix.GetBasisRangeName(ExcelMatrix.SegmentId);
                        ExcelMatrix.GetHeaderRange().GetTopLeftCell().Offset[0, StartColumn - 1].Name = basisRangeName;
                    }

                    base.ModifyRange();
                }
            }
        }
    }
}