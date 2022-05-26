using Microsoft.Office.Interop.Excel;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Subline;

namespace SubmissionCollector.Models.Segment.DataComponents
{
    public abstract class BaseSegmentExcelMatrix : BaseExcelMatrix, ISegmentExcelMatrix
    {
        protected BaseSegmentExcelMatrix()
        {

        }

        protected BaseSegmentExcelMatrix(int segmentId)
        {
            SegmentId = segmentId;
        }


        public int SegmentId { get; set; }

        public int InterDisplayOrder { get; set; }
        public int IntraDisplayOrder { get; set; }

        public string HeaderRangeName => $"{RangeName}.{ExcelConstants.HeaderRangeName}";

        public void ShowColumns(bool hasEmptyColumnToRight = true)
        {
            if (hasEmptyColumnToRight)
            {
                RangeName.GetRange().AppendColumn().ShowColumns();
            }
            else
            {
                RangeName.GetRange().ShowColumns();
            }
        }

        public void HideColumns(bool hasEmptyColumnToRight = true)
        {
            if (hasEmptyColumnToRight)
            {
                RangeName.GetRange().AppendColumn().HideColumns();
            }
            else
            {
                RangeName.GetRange().HideColumns();
            }
        }

        public Range GetHeaderRange()
        {
            return RangeName.GetRange().GetRow(0).Offset[-1, 0];
        }

        public ISegment GetSegment()
        {
            return Globals.ThisWorkbook.ThisExcelWorkspace.Package.GetSegment(SegmentId);
        }

        public virtual void SetColumnVisibility(ISubline subline)
        {
            if (!(this is MultipleOccurrenceSegmentExcelMatrix multipleOccurrenceSegmentExcelMatrix)) return;

            if (multipleOccurrenceSegmentExcelMatrix.Contains(subline))
            {
                ShowColumns(HasEmptyColumnToRight);
            }
            else
            {
                HideColumns(HasEmptyColumnToRight);
            }
        }

        public abstract Range GetBodyHeaderRange();

        public abstract Range GetBodyRange();

        public virtual void MoveRangesWhenSublinesChange(int sublineCount)
        {
            const int inBetweenRowCount = ExcelConstants.InBetweenRowCount;

            var segment = GetSegment();
            var prospectiveRange = segment.ProspectiveExposureAmountExcelMatrix.GetInputRange();
            var prospectiveBottomRow = prospectiveRange.Row;
            var desiredTopRow = prospectiveBottomRow + sublineCount + inBetweenRowCount;

            var headerRange = GetHeaderRange();
            var headerRangeTopRow = headerRange.GetTopLeftCell().Row;

            if (desiredTopRow < headerRangeTopRow)
            {
                var delta = headerRangeTopRow - desiredTopRow;
                headerRange.Offset[-delta - inBetweenRowCount, 0].Resize[delta, headerRange.Columns.Count].DeleteRangeUp();
            }
            else if (headerRangeTopRow < desiredTopRow)
            {
                InsertHeaderRows(prospectiveRange, desiredTopRow);
            }
        }

        public virtual void InsertHeaderRows(Range prospectiveRange, int desiredTopRow)
        {
            const int inBetweenRowCount = ExcelConstants.InBetweenRowCount; 
            var headerRange = GetHeaderRange();
            var headerRangeTopRow = headerRange.GetTopLeftCell().Row;

            var delta = desiredTopRow - headerRangeTopRow;
            var insertRange = headerRange.Offset[-inBetweenRowCount, 0].Resize[delta, headerRange.Columns.Count];
            insertRange.InsertRangeDown();
        }
    }
}

