using SubmissionCollector.Models.Segment.DataComponents;

namespace SubmissionCollector.Models.DataComponents
{
    public abstract class SingleOccurrenceSegmentExcelMatrix : BaseSegmentExcelMatrix
    {
        public override string RangeName => $"segment{SegmentId}.{ExcelRangeName}";
        
        protected SingleOccurrenceSegmentExcelMatrix(int segmentId) : base(segmentId)
        {

        }
    }
}
