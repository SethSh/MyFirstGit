namespace SubmissionCollector.Models.DataComponents
{
    public abstract class SingleOccurrenceProfileExcelMatrix : BaseProfileExcelMatrix
    {
        protected SingleOccurrenceProfileExcelMatrix(int segmentId) : base(segmentId)
        {

        }
        public override string RangeName => $"segment{SegmentId}.{ExcelRangeName}";
    }
}