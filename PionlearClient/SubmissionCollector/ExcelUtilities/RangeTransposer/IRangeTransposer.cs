using SubmissionCollector.Models.DataComponents;
using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.ExcelUtilities.RangeTransposer
{
    public interface IRangeTransposer
    {
        IExcelMatrix ExcelMatrix { get; set; }
        ISegment Segment { get; set; }
        void TransposeWrapper();
    }
}