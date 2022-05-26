using Microsoft.Office.Interop.Excel;
using SubmissionCollector.Models.DataComponents;

namespace SubmissionCollector.Models.Segment.DataComponents
{
    public interface ISegmentExcelMatrix : IExcelMatrix , IRangeHidable
    {
        int SegmentId { get; set; }
        int InterDisplayOrder { get; }
        int IntraDisplayOrder { get; set; }
        string HeaderRangeName { get; }
        Range GetBodyHeaderRange();
        Range GetBodyRange();
        Range GetHeaderRange();
        ISegment GetSegment();
    }
}
