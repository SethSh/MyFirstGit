using SubmissionCollector.Models.Segment;

namespace SubmissionCollector.ExcelUtilities.StateSort
{
    internal interface IStateSorter
    {
        ISegment Segment { get; set; }
        bool Validate();
        void Sort();
    }
}
