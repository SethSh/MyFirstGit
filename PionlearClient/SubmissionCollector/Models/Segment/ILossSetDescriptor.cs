namespace SubmissionCollector.Models.Segment
{
    public interface ILossSetDescriptor
    {
        bool IsLossAndAlaeCombined { get; set; }
        bool IsPaidAvailable { get; set; }
    }
}