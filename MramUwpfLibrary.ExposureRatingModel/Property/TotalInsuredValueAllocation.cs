namespace MramUwpfLibrary.ExposureRatingModel.Property
{
    public interface ITotalInsuredValueAllocation
    {
        double TotalInsuredValue { get; set; }
        double? Share { get; set; }
        double? Limit { get; set; }
        double? Attachment { get; set; }
        double Amount { get; set; }
    }

    public class TotalInsuredValueAllocation : ITotalInsuredValueAllocation
    {
        public double TotalInsuredValue { get; set; }
        public double? Share { get; set; }
        public double? Limit { get; set; }
        public double? Attachment { get; set; }
        public double Amount { get; set; }
    }
}