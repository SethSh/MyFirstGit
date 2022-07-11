namespace MramUwpfLibrary.ExposureRatingModel
{
    public interface IProfileAllocation
    {
        long Id { get; set; }
        double Amount { get; set; }
    }
    
    public class ProfileAllocation: IProfileAllocation
    {
        public long Id { get; set; }
        public double Amount { get; set; }
    }
}
