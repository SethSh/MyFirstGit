namespace MramUwpfLibrary.ExposureRatingModel.Discretize
{
    public interface IDiscretizeInput
    {
        int BucketWidth { get; set; }
        int BucketCount { get; set; }
    }

    public class DiscretizeInput : IDiscretizeInput
    {
        public DiscretizeInput(int bucketCount, int bucketWidth)
        {
            BucketCount = bucketCount;
            BucketWidth = bucketWidth;
        }
        public int BucketWidth { get; set; }
        public int BucketCount { get; set; }
    }
}