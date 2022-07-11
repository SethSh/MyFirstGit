using MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials;

namespace MramUwpfLibrary.ExposureRatingModel.Property
{
    public class PropertyCurveParameters: IMixedExponentialParametersCore
    {
        public double[] Means { get; set; }
        public double[] Weights { get; set; }
    }
}
