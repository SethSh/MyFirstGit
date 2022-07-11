namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials
{
    public interface IParameters : IMixedExponentialParametersCore
    {
        double AlaeForClaimsWithoutPay { get; set; }
        double[] AlaePercents { get; set; }
        double ProbabilityOfNoLoss { get; set; }
    }

    public interface IMixedExponentialParametersCore
    {
        double[] Means { get; set; }
        double[] Weights { get; set; }
    }
}