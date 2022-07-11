namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials
{
    public class Parameters : IParameters
    {
        public Parameters(
            double[] means,
            double[] weights,
            double[] alaePercents,
            double probabilityOfNoLoss,
            double alaeForClaimsWithoutPay)
        {
            Means = means;
            Weights = weights;
            AlaePercents = alaePercents;
            ProbabilityOfNoLoss = probabilityOfNoLoss;
            AlaeForClaimsWithoutPay = alaeForClaimsWithoutPay;
        }

        public double ProbabilityOfNoLoss { get; set; }

        public double AlaeForClaimsWithoutPay { get; set; }

        public double[] AlaePercents { get; set; }

        public double[] Means { get; set; }

        public double[] Weights { get; set; }
    }
}