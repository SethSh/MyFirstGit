namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.TruncatedParetos
{
    public class Parameters : IParameters
    {
        public Parameters(
            double scale,
            double shape,
            double probabilityLessThanTruncation,
            double averageSeverityForSmallLoss,
            double truncation,
            double flatAlaeForLargeLoss,
            double variableAlaeForLargeLoss,
            double variableAlaeForSmallLoss)
        {
            Scale = scale;
            Shape = shape;
            ProbabilityLessThanTruncation = probabilityLessThanTruncation;
            AverageSeverityForSmallLoss = averageSeverityForSmallLoss;
            Truncation = truncation;
            FlatAlaeForLargeLoss = flatAlaeForLargeLoss;
            VariableAlaeForLargeLoss = variableAlaeForLargeLoss;
            VariableAlaeForSmallLoss = variableAlaeForSmallLoss;
        }

        public Parameters(
            double scale,
            double shape,
            double probabilityLessThanTruncation,
            double averageSeverityForSmallLoss,
            double truncation)
        {
            Scale = scale;
            Shape = shape;
            ProbabilityLessThanTruncation = probabilityLessThanTruncation;
            AverageSeverityForSmallLoss = averageSeverityForSmallLoss;
            Truncation = truncation;
        }

        public Parameters()
        {
                
        }
            

        public double VariableAlaeForSmallLoss { get; set; }

        public double VariableAlaeForLargeLoss { get; set; }

        public double FlatAlaeForLargeLoss { get; set; }

        public double Truncation { get; set; }

        public double AverageSeverityForSmallLoss { get; set; }

        public double ProbabilityLessThanTruncation { get; set; }

        public double Scale { get; set; }

        public double Shape { get; set; }
    }
}