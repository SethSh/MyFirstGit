using System.Linq;
using MramUwpfLibrary.ExposureRatingModel.Discretize;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.TruncatedParetos
{
    public interface IParameters
    {
        double AverageSeverityForSmallLoss { get; set; }
        double FlatAlaeForLargeLoss { get; set; }
        double ProbabilityLessThanTruncation { get; set; }
        double Scale { get; set; }
        double Shape { get; set; }
        double Truncation { get; set; }
        double VariableAlaeForLargeLoss { get; set; }
        double VariableAlaeForSmallLoss { get; set; }
    }

    internal interface IParametersProcessor
    {
        IDiscretization Discretize(IDiscretizeInput discretizeInput);
    }

    public class ParametersProcessor : IParametersProcessor
    {
        private readonly IParameters _parameters;
        
        public ParametersProcessor(IParameters parameters)
        {
            _parameters = parameters;
        }

        public IDiscretization Discretize(IDiscretizeInput discretizeInput)
        {
            var cumulativeDiscretization = new Discretization();

            for (var index = 0; index < discretizeInput.BucketCount; index++)
            {
                var loss = index * discretizeInput.BucketWidth;
                cumulativeDiscretization.Add(
                    new DiscretizationItem
                    {
                        Loss = loss,
                        Probability = BaseCurve.GetCdfIgnoringAlae(loss, _parameters)
                    });
            }

            cumulativeDiscretization.Last().Probability = 1d;
            
            var discretization = cumulativeDiscretization.MapCumulativeToIncremental();
            return discretization;
        }
    }
}