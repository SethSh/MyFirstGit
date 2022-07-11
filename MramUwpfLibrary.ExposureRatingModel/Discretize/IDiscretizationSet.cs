using System.Collections.Generic;
using System.Linq;
using MramUwpfLibrary.Common.Extensions;

namespace MramUwpfLibrary.ExposureRatingModel.Discretize
{
    public interface IDiscretizationSet
    {
        IDiscretization Discretization { get; set; }
        double Weight { get; set; }
    }

    public class DiscretizationSet : IDiscretizationSet
    {
        public IDiscretization Discretization { get; set; }
        public double Weight { get; set; }

        public static IDiscretization Aggregate(IList<IDiscretizationSet> discretizationSets)
        {
            var aggregatedDiscretization = new Discretization();
            var firstDiscretizationSet = discretizationSets.First();
            foreach (var item in firstDiscretizationSet.Discretization)
            {
                aggregatedDiscretization.Add(new DiscretizationItem { Loss = item.Loss });
            }
            

            var totalWeight = discretizationSets.Sum(distSet => distSet.Weight);
            foreach (var discretizationSet in discretizationSets)
            {
                var relativeWeight = discretizationSet.Weight.DivideByWithTrap(totalWeight);
                var counter = 0;
                foreach (var discretizedItem in discretizationSet.Discretization)
                {
                    aggregatedDiscretization[counter++].Probability += discretizedItem.Probability * relativeWeight;
                }
            }

            return aggregatedDiscretization;
        }
    }
}