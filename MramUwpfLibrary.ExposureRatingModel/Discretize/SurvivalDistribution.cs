using System.Collections.Generic;
using System.Linq;
using MramUwpfLibrary.Common.Layers;

namespace MramUwpfLibrary.ExposureRatingModel.Discretize
{
    public class SurvivalDistribution
    {
        public IList<SurvivalDistributionItem> Items { get; set; }

        public SurvivalDistribution()
        {
            Items = new List<SurvivalDistributionItem>();
        }

        public IDiscretization MapToDiscretization()
        {
            var discretization = new Discretization();
            for (var i = 0; i < Items.Count - 1; i++)
            {
                discretization.Add(new DiscretizationItem
                {
                    Loss = Items[i].Loss,
                    Probability = Items[i].Probability - Items[i + 1].Probability
                });
            }
            discretization.Add(new DiscretizationItem { Loss = Items.Last().Loss, Probability = Items.Last().Probability });
            return discretization;
        }

        public static double GetFrequency(ITower tower, IDiscretization discretization)
        {
            var gX = GetSurvivalDistribution(discretization);

            var firstTowerItem = tower.First();
            var firstTowerItemTop = firstTowerItem.Value.Attachment + firstTowerItem.Value.Limit;
            var freqTmp = gX.Items
                .Where(item => item.Loss > firstTowerItem.Value.Attachment && item.Loss <= firstTowerItemTop)
                .Sum(item => item.Probability);
            return firstTowerItem.Value.Loss / discretization[1].Loss / freqTmp;
        }

        public static SurvivalDistribution GetSurvivalDistribution(IDiscretization discretization)
        {
            var survivalDistribution = new SurvivalDistribution();
            survivalDistribution.Items.Add(new SurvivalDistributionItem { Loss = 0d, Probability = 1d });
            for (var i = 1; i < discretization.Count; i++)
            {
                survivalDistribution.Items.Add(new SurvivalDistributionItem
                {
                    Loss = discretization[i].Loss,
                    Probability = survivalDistribution.Items[i - 1].Probability - discretization[i - 1].Probability
                });
            }
            return survivalDistribution;
        }

        public double SumBand(double limit, double attachment)
        {
            var bottom = attachment;
            var top = limit + attachment;
            return Items.Where(item => item.Loss > bottom && item.Loss <= top).Sum(item => item.Probability);
        }

    }
    public class SurvivalDistributionItem
    {
        public double Loss { get; set; }
        public double Probability { get; set; }
    }

}