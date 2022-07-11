using System.Collections.Generic;
using MramUwpfLibrary.Common.Extensions;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.ResultAggregators
{
    internal static class LayerResultsMapper
    {
        public static ExposureRatingResultItem Map(IEnumerable<LayerResultItem> layerResults,
            double subjectBase,
            GrossUpLossRatioFactors grossUpLossRatio,
            double limitedLoss, 
            double unlimitedOverLimitedFactor,
            double? share = new double?())
        {
            var frequency = 0d;
            var layerLossCost = 0d;
            foreach (var result in layerResults)
            {
                var limitedLayerLoss = limitedLoss * result.LayerLevOverPolicyLev;
                var thisLayerLossCost = share * limitedLayerLoss * unlimitedOverLimitedFactor ?? limitedLayerLoss * unlimitedOverLimitedFactor;
                var weight = result.Weight;

                layerLossCost += thisLayerLossCost * weight;

                var thisFrequency = thisLayerLossCost.DivideByWithTrap(result.Severity);
                frequency += thisFrequency * weight;
            }

            var layerLossCostPercent = layerLossCost / subjectBase;
            var averageSeverity = layerLossCost.DivideByWithTrap(frequency);

            return new ExposureRatingResultItem
            {
                AlaeToLoss = grossUpLossRatio.AlaeToLoss.Custom,
                BenchmarkAlaeToLoss = grossUpLossRatio.AlaeToLoss.Benchmark,
                UnlimitedLossPlusAlaeRatio = grossUpLossRatio.LimitedToUnlimited.CustomUnlimitedLossRatio,
                BenchmarkUnlimitedLossPlusAlaeRatio = grossUpLossRatio.LimitedToUnlimited.BenchmarkUnlimitedLossRatio,

                Severity = averageSeverity,
                LayerLossCostPercent = layerLossCostPercent,
                LayerLossCostAmount = layerLossCost,
                Frequency = frequency
            };
        }
    }
}