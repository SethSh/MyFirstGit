using System.Collections.Generic;
using System.Linq;
using MramUwpfLibrary.Common.Extensions;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.ResultAggregators
{
    internal static class ExposureRatingResultAggregator
    {
        public static ExposureRatingResultItem Aggregate(List<ExposureRatingResultItem> exposureRatingResultItems)
        {
            var weightTotal = exposureRatingResultItems.Sum(result => result.Weight);

            var layerLossCost = 0d;
            var layerLossCostPercent = 0d;
            var frequency = 0d;
            var alae = 0d;
            var benchmarkAlae = 0d;
            var unlimitedLossPlusAlaeRatio = 0d;
            var benchmarkUnlimitedLossPlusAlaeRatio = 0d;
            var weight = 0d;

            foreach (var row in exposureRatingResultItems)
            {
                var normalizedWeight = row.Weight.DivideByWithTrap(weightTotal);

                layerLossCost += row.LayerLossCostAmount * normalizedWeight;
                layerLossCostPercent += row.LayerLossCostPercent * normalizedWeight;
                frequency += row.Frequency * normalizedWeight;
                alae += row.LayerLossCostAmount * row.AlaeToLoss * normalizedWeight;
                benchmarkAlae += row.LayerLossCostAmount * row.BenchmarkAlaeToLoss * normalizedWeight;
                unlimitedLossPlusAlaeRatio += row.UnlimitedLossPlusAlaeRatio * normalizedWeight;
                benchmarkUnlimitedLossPlusAlaeRatio += row.BenchmarkUnlimitedLossPlusAlaeRatio * normalizedWeight;
                weight += normalizedWeight;
            }

            var alaeToLoss = alae.DivideByWithTrap(layerLossCost);
            var benchmarkAlaeToLoss = benchmarkAlae.DivideByWithTrap(layerLossCost);

            return new ExposureRatingResultItem
            {
                AlaeToLoss = alaeToLoss,
                BenchmarkAlaeToLoss = benchmarkAlaeToLoss,
                Severity = layerLossCost.DivideByWithTrap(frequency),
                UnlimitedLossPlusAlaeRatio = unlimitedLossPlusAlaeRatio,
                BenchmarkUnlimitedLossPlusAlaeRatio = benchmarkUnlimitedLossPlusAlaeRatio,
                LayerLossCostPercent = layerLossCostPercent,
                LayerLossCostAmount = layerLossCost,
                Frequency = frequency,
                Weight = weight
            };
        }
    }
}

