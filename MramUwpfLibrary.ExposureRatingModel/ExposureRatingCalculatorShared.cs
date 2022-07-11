using System.Collections.Generic;
using System.Linq;
using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.Common.Extensions;

namespace MramUwpfLibrary.ExposureRatingModel
{
    internal class ExposureRatingCalculatorShared
    {
        internal static GrossUpLossRatioFactors AggregateLossRatioSets(IList<LossRatioResultSet> inputs)
        {
            var totalSubjBase = inputs.Sum(x => x.SubjBase);
            var totalLimitedOverUnlimitedFactor = inputs.Sum(x => x.SubjBase * x.GrossUpLossRatio.LimitedToUnlimited.Custom);
            var totalBenchmarkLimitedOverUnlimitedFactor = inputs.Sum(x => x.SubjBase * x.GrossUpLossRatio.LimitedToUnlimited.Benchmark);

            var blendedLimitedOverUnlimitedFactor = totalSubjBase > 0 ? totalLimitedOverUnlimitedFactor / totalSubjBase : 1;
            var blendedBenchmarkLimitedOverUnlimitedFactor = totalSubjBase > 0 ? totalBenchmarkLimitedOverUnlimitedFactor / totalSubjBase : 1;

            var totalLimitedLossPlusAlae = inputs.Sum(x => x.SubjBase * x.LimitedLossRatio);

            var totalAlae = 0d;
            var totalBenchmarkAlae = 0d;
            var totalLoss = 0d;
            var totalBenchmarkLoss = 0d;

            foreach (var input in inputs)
            {
                var restatedLossRatio = input.GrossUpLossRatio;
                var limitedLossPlusAlae = input.SubjBase * input.LimitedLossRatio;

                var unlimitedLossPlusAlae = limitedLossPlusAlae * blendedLimitedOverUnlimitedFactor;
                var benchmarkUnlimitedLossPlusAlae = limitedLossPlusAlae * blendedBenchmarkLimitedOverUnlimitedFactor;

                var alae = unlimitedLossPlusAlae * restatedLossRatio.AlaeToLoss.Custom / (1 + restatedLossRatio.AlaeToLoss.Custom);
                var benchmarkAlae = benchmarkUnlimitedLossPlusAlae * restatedLossRatio.AlaeToLoss.Benchmark /
                                    (1 + restatedLossRatio.AlaeToLoss.Benchmark);

                var loss = unlimitedLossPlusAlae - alae;
                var benchmarkLoss = benchmarkUnlimitedLossPlusAlae - benchmarkAlae;

                totalAlae += alae;
                totalBenchmarkAlae += benchmarkAlae;
                totalLoss += loss;
                totalBenchmarkLoss += benchmarkLoss;
            }

            var totalUnlimitedLossPlusAlae = totalLimitedLossPlusAlae.DivideByWithTrap(blendedLimitedOverUnlimitedFactor);
            var totalBenchmarkUnlimitedLossPlusAlae = totalLimitedLossPlusAlae.DivideByWithTrap(blendedBenchmarkLimitedOverUnlimitedFactor);

            var limitedToUnlimited = new LimitedToUnlimitedFactors
            {
                Custom = blendedLimitedOverUnlimitedFactor,
                Benchmark = blendedBenchmarkLimitedOverUnlimitedFactor,
                CustomUnlimitedLossRatio = totalUnlimitedLossPlusAlae.DivideByWithTrap(totalSubjBase, double.NaN),
                BenchmarkUnlimitedLossRatio = totalBenchmarkUnlimitedLossPlusAlae.DivideByWithTrap(totalSubjBase, double.NaN)
            };
            limitedToUnlimited.CheckForNan();

            var alaeToLoss = new AlaeToLossFactors
            {
                Custom = totalAlae.DivideByWithTrap(totalLoss),
                Benchmark = totalBenchmarkAlae.DivideByWithTrap(totalBenchmarkLoss),
            };
            alaeToLoss.CheckForNan();

            return new GrossUpLossRatioFactors(limitedToUnlimited, alaeToLoss);
        }

        internal static GrossUpLossRatioFactors AggregateLossRatioSets(bool useAlternative, IList<LossRatioResultSet> lossRatioSets)
        {
            return !useAlternative ? AggregateLossRatioSets(lossRatioSets) : null;
        }


        internal static double GetUnlimitedOverLimitedFactor(ReinsuranceAlaeTreatmentType reinsuranceAlaeTreatment, GrossUpLossRatioFactors grossUpLossRatio, double limitedLoss)
        {
            var limitedToUnlimited = grossUpLossRatio.LimitedToUnlimited;
            var alaeToLoss = grossUpLossRatio.AlaeToLoss;

            var unlimitedOverLimitedFactor = 1 / limitedToUnlimited.Custom;
            var benchmarkUnlimitedOverLimitedFactor = 1 / limitedToUnlimited.Benchmark;

            var unlimitedLoss = limitedLoss * unlimitedOverLimitedFactor;
            var benchmarkUnlimitedLoss = limitedLoss * benchmarkUnlimitedOverLimitedFactor;

            var alae = unlimitedLoss * alaeToLoss.Custom / (1 + alaeToLoss.Custom);
            var benchmarkAlae = benchmarkUnlimitedLoss * alaeToLoss.Benchmark / (1 + alaeToLoss.Benchmark);

            var alaeToLossFactor = alae / (unlimitedLoss - alae);
            var benchmarkAlaeToLossFactor = benchmarkAlae / (benchmarkUnlimitedLoss - benchmarkAlae);

            if (reinsuranceAlaeTreatment == ReinsuranceAlaeTreatmentType.Excluded)
            {
                unlimitedOverLimitedFactor /= 1 + alaeToLossFactor;
                // ReSharper disable once RedundantAssignment
                // this is from MATLAB - not doing anything
                benchmarkUnlimitedOverLimitedFactor /= 1 + benchmarkAlaeToLossFactor;
            }

            return unlimitedOverLimitedFactor;
        }
    }
}