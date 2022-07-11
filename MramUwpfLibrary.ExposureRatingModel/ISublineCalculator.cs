using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.Common.Extensions;
using MramUwpfLibrary.ExposureRatingModel.Casualty.CurveInputHelpers;
using MramUwpfLibrary.ExposureRatingModel.Casualty.Curves;
using MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials;
using MramUwpfLibrary.ExposureRatingModel.Input;

[assembly: InternalsVisibleTo("MramUwpfLibrary.Tests")]

namespace MramUwpfLibrary.ExposureRatingModel
{
    internal interface ISublineCalculator
    {
        ExposureRatingResultItem Calculate(GrossUpLossRatioFactors grossUpLossRatioFactors);
        GrossUpLossRatioFactors GrossUpLossRatio();
    }

    internal abstract class BaseSublineCalculator : ISublineCalculator
    {
        private readonly ReinsuranceParameters _reinsuranceParameters;
        private readonly ISublineExposureRatingInput _sublineInput;
        private readonly PolicyAlaeTreatmentType _policyAlaeTreatment;

        protected BaseSublineCalculator(ReinsuranceParameters reinsuranceParameters, 
            ISublineExposureRatingInput sublineInput, 
            PolicyAlaeTreatmentType policyAlaeTreatment)
        {
            _reinsuranceParameters = reinsuranceParameters;
            _sublineInput = sublineInput;
            _policyAlaeTreatment = policyAlaeTreatment;
        }

        public abstract ExposureRatingResultItem Calculate(GrossUpLossRatioFactors grossUpLossRatioFactors);

        public abstract GrossUpLossRatioFactors GrossUpLossRatio();

        internal ExposureRatingResultItem Calculate()
        {
            var lossRatio = GrossUpLossRatio();
            return Calculate(lossRatio);
        }


        internal FactorPair AggregateNonZeroFactorPair(List<FactorPairAndWeight> pairAndWeights)
        {
            // for layer with respect to ground, we need to discard policies that don't expose the limited layer
            // throwing out zero value factors and then normalizing the remaining ones prior to summing
            var factor = GetFactor(pairAndWeights);
            var benchmarkFactor = GetBenchmarkFactor(pairAndWeights);

            return new FactorPair
            {
                Custom = factor,
                Benchmark = benchmarkFactor
            };
        }
        
        internal double GetLayerLimitedExpectedValue(double policyLimit, double policySir, 
            CurveInputs curveInputs, ICalculator curveCalculator, Parameters parameters)
        {
            return curveCalculator.GetLayerLimitedExpectedValue(
                curveInputs.TopLimit,
                curveInputs.BottomLimit,
                policyLimit,
                policySir,
                curveInputs.AlaeAdjustmentFactor,
                curveInputs.ReinsurancePerspective,
                parameters);
        }

        protected AlaeToLossFactors CalculateAlaeToLoss(double policyLimit,
            double? policySir,
            ICalculator selectedAlaeCalculator,
            Parameters parameters,
            bool forceWithinLimits = false)
        {
            var policySirAsDouble = policySir ?? 0;
            var policyAlaeTreatment = forceWithinLimits ? PolicyAlaeTreatmentType.WithinLimit : _policyAlaeTreatment;

            if (policyAlaeTreatment != PolicyAlaeTreatmentType.InAdditionToLimit)
            {
                return new AlaeToLossFactors
                {
                    Custom = 0,
                    Benchmark = 0
                };
            }

            var selectedAlaeHelper = new SelectedAlaeHelper(_sublineInput, policyLimit, policySirAsDouble);

            var policyLev = GetLayerLimitedExpectedValue(policyLimit,
                policySirAsDouble, selectedAlaeHelper.GetNumeratorCurveInputs(), selectedAlaeCalculator, parameters);

            var policyLevUsingLossOnly = GetLayerLimitedExpectedValue(policyLimit,
                policySirAsDouble, selectedAlaeHelper.GetDenominatorCurveInputs(), selectedAlaeCalculator, parameters);

            var benchmarkPolicyLev = GetLayerLimitedExpectedValue(policyLimit,
                policySirAsDouble, selectedAlaeHelper.GetNumeratorCurveBenchmarkInputs(), selectedAlaeCalculator, parameters);

            var alaeRatio = policyLev.DivideByWithTrap(policyLevUsingLossOnly, 1) - 1;
            var benchmarkAlaeRatio = benchmarkPolicyLev.DivideByWithTrap(policyLevUsingLossOnly, 1) - 1;

            return new AlaeToLossFactors
            {
                Custom = alaeRatio,
                Benchmark = benchmarkAlaeRatio
            };
        }

        protected LimitedToUnlimitedFactors CalculateLimitedToUnlimited(double policyLimit,
            double? policySir,
            ICalculator limitedLossRatioCurveCalculator,
            ICalculator unlimitedLossRatioCurveCalculator,
            Parameters parameters,
            bool forceWithinLimits = false)
        {
            var policySirAsDouble = policySir ?? 0;
            var policyAlaeTreatment = forceWithinLimits ? PolicyAlaeTreatmentType.WithinLimit : _policyAlaeTreatment;
            var lossRatioHelper = new LossRatioHelper(policyAlaeTreatment, _sublineInput, policyLimit, policySirAsDouble, _reinsuranceParameters.PerspectiveHandler);

            var limitedLossRatioLev = GetLayerLimitedExpectedValue(policyLimit, policySirAsDouble,
                lossRatioHelper.GetNumeratorCurveInputs(), limitedLossRatioCurveCalculator, parameters);

            var unlimitedLossRatioLev = GetLayerLimitedExpectedValue(policyLimit, policySirAsDouble,
                lossRatioHelper.GetDenominatorCurveInputs(), unlimitedLossRatioCurveCalculator, parameters);


            var benchmarkLimitedLossRatioLev = GetLayerLimitedExpectedValue(policyLimit, policySirAsDouble,
                lossRatioHelper.GetNumeratorCurveBenchmarkInputs(), limitedLossRatioCurveCalculator, parameters);

            var benchmarkUnlimitedLossRatioLev = GetLayerLimitedExpectedValue(policyLimit, policySirAsDouble,
                lossRatioHelper.GetDenominatorCurveBenchmarkInputs(), unlimitedLossRatioCurveCalculator, parameters);


            var limitedToUnlimited = limitedLossRatioLev.DivideByWithTrap(unlimitedLossRatioLev);
            var benchmarkLimitedToUnlimited = benchmarkLimitedLossRatioLev.DivideByWithTrap(benchmarkUnlimitedLossRatioLev);

            return new LimitedToUnlimitedFactors
            {
                Custom = limitedToUnlimited,
                Benchmark = benchmarkLimitedToUnlimited,
            };
        }
        
        private static double GetFactor(List<FactorPairAndWeight> pairAndWeights)
        {
            if (pairAndWeights.All(pw => pw.Custom.IsEqualToZero())) throw new InvalidDataException("Can't gross-up limited loss ratio: No exposed subject policies"); 
            
            var totalWeight = pairAndWeights.Sum(pw => pw.Weight);
            double factor;
            if (pairAndWeights.Any(pw => pw.Custom.IsEqualToZero()))
            {
                var nonZeroWeightTotal = pairAndWeights.Where(pw => pw.Custom > 0).Sum(pw => pw.Weight);
                var grossUp = totalWeight / nonZeroWeightTotal;
                factor = pairAndWeights.Where(pw => pw.Custom > 0).Sum(pw => pw.Custom * pw.Weight * grossUp);
            }
            else
            {
                factor = pairAndWeights.Sum(pw => pw.Custom * pw.Weight);
            }

            return factor;
        }

        private static double GetBenchmarkFactor(List<FactorPairAndWeight> pairAndWeights)
        {
            if (pairAndWeights.All(pw => pw.Benchmark.IsEqualToZero())) throw new InvalidDataException("Can't gross-up benchmark limited loss ratio: No exposed subject policies"); 
            
            var totalWeight = pairAndWeights.Sum(pw => pw.Weight);
            double benchmarkFactor;
            if (pairAndWeights.Any(pw => pw.Benchmark.IsEqualToZero()))
            {
                var nonZeroWeightTotal = pairAndWeights.Where(pw => pw.Benchmark > 0).Sum(pw => pw.Weight);
                var grossUp = totalWeight / nonZeroWeightTotal;
                benchmarkFactor = pairAndWeights.Where(pw => pw.Benchmark > 0).Sum(pw => pw.Benchmark * pw.Weight * grossUp);
            }
            else
            {
                benchmarkFactor = pairAndWeights.Sum(pw => pw.Benchmark * pw.Weight);
            }

            return benchmarkFactor;
        }
    }

    internal struct LayerResultItem
    {
        internal double LayerLevOverPolicyLev { get; set; }
        internal double Severity { get; set; }
        internal double Weight { get; set; }
    }

    internal class FactorPair
    {
        internal double Custom { get; set; }
        internal double Benchmark { get; set; }

        internal static FactorPair Add(FactorPair factorPair, FactorPair otherFactorPair)
        {
            return new FactorPair
            {
                Custom = factorPair.Custom + otherFactorPair.Custom,
                Benchmark = factorPair.Benchmark + otherFactorPair.Benchmark
            };
        }
    }

    internal class FactorPairAndWeight : FactorPair
    {
        public double Weight { get; set; }
    }

    internal class LossRatioResultSet
    {
        public string Id { get; set; }
        public double SubjBase { get; set; }
        public GrossUpLossRatioFactors GrossUpLossRatio { get; set; }
        public double LimitedLossRatio { get; set; }
        public static LossRatioResultSet CopyValues(string sublineId, LossRatioResultSet lossRatioResultSet)
        {
            return new LossRatioResultSet
            {
                Id = sublineId,
                GrossUpLossRatio = new GrossUpLossRatioFactors
                {
                    AlaeToLoss = lossRatioResultSet.GrossUpLossRatio.AlaeToLoss,
                    LimitedToUnlimited = lossRatioResultSet.GrossUpLossRatio.LimitedToUnlimited,
                },
                LimitedLossRatio = lossRatioResultSet.LimitedLossRatio,
                SubjBase = lossRatioResultSet.SubjBase
            };
        }
    }
    
    internal class GrossUpLossRatioFactors
    {
        public LimitedToUnlimitedFactors LimitedToUnlimited { get; set; }
        public AlaeToLossFactors AlaeToLoss { get; set; }

        public GrossUpLossRatioFactors(LimitedToUnlimitedFactors limitedToUnlimited, AlaeToLossFactors alaeToLoss)
        {
            LimitedToUnlimited = limitedToUnlimited;
            AlaeToLoss = alaeToLoss;
        }

        public GrossUpLossRatioFactors()
        {
            LimitedToUnlimited = new LimitedToUnlimitedFactors
            {
                Custom = 1,
                Benchmark = 1,
                CustomUnlimitedLossRatio = 1,
                BenchmarkUnlimitedLossRatio = 1,
            };

            AlaeToLoss = new AlaeToLossFactors
            {
                Custom = 0,
                Benchmark = 0,
            };
        }
    }

    internal class LimitedToUnlimitedFactors
    {
        internal double Custom { get; set; }
        internal double Benchmark { get; set; }
        
        public double CustomUnlimitedLossRatio { get; set; }
        public double BenchmarkUnlimitedLossRatio { get; set; }
        
        public void CheckForNan()
        {
            var nanList = new List<string>();
            if (double.IsNaN(Custom)) nanList.Add("Limit Over Unlimited Factor");
            if (double.IsNaN(Benchmark)) nanList.Add("Benchmark Limit Over Unlimited Factor");
            if (double.IsNaN(CustomUnlimitedLossRatio)) nanList.Add("Unlimited Loss Ratio");
            if (double.IsNaN(BenchmarkUnlimitedLossRatio)) nanList.Add("Benchmark Unlimited Loss Ratio");

            if (nanList.Count > 0) throw new ArgumentException("Restate Loss Ratio As Unlimited failed for " + string.Join(", ", nanList));
        }
    }
    internal class AlaeToLossFactors
    {
        internal double Custom { get; set; }
        internal double Benchmark { get; set; }
        public void CheckForNan()
        {
            var nanList = new List<string>();
            if (double.IsNaN(Custom)) nanList.Add("ALAE To Loss");
            if (double.IsNaN(Benchmark)) nanList.Add("Benchmark ALAE To Loss");
        
            if (nanList.Count > 0) throw new ArgumentException("ALAE to Loss failed for " + string.Join(", ", nanList));
        }
    }
}
