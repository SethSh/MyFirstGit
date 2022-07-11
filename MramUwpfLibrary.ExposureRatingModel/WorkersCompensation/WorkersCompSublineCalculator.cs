using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.Common.Extensions;
using MramUwpfLibrary.Common.Policies;
using MramUwpfLibrary.ExposureRatingModel.Casualty;
using MramUwpfLibrary.ExposureRatingModel.Casualty.CurveInputHelpers;
using MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials;
using MramUwpfLibrary.ExposureRatingModel.Casualty.ResultAggregators;
using MramUwpfLibrary.ExposureRatingModel.Input;

namespace MramUwpfLibrary.ExposureRatingModel.WorkersCompensation
{
    internal class WorkersCompSublineCalculator : BaseSublineCalculator
    {
        private readonly ReinsuranceParameters _reinsuranceParameters;
        private readonly PolicyAlaeTreatmentType _policyAlaeTreatment;
        private readonly ISublineExposureRatingInput _sublineInput;
        private readonly IList<MixedExponentialCurve> _curves;


        public WorkersCompSublineCalculator(ReinsuranceParameters reinsuranceParameters,
            ISublineExposureRatingInput sublineInput,
            PolicyAlaeTreatmentType policyAlaeTreatment,
            IList<MixedExponentialCurve> mixedExponentials) : base(reinsuranceParameters, sublineInput, policyAlaeTreatment)
        {
            _reinsuranceParameters = reinsuranceParameters;
            _policyAlaeTreatment = policyAlaeTreatment;
            _sublineInput = sublineInput;
            _curves = mixedExponentials;
        }

        public override ExposureRatingResultItem Calculate(GrossUpLossRatioFactors grossUpLossRatioFactors)
        {
            var exposureRatingResultItems = new List<ExposureRatingResultItem>();
            var limitedLoss = _sublineInput.AllocatedExposureAmount * _sublineInput.LimitedLossRatio;
            var unlimitedOverLimitedFactor =
                ExposureRatingCalculatorShared.GetUnlimitedOverLimitedFactor(_reinsuranceParameters.AlaeTreatment, grossUpLossRatioFactors, limitedLoss);

            var tasks = new List<Task<ExposureRatingResultItem>>();

            var layerCalculator = CalculatorFactory.Create(_reinsuranceParameters.AlaeTreatment, _policyAlaeTreatment);
            var reinsuranceAlaeTreatment = _policyAlaeTreatment.MapToReinsuranceAlaeTreatmentTypeForLayer(_reinsuranceParameters.AlaeTreatment);
            var policyCalculator = CalculatorFactory.Create(reinsuranceAlaeTreatment, _policyAlaeTreatment);

            var withinLimitsLayerCalculator = CalculatorFactory.Create(_reinsuranceParameters.AlaeTreatment, PolicyAlaeTreatmentType.WithinLimit);
            var withinLimitsReinsuranceAlaeTreatment = PolicyAlaeTreatmentType.WithinLimit.MapToReinsuranceAlaeTreatmentTypeForLayer(_reinsuranceParameters.AlaeTreatment);
            var withinLimitsPolicyCalculator = CalculatorFactory.Create(withinLimitsReinsuranceAlaeTreatment, PolicyAlaeTreatmentType.WithinLimit);

            foreach (var curve in _curves)
            {
                var task = new Task<ExposureRatingResultItem>(() =>
                {
                    var forceWithinLimits = curve.ForceWithinLimits;
                    var policyLayerResultsItems = CalculateReinsuranceLayerResults(curve,
                        forceWithinLimits ? withinLimitsLayerCalculator.CreateNew() : layerCalculator.CreateNew(),
                        forceWithinLimits ? withinLimitsPolicyCalculator.CreateNew() : policyCalculator.CreateNew());
                    
                    var layerResult = LayerResultsMapper.Map(policyLayerResultsItems,  _sublineInput.AllocatedExposureAmount, 
                        grossUpLossRatioFactors,  limitedLoss, unlimitedOverLimitedFactor);

                    layerResult.SublineId = _sublineInput.Id;

                    return new ExposureRatingResultItem
                    {
                        AlaeToLoss = layerResult.AlaeToLoss,
                        Severity = layerResult.Severity,
                        BenchmarkAlaeToLoss = layerResult.BenchmarkAlaeToLoss,
                        BenchmarkUnlimitedLossPlusAlaeRatio = layerResult.BenchmarkUnlimitedLossPlusAlaeRatio,
                        Frequency = layerResult.Frequency,
                        SublineId = layerResult.SublineId,
                        LayerLossCostAmount = layerResult.LayerLossCostAmount,
                        LayerLossCostPercent = layerResult.LayerLossCostPercent,
                        UnlimitedLossPlusAlaeRatio = layerResult.UnlimitedLossPlusAlaeRatio,
                        Weight = curve.Weight * _sublineInput.AllocatedExposureAmount
                    };
                });
                task.Start();
                tasks.Add(task);
            }


            // ReSharper disable once CoVariantArrayConversion
            Task.WaitAll(tasks.ToArray());
            exposureRatingResultItems.AddRange(tasks.Select(t => t.Result));

            var overallLayerResult = ExposureRatingResultAggregator.Aggregate(exposureRatingResultItems);
            overallLayerResult.SublineId = _sublineInput.Id;

            overallLayerResult.CheckForNan();
            return overallLayerResult;
        }

        public override GrossUpLossRatioFactors GrossUpLossRatio()
        {
            var limitedLossRatioReinsuranceAlae = _sublineInput.LossRatioAlaeTreatment.MapToReinsuranceAlaeTreatmentType();
            var limitedLossRatioCurveCalculator = CalculatorFactory.Create(limitedLossRatioReinsuranceAlae, _policyAlaeTreatment);
            var withinLimitsLimitedLossRatioCurveCalculator = CalculatorFactory.Create(limitedLossRatioReinsuranceAlae, PolicyAlaeTreatmentType.WithinLimit);

            var unlimitedLossRatioReinsuranceAlae = _policyAlaeTreatment.MapToReinsuranceAlaeTreatmentTypeForUnlimitedLossRatio();
            var unlimitedLossRatioCurveCalculator = CalculatorFactory.Create(unlimitedLossRatioReinsuranceAlae, _policyAlaeTreatment);

            var withinLimitsUnlimitedLossRatioReinsuranceAlae = PolicyAlaeTreatmentType.WithinLimit.MapToReinsuranceAlaeTreatmentTypeForUnlimitedLossRatio();
            var withinLimitsUnlimitedLossRatioCurveCalculator = CalculatorFactory.Create(withinLimitsUnlimitedLossRatioReinsuranceAlae, PolicyAlaeTreatmentType.WithinLimit);


            var proRata = ReinsuranceAlaeTreatmentType.ProRata;
            var inAdditionToLimit = PolicyAlaeTreatmentType.InAdditionToLimit;
            var selectedAlaeCalculator = CalculatorFactory.Create(proRata, inAdditionToLimit);


            var limitedToUnlimitedFactors = CalculateLimitedToUnlimitedLossFactors(
                _curves,
                limitedLossRatioCurveCalculator.CreateNew,
                unlimitedLossRatioCurveCalculator.CreateNew,
                withinLimitsLimitedLossRatioCurveCalculator.CreateNew,
                withinLimitsUnlimitedLossRatioCurveCalculator.CreateNew);


            var alaeToLossFactors = CalculateAlaeToLossFactors(
                _curves,
                limitedToUnlimitedFactors,
                selectedAlaeCalculator.CreateNew);

            var limitedToUnlimited = new LimitedToUnlimitedFactors
            {
                Custom = limitedToUnlimitedFactors.Custom,
                Benchmark = limitedToUnlimitedFactors.Benchmark,
                CustomUnlimitedLossRatio = _sublineInput.LimitedLossRatio / limitedToUnlimitedFactors.Custom,
                BenchmarkUnlimitedLossRatio = _sublineInput.LimitedLossRatio / limitedToUnlimitedFactors.Benchmark
            };

            var alaeToLoss = new AlaeToLossFactors
            {
                Custom = alaeToLossFactors.Custom,
                Benchmark = alaeToLossFactors.Benchmark,
            };

            return new GrossUpLossRatioFactors(limitedToUnlimited, alaeToLoss);
        }

        private IEnumerable<LayerResultItem> CalculateReinsuranceLayerResults(MixedExponentialCurve meCurve,
            ICalculator reinsuranceCurveCalculator,
            ICalculator policyCurveCalculator)
        {
            return meCurve.PolicySet
                .Select(policyAndNormalizedWeight =>
                    CalculateReinsuranceLayerResult(meCurve.CurveParameters,
                        policyAndNormalizedWeight, reinsuranceCurveCalculator, policyCurveCalculator));
        }

        private LayerResultItem CalculateReinsuranceLayerResult(
            Parameters parameters, IPolicy policy,
            ICalculator layerCurveCalculator,
            ICalculator policyCurveCalculator)
        {
            var policySir = policy.Sir ?? 0;

            IHelper helper = new ReinsuranceLayerHelper(_reinsuranceParameters,
                _policyAlaeTreatment,
                _sublineInput,
                policy.Limit, policySir);

            var layerParameters = helper.GetNumeratorCurveInputs();
            var policyParameters = helper.GetDenominatorCurveInputs();

            var layerLev = GetLayerLimitedExpectedValue(policy.Limit,
                policySir,
                layerParameters,
                layerCurveCalculator,
                parameters);

            var probLessThanLayerBottom = layerCurveCalculator.GetProbabilityLessThanLimit(
                layerParameters.BottomLimit,
                policy.Limit,
                policySir,
                layerParameters.AlaeAdjustmentFactor,
                layerParameters.ReinsurancePerspective,
                parameters);

            var policyLev = GetLayerLimitedExpectedValue(policy.Limit,
                policySir,
                policyParameters,
                policyCurveCalculator,
                parameters);

            var layerLevOverPolicyLev = layerLev.DivideByWithTrap(policyLev);
            var severity = probLessThanLayerBottom < 1
                ? layerLev / (1 - probLessThanLayerBottom)
                : 0;

            return new LayerResultItem
            {
                LayerLevOverPolicyLev = layerLevOverPolicyLev,
                Severity = severity,
                Weight = policy.Amount
            };
        }

        private FactorPair CalculateLimitedToUnlimitedLossFactors(IEnumerable<MixedExponentialCurve> mixedExponentialCurves,
            Func<ICalculator> limitedLossRatioCurveCalculator,
            Func<ICalculator> unlimitedLossRatioCurveCalculator,
            Func<ICalculator> withinLimitsLimitedLossRatioCurveCalculator,
            Func<ICalculator> withinLimitsUnlimitedLossRatioCurveCalculator)
        {
            if (mixedExponentialCurves == null)
            {
                return new FactorPair();
            }

            var tasks = new List<Task<FactorPair>>();
            foreach (var curve in mixedExponentialCurves)
            {
                var task = new Task<FactorPair>(() =>
                {
                    var forceWithinLimits = curve.ForceWithinLimits;
                    var pairAndWeights = new List<FactorPairAndWeight>();
                    var parameters = curve.CurveParameters;

                    foreach (var policy in curve.PolicySet)
                    {
                        var policyLimit = policy.Limit;
                        var policySir = policy.Sir;
                        var policyWeight = policy.Amount;

                        if (policyWeight.IsEqualToZero()) continue;

                        LimitedToUnlimitedFactors factors;
                        if (forceWithinLimits)
                        {
                            factors = CalculateLimitedToUnlimited(policyLimit,
                                policySir,
                                withinLimitsLimitedLossRatioCurveCalculator(),
                                withinLimitsUnlimitedLossRatioCurveCalculator(),
                                parameters, true);
                        }
                        else
                        {
                            factors = CalculateLimitedToUnlimited(policyLimit,
                                policySir,
                                limitedLossRatioCurveCalculator(),
                                unlimitedLossRatioCurveCalculator(),
                                parameters);
                        }
                        

                        var weight = policyWeight * curve.Weight;

                        pairAndWeights.Add(new FactorPairAndWeight
                        {
                            Custom = factors.Custom,
                            Benchmark = factors.Benchmark,
                            Weight = weight
                        });
                    }

                    return AggregateNonZeroFactorPair(pairAndWeights);
                });


                task.Start();
                tasks.Add(task);
            }

            // ReSharper disable once CoVariantArrayConversion
            Task.WaitAll(tasks.ToArray());

            return new FactorPair
            {
                Custom = tasks.Select(x => x.Result.Custom).Sum(),
                Benchmark = tasks.Select(x => x.Result.Benchmark).Sum()
            };
        }

        private FactorPair CalculateAlaeToLossFactors(IEnumerable<MixedExponentialCurve> mixedExponentialCurves,
            FactorPair limitToUnlimitedFactors,
            Func<ICalculator> selectedAlaeCurveCalculator)
        {
            if (mixedExponentialCurves == null)
            {
                return new FactorPair();
            }

            var tasks = new List<Task<FactorPair>>();
            foreach (var curve in mixedExponentialCurves)
            {
                var alae = 0d;
                var benchmarkAlae = 0d;
                
                var task = new Task<FactorPair>(() =>
                {
                    foreach (var policy in curve.PolicySet)
                    {
                        var policyLimit = policy.Limit;
                        var policySir = policy.Sir;
                        var policyWeight = policy.Amount;

                        if (policyWeight <= 0) continue;

                        var result = CalculateAlaeToLoss(policyLimit, policySir, selectedAlaeCurveCalculator(), curve.CurveParameters, curve.ForceWithinLimits);

                        var weight = policyWeight * curve.Weight;
                        alae += 1 / limitToUnlimitedFactors.Custom * result.Custom / (1 + result.Custom) * weight;
                        benchmarkAlae += 1 / limitToUnlimitedFactors.Benchmark * result.Benchmark /
                            (1 + result.Benchmark) * weight;
                    }

                    return new FactorPair {Custom = alae, Benchmark = benchmarkAlae};
                });

                task.Start();
                tasks.Add(task);
            }

            // ReSharper disable once CoVariantArrayConversion
            Task.WaitAll(tasks.ToArray());

            var overallAlae = tasks.Select(x => x.Result.Custom).Sum();
            var overallBenchmarkAlae = tasks.Select(x => x.Result.Benchmark).Sum();

            return new FactorPair
            {
                Custom = overallAlae / (1 / limitToUnlimitedFactors.Custom - overallAlae),
                Benchmark =
                    overallBenchmarkAlae /
                    (1 / limitToUnlimitedFactors.Benchmark - overallBenchmarkAlae)
            };
        }
    }
}