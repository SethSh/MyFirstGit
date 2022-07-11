using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.Common.Extensions;
using MramUwpfLibrary.Common.Layers;
using MramUwpfLibrary.ExposureRatingModel.Casualty.CurveInputHelpers;
using MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials;
using MramUwpfLibrary.ExposureRatingModel.Casualty.ResultAggregators;
using MramUwpfLibrary.ExposureRatingModel.Input;

namespace MramUwpfLibrary.ExposureRatingModel.Property
{
    internal class PropertySublineCalculator : BaseSublineCalculator
    {
        private const double ProbabilityOfNoLoss = 0;
        private const double AlaeForClaimsWithoutPay = 0;
        private const double AlaeAdjustmentFactor = 1;

        private readonly ReinsuranceParameters _reinsuranceParameters;
        private readonly PolicyAlaeTreatmentType _policyAlaeTreatment;
        private readonly ISublineExposureRatingInput _sublineInput;
        private readonly IList<ITotalInsuredValueItem> _totalInsuredValueItems;
        private readonly double _alaeToLossPercent;

        public PropertySublineCalculator(ReinsuranceParameters reinsuranceParameters,
            PolicyAlaeTreatmentType policyAlaeTreatment,
            ISublineExposureRatingInput sublineInput,
            IList<ITotalInsuredValueItem> totalInsuredValueItems,
            double alaeToLossPercent) : base(reinsuranceParameters, sublineInput, policyAlaeTreatment)
        {
            _reinsuranceParameters = reinsuranceParameters;
            _policyAlaeTreatment = policyAlaeTreatment;
            _sublineInput = sublineInput;
            _totalInsuredValueItems = totalInsuredValueItems;
            _alaeToLossPercent = alaeToLossPercent;

            _sublineInput.AlaeAdjustmentFactor = AlaeAdjustmentFactor;
        }

        public override ExposureRatingResultItem Calculate(GrossUpLossRatioFactors grossUpLossRatioFactors)
        {
            var exposureRatingResults = new List<ExposureRatingResultItem>();
            
            var limitedLoss = _sublineInput.AllocatedExposureAmount * _sublineInput.LimitedLossRatio;
            var unlimitedOverLimitedFactor = ExposureRatingCalculatorShared.GetUnlimitedOverLimitedFactor(
                _reinsuranceParameters.AlaeTreatment, grossUpLossRatioFactors, limitedLoss);

            var tasks = new List<Task<ExposureRatingResultItem>>();

            var layerCalculator = CalculatorFactory.Create(_reinsuranceParameters.AlaeTreatment, _policyAlaeTreatment);
            var reinsuranceAlaeTreatment = _policyAlaeTreatment.MapToReinsuranceAlaeTreatmentTypeForLayer(_reinsuranceParameters.AlaeTreatment);
            var policyCalculator = CalculatorFactory.Create(reinsuranceAlaeTreatment, _policyAlaeTreatment);

            foreach (var totalInsuredValueItem in _totalInsuredValueItems.Where(item => item.Weight > 0))
            {
                var task = new Task<ExposureRatingResultItem>(() =>
                {
                    var curveLayerResultItems = CalculateReinsuranceLayerItems(totalInsuredValueItem, 
                        layerCalculator.CreateNew(), 
                        policyCalculator.CreateNew());
                     
                    // PolicyAggregator probably not necessary, but aligned with casualty
                    // todo aggregate is not a good name
                    var layerResult = LayerResultsMapper.Map(curveLayerResultItems,
                        _sublineInput.AllocatedExposureAmount,
                        grossUpLossRatioFactors,
                        limitedLoss,
                        unlimitedOverLimitedFactor,
                        totalInsuredValueItem.Share);

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
                        Weight = totalInsuredValueItem.Weight
                    };
                });
                task.Start();
                tasks.Add(task);
            }

            // ReSharper disable once CoVariantArrayConversion
            Task.WaitAll(tasks.ToArray());
            exposureRatingResults.AddRange(tasks.Select(t => t.Result));

            var overallLayerResult = ExposureRatingResultAggregator.Aggregate(exposureRatingResults);
            overallLayerResult.SublineId = _sublineInput.Id;

            overallLayerResult.CheckForNan();
            return overallLayerResult;
        }

        public override GrossUpLossRatioFactors GrossUpLossRatio()
        {
            var reinsuranceAlaeTreatmentForLimitedLossRatio = _sublineInput.LossRatioAlaeTreatment.MapToReinsuranceAlaeTreatmentType(); 
            var reinsuranceAlaeTreatmentForUnlimitedLossRatio = _policyAlaeTreatment.MapToReinsuranceAlaeTreatmentTypeForUnlimitedLossRatio();

            var limitedLossRatioCurveCalculator = CalculatorFactory.Create(reinsuranceAlaeTreatmentForLimitedLossRatio, _policyAlaeTreatment);
            var unlimitedLossRatioCurveCalculator = CalculatorFactory.Create(reinsuranceAlaeTreatmentForUnlimitedLossRatio, _policyAlaeTreatment);
            var selectedAlaeCalculator = CalculatorFactory.Create(ReinsuranceAlaeTreatmentType.ProRata, PolicyAlaeTreatmentType.InAdditionToLimit);

            var limitedToUnlimitedLossFactors = CalculateLimitedToUnlimitedLossFactors(
                _totalInsuredValueItems,
                limitedLossRatioCurveCalculator.CreateNew,
                unlimitedLossRatioCurveCalculator.CreateNew);

            var alaeToLossFactors = CalculateAlaeToLossFactors(
                _totalInsuredValueItems,
                limitedToUnlimitedLossFactors,
                selectedAlaeCalculator.CreateNew);

            var limitedToUnlimited = new LimitedToUnlimitedFactors
            {
                Custom = limitedToUnlimitedLossFactors.Custom,
                Benchmark = limitedToUnlimitedLossFactors.Benchmark,
                CustomUnlimitedLossRatio = _sublineInput.LimitedLossRatio / limitedToUnlimitedLossFactors.Custom,
                BenchmarkUnlimitedLossRatio = _sublineInput.LimitedLossRatio / limitedToUnlimitedLossFactors.Benchmark
            };

            var alaeToLoss = new AlaeToLossFactors
            {
                Custom = alaeToLossFactors.Custom,
                Benchmark = alaeToLossFactors.Benchmark,
            };
            
            return new GrossUpLossRatioFactors(limitedToUnlimited, alaeToLoss);
        }


        private FactorPair CalculateLimitedToUnlimitedLossFactors(
            IEnumerable<ITotalInsuredValueItem> totalInsuredValueItems,
            Func<ICalculator> limitedLossRatioCurveCalculator,
            Func<ICalculator> unlimitedLossRatioCurveCalculator)
        {
            var tasks = new List<Task<FactorPair>>();
            foreach (var totalInsuredValueItem in totalInsuredValueItems.Where(item => item.Weight > 0))
            {
                var task = new Task<FactorPair>(() =>
                {
                    var pairAndWeights = new List<FactorPairAndWeight>();
                    
                    foreach (var curve in totalInsuredValueItem.PropertyCurveProcessors)
                    {
                        var curveWeight = curve.Weight;
                        if (curveWeight.IsEqualToZero()) continue; 
                        
                        var parameters = MapParameters(curve.PropertyCurve.MixedExponentialCurveParameters);

                        //full means at 100 percent
                        var fullPolicyLimit = GetFullLimit(totalInsuredValueItem.Limit, totalInsuredValueItem.Share);
                        var fullPolicySir = GetFullAttachment(totalInsuredValueItem.Attachment, totalInsuredValueItem.Share);

                        var limitedToUnlimited = CalculateLimitedToUnlimited(fullPolicyLimit,
                            fullPolicySir,
                            limitedLossRatioCurveCalculator(),
                            unlimitedLossRatioCurveCalculator(),
                            parameters);

                        var weight = totalInsuredValueItem.Weight * curveWeight; 
                        
                        pairAndWeights.Add(new FactorPairAndWeight
                        {
                            Custom = limitedToUnlimited.Custom,
                            Benchmark = limitedToUnlimited.Benchmark,
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
        
        private FactorPair CalculateAlaeToLossFactors(IEnumerable<ITotalInsuredValueItem> totalInsuredValueItems,
            FactorPair limitedOverUnlimitedFactors,
            Func<ICalculator> selectedAlaeCurveCalculator)
        {
            var tasks = new List<Task<FactorPair>>();
            foreach (var totalInsuredValueItem in totalInsuredValueItems.Where(item => item.Weight > 0))
            {
                var alae = 0d;
                var benchmarkAlae = 0d;

                var task = new Task<FactorPair>(() =>
                {
                    foreach (var curve in totalInsuredValueItem.PropertyCurveProcessors)
                    {
                        var curveWeight = curve.Weight;
                        if (curveWeight.IsEqualToZero()) continue;

                        var parameters = MapParameters(curve.PropertyCurve.MixedExponentialCurveParameters);

                        //full means at 100 percent
                        var fullPolicyLimit = GetFullLimit(totalInsuredValueItem.Limit, totalInsuredValueItem.Share);
                        var fullPolicySir = GetFullAttachment(totalInsuredValueItem.Attachment, totalInsuredValueItem.Share);

                        var alaeToLoss = CalculateAlaeToLoss(fullPolicyLimit, fullPolicySir, selectedAlaeCurveCalculator(), parameters);

                        var weight = curveWeight * totalInsuredValueItem.Weight;
                        alae += 1 / limitedOverUnlimitedFactors.Custom * alaeToLoss.Custom / (1 + alaeToLoss.Custom) * weight;
                        benchmarkAlae += 1 / limitedOverUnlimitedFactors.Benchmark * alaeToLoss.Benchmark /
                            (1 + alaeToLoss.Benchmark) * weight;
                    }

                    return new FactorPair
                    {
                        Custom = alae, 
                        Benchmark = benchmarkAlae
                    };
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
                Custom = overallAlae / (1 / limitedOverUnlimitedFactors.Custom - overallAlae),
                Benchmark = 
                    overallBenchmarkAlae /
                    (1 / limitedOverUnlimitedFactors.Benchmark - overallBenchmarkAlae)
            };
        }

        private IEnumerable<LayerResultItem> CalculateReinsuranceLayerItems(ITotalInsuredValueItem totalInsuredValueItem,
            ICalculator layerCurveCalculator,
            ICalculator policyCurveCalculator)
        {
            //full means at 100 percent
            var fullPolicyLimit = GetFullLimit(totalInsuredValueItem.Limit, totalInsuredValueItem.Share);
            var fullPolicySir = GetFullAttachment(totalInsuredValueItem.Attachment, totalInsuredValueItem.Share);
            var fullLayer = GetFullReinsuranceLayer(_reinsuranceParameters.Layer, totalInsuredValueItem.Share);

            var fullReinsuranceParameters = new ReinsuranceParameters(fullLayer, _reinsuranceParameters.AlaeTreatment, _reinsuranceParameters.PerspectiveHandler);

            var helper = new ReinsuranceLayerHelper(fullReinsuranceParameters,
                _policyAlaeTreatment,
                _sublineInput,
                fullPolicyLimit, fullPolicySir);

            var fullLayerParameters = helper.GetNumeratorCurveInputs();
            var fullPolicyParameters = helper.GetDenominatorCurveInputs();

            var layerResultItems = new List<LayerResultItem>();
            foreach (var curve in totalInsuredValueItem.PropertyCurveProcessors)
            {
                var parameters = MapParameters(curve.PropertyCurve.MixedExponentialCurveParameters);
                
                var layerLev = GetLayerLimitedExpectedValue(fullPolicyLimit,
                    fullPolicySir,
                    fullLayerParameters,
                    layerCurveCalculator, 
                    parameters);

                var probLessThanLayerBottom = layerCurveCalculator.GetProbabilityLessThanLimit(
                    fullLayerParameters.BottomLimit,
                    fullPolicyLimit,
                    fullPolicySir,
                    fullLayerParameters.AlaeAdjustmentFactor,
                    fullLayerParameters.ReinsurancePerspective,
                    parameters);

                var policyLev = GetLayerLimitedExpectedValue(fullPolicyLimit,
                    fullPolicySir,
                    fullPolicyParameters,
                    policyCurveCalculator,
                    parameters);

                var layerLevOverPolicyLev = layerLev.DivideByWithTrap(policyLev);
                var severity = probLessThanLayerBottom < 1
                    ? layerLev / (1 - probLessThanLayerBottom)
                    : 0;

                var severityAtShare = severity * totalInsuredValueItem.Share ?? severity;

                layerResultItems.Add(new LayerResultItem
                {
                    LayerLevOverPolicyLev = layerLevOverPolicyLev,
                    Severity = severityAtShare,
                    Weight = curve.Weight
                });
            }

            return layerResultItems;
        }

        private static ILayer GetFullReinsuranceLayer(ILayer layer, double? share)
        {
            var layerAtOneHundredPercent = (ILayer)layer.Clone();
            layerAtOneHundredPercent.Limit = GetFullLimit(layerAtOneHundredPercent.Limit, share);
            layerAtOneHundredPercent.Attachment = GetFullAttachment(layerAtOneHundredPercent.Attachment, share);
            return layerAtOneHundredPercent;
        }

        private static double GetFullLimit(double? limit, double? share)
        {
            double limitAtOneHundredPercent;
            if (limit.HasValue)
            {
                limitAtOneHundredPercent = share.HasValue 
                    ? limit.Value.DivideByWithTrap(share.Value) 
                    : limit.Value;
            }
            else
            {
                limitAtOneHundredPercent = double.MaxValue;
            }

            return limitAtOneHundredPercent;
        }

        private static double GetFullAttachment(double? attachment, double? share)
        {
            double attachmentAtOneHundredPercent;
            if (attachment.HasValue)
            {
                attachmentAtOneHundredPercent = share.HasValue
                    ? attachment.Value.DivideByWithTrap(share.Value)
                    : attachment.Value;
            }
            else
            {
                attachmentAtOneHundredPercent = 0d;
            }

            return attachmentAtOneHundredPercent;
        }

        private Parameters MapParameters(IMixedExponentialParametersCore meParameters)
        {
            var means = meParameters.Means;
            var weights = meParameters.Weights;
            var alaePercents = Enumerable.Repeat(_alaeToLossPercent, means.Length).ToArray();
            return new Parameters(means, weights, alaePercents, ProbabilityOfNoLoss, AlaeForClaimsWithoutPay);
        }

        
    }
}