using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.Common.Extensions;
using MramUwpfLibrary.Common.Policies;
using MramUwpfLibrary.ExposureRatingModel.Casualty.CurveInputHelpers;
using MramUwpfLibrary.ExposureRatingModel.Casualty.Curves;
using MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials;
using MramUwpfLibrary.ExposureRatingModel.Casualty.ResultAggregators;
using MramUwpfLibrary.ExposureRatingModel.Input;
using MramUwpfLibrary.ExposureRatingModel.Input.Casualty;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty
{
    internal class CasualtySublineCalculator : BaseSublineCalculator
    {
        private readonly ReinsuranceParameters _reinsuranceParameters;
        private readonly PolicyAlaeTreatmentType _policyAlaeTreatment;
        private readonly ICasualtyPrimarySublineInput _sublineInput;
        private readonly CasualtyCurveProviderResultSet _casualtyCurveResultSet;
        private readonly IList<MixedExponentialCurve> _meCurves;
        private readonly IList<TruncatedParetoCurve> _tpCurves;

        public CasualtySublineCalculator(ReinsuranceParameters reinsuranceParameters, 
            PolicyAlaeTreatmentType policyAlaeTreatment,
            ICasualtyPrimarySublineInput sublineInput, 
            CasualtyCurveProviderResultSet casualtyCurveResultSet) : base(reinsuranceParameters, sublineInput, policyAlaeTreatment)
        {
            _reinsuranceParameters = reinsuranceParameters;
            _policyAlaeTreatment = policyAlaeTreatment;
            _sublineInput = sublineInput;
            _casualtyCurveResultSet = casualtyCurveResultSet;
            
            _meCurves = new List<MixedExponentialCurve>();
            _tpCurves = new List<TruncatedParetoCurve>();
        }

        public override ExposureRatingResultItem Calculate(GrossUpLossRatioFactors grossUpLossRatioFactors)
        {
            var exposureRatingResultItems = new List<ExposureRatingResultItem>();
            var limitedLoss = _sublineInput.AllocatedExposureAmount * _sublineInput.LimitedLossRatio;
            var unlimitedOverLimitedFactor = ExposureRatingCalculatorShared.GetUnlimitedOverLimitedFactor(_reinsuranceParameters.AlaeTreatment, grossUpLossRatioFactors, limitedLoss);

            var tasks = new List<Task<ExposureRatingResultItem>>();
            if (_meCurves != null)
            {
                var layerCalculator = CalculatorFactory.Create(_reinsuranceParameters.AlaeTreatment, _policyAlaeTreatment);
                var reinsuranceAlaeTreatment = _policyAlaeTreatment.MapToReinsuranceAlaeTreatmentTypeForLayer(_reinsuranceParameters.AlaeTreatment);
                var policyCalculator = CalculatorFactory.Create(reinsuranceAlaeTreatment, _policyAlaeTreatment);

                foreach (var mixedExponentialCurve in _meCurves)
                {
                    var task = new Task<ExposureRatingResultItem>(() =>
                    {
                        var policyLayerResultsItems = CalculateReinsuranceLayerResults(mixedExponentialCurve,
                            layerCalculator.CreateNew(),
                            policyCalculator.CreateNew());

                        var layerResult = LayerResultsMapper.Map(policyLayerResultsItems, 
                            _sublineInput.AllocatedExposureAmount,
                            grossUpLossRatioFactors,
                            limitedLoss, 
                            unlimitedOverLimitedFactor);

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
                            Weight = mixedExponentialCurve.Weight * _sublineInput.AllocatedExposureAmount
                        };
                    });
                    task.Start();
                    tasks.Add(task);
                }
            }

            if (_tpCurves != null)
            {
                foreach (var truncatedParetoCurve in _tpCurves)
                {
                    var layerCalculator = Curves.TruncatedParetos.CalculatorFactory.Create(_reinsuranceParameters.AlaeTreatment, _policyAlaeTreatment);
                    var reinsuranceAlaeTreatment = _policyAlaeTreatment.MapToReinsuranceAlaeTreatmentTypeForLayer(_reinsuranceParameters.AlaeTreatment);
                    var policyCalculator = Curves.TruncatedParetos.CalculatorFactory.Create(reinsuranceAlaeTreatment, _policyAlaeTreatment);

                    var task = new Task<ExposureRatingResultItem>(() =>
                    {
                        var policyLayerResultsItems = CalculateReinsuranceLayerResults(truncatedParetoCurve,
                            layerCalculator.CreateNew(),
                            policyCalculator.CreateNew());

                        var layerResult = LayerResultsMapper.Map(policyLayerResultsItems , 
                            _sublineInput.AllocatedExposureAmount,
                            grossUpLossRatioFactors, 
                            limitedLoss,
                            unlimitedOverLimitedFactor);

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
                            Weight = truncatedParetoCurve.Weight * _sublineInput.AllocatedExposureAmount
                        };
                    });
                    task.Start();
                    tasks.Add(task);
                }
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
            if (!_meCurves.Any() && !_tpCurves.Any())
            {
                SetCurves();
            }

            var limitedAlaeTreatment = _sublineInput.LossRatioAlaeTreatment.MapToReinsuranceAlaeTreatmentType();
            var unlimitedAlaeTreatment = _policyAlaeTreatment.MapToReinsuranceAlaeTreatmentTypeForUnlimitedLossRatio();
            
            var meLimitedLossRatioCurveCalculator = CalculatorFactory.Create(limitedAlaeTreatment, _policyAlaeTreatment);
            var meUnlimitedLossRatioCurveCalculator = CalculatorFactory.Create(unlimitedAlaeTreatment, _policyAlaeTreatment);
            var meSelectedAlaeCalculator = CalculatorFactory.Create(ReinsuranceAlaeTreatmentType.ProRata, PolicyAlaeTreatmentType.InAdditionToLimit);

            var tpLimitedLossRatioCurveCalculator = Curves.TruncatedParetos.CalculatorFactory.Create(limitedAlaeTreatment, _policyAlaeTreatment);
            var tpUnlimitedLossRatioCurveCalculator = Curves.TruncatedParetos.CalculatorFactory.Create(unlimitedAlaeTreatment, _policyAlaeTreatment);
            var tpSelectedAlaeCalculator = Curves.TruncatedParetos.CalculatorFactory.Create(ReinsuranceAlaeTreatmentType.ProRata, PolicyAlaeTreatmentType.InAdditionToLimit);

            var meLimitedToUnlimitedFactors = CalculateLimitedToUnlimitedLossFactors(
                _meCurves,
                meLimitedLossRatioCurveCalculator.CreateNew,
                meUnlimitedLossRatioCurveCalculator.CreateNew);

            var tpLimitedToUnlimitedFactors = CalculateLimitedToUnlimitedLossFactors(
                _tpCurves,
                tpLimitedLossRatioCurveCalculator.CreateNew,
                tpUnlimitedLossRatioCurveCalculator.CreateNew,
                tpSelectedAlaeCalculator.CreateNew);

            var limitedToUnlimitedFactors = FactorPair.Add(meLimitedToUnlimitedFactors, tpLimitedToUnlimitedFactors);

            var meAlaeToLossFactors = CalculateAlaeToLossFactors(
                _meCurves,
                limitedToUnlimitedFactors,
                meSelectedAlaeCalculator.CreateNew);

            var tpAlaeToLossFactors = CalculateAlaeToLossFactors(
                _tpCurves,
                limitedToUnlimitedFactors,
                tpLimitedLossRatioCurveCalculator.CreateNew,
                tpUnlimitedLossRatioCurveCalculator.CreateNew,
                tpSelectedAlaeCalculator.CreateNew);

            var alaeToLossFactors = FactorPair.Add(meAlaeToLossFactors, tpAlaeToLossFactors);

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

        internal void SetCurves()
        {
            _sublineInput.CasualtyCurveContainer.PolicySet.NormalizeAmounts();

            foreach (var curveResult in _casualtyCurveResultSet.Curves.Where(curve => curve.MixedExponentialCurveParameters != null))
            {
                var curveId = curveResult.CurveId.ToString();
                _meCurves.Add(new MixedExponentialCurve
                {
                    Id = curveId,
                    Weight = curveResult.Weight,
                    CurveParameters = curveResult.MixedExponentialCurveParameters,
                    PolicySet = _sublineInput.CasualtyCurveContainer.PolicySet
                });
            }

            foreach (var curveResult in _casualtyCurveResultSet.Curves.Where(curve => curve.TruncatedParetoCurveParameters != null))
            {
                var curveId = curveResult.CurveId.ToString();
                _tpCurves.Add(new TruncatedParetoCurve
                {
                    Id = curveId,
                    Weight = curveResult.Weight,
                    CurveParameters = curveResult.TruncatedParetoCurveParameters,
                    PolicySet = _sublineInput.CasualtyCurveContainer.PolicySet
                });
            }
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

        private IEnumerable<LayerResultItem> CalculateReinsuranceLayerResults(TruncatedParetoCurve tpCurve,
            Curves.TruncatedParetos.ICalculator reinsCurveCalculator, Curves.TruncatedParetos.ICalculator policyCurveCalculator)
        {
            return tpCurve.PolicySet
                .Select(policyAndNormalizedWeight => 
                    CalculateReinsuranceLayerResult(_reinsuranceParameters, tpCurve.CurveParameters, 
                        policyAndNormalizedWeight, reinsCurveCalculator, policyCurveCalculator)).ToList();
        }

        private FactorPair CalculateLimitedToUnlimitedLossFactors(IEnumerable<MixedExponentialCurve> meCurves,
            Func<ICalculator> limitedLossRatioCurveCalculator,
            Func<ICalculator> unlimitedLossRatioCurveCalculator)
        {
            if (meCurves == null)
            {
                return new FactorPair();
            }

            var tasks = new List<Task<FactorPair>>();
            foreach (var mixedExponentialCurve in meCurves)
            {
                var task = new Task<FactorPair>(() =>
                {
                    var pairAndWeights = new List<FactorPairAndWeight>(); 
                    
                    var parameters = mixedExponentialCurve.CurveParameters;
                    
                    foreach (var policy in mixedExponentialCurve.PolicySet)
                    {
                        var policyLimit = policy.Limit;
                        var policySir = policy.Sir;
                        var policyWeight = policy.Amount;

                        if (policyWeight.IsEqualToZero()) continue;

                        var result = CalculateLimitedToUnlimited(policyLimit,
                            policySir,
                            limitedLossRatioCurveCalculator(),
                            unlimitedLossRatioCurveCalculator(),
                            parameters);

                        var weight = policyWeight * mixedExponentialCurve.Weight;
                        
                        pairAndWeights.Add(new FactorPairAndWeight
                        {
                            Custom = result.Custom,
                            Benchmark = result.Benchmark,
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

        private FactorPair CalculateLimitedToUnlimitedLossFactors(
            IEnumerable<TruncatedParetoCurve> tpCurves,
            Func<Curves.TruncatedParetos.ICalculator> limitedLossRatioCurveCalculator,
            Func<Curves.TruncatedParetos.ICalculator> unlimitedLossRatioCurveCalculator,
            Func<Curves.TruncatedParetos.ICalculator> selectedAlaeCalculator)
        {
            if (tpCurves == null)
            {
                return new FactorPair();
            }

            var tasks = new List<Task<FactorPair>>();
            foreach (var truncatedParetoCurve in tpCurves)
            {
                var task = new Task<FactorPair>(() =>
                {
                    var limitedToUnlimitedRatio = 0d;
                    var benchmarkLimitedToUnlimitedRatio = 0d;
                    var parameters = truncatedParetoCurve.CurveParameters;

                    foreach (var policy in truncatedParetoCurve.PolicySet)
                    {
                        var policyLimit = policy.Limit;
                        var policySir = policy.Sir ?? 0;
                        var policyWeight = policy.Amount;
                        
                        if (policyWeight.IsEqualToZero()) continue;

                        var result = CalculateAlaeResult(
                            parameters,
                            policyLimit,
                            policySir,
                            limitedLossRatioCurveCalculator(),
                            unlimitedLossRatioCurveCalculator(),
                            selectedAlaeCalculator());

                        if (result.LimitedToUnlimited.Custom.IsEqualToZero()) result.LimitedToUnlimited.Custom = 1;
                        if (result.LimitedToUnlimited.Benchmark.IsEqualToZero()) result.LimitedToUnlimited.Benchmark = 1;

                        var weight = policyWeight * truncatedParetoCurve.Weight;
                        limitedToUnlimitedRatio += result.LimitedToUnlimited.Custom * weight;
                        benchmarkLimitedToUnlimitedRatio += result.LimitedToUnlimited.Benchmark * weight;
                    }
                    return new FactorPair {Custom = limitedToUnlimitedRatio, Benchmark = benchmarkLimitedToUnlimitedRatio};
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

        private FactorPair CalculateAlaeToLossFactors(IEnumerable<MixedExponentialCurve> meCurves,
            FactorPair limitToUnlimitedFactors,
            Func<ICalculator> selectedAlaeCurveCalculator)
        {
            if (meCurves == null)
            {
                return new FactorPair();
            }

            var tasks = new List<Task<FactorPair>>();
            foreach (var mixedExponentialCurve in meCurves)
            {
                var alae = 0d;
                var benchmarkAlae = 0d;
                
                var task = new Task<FactorPair>(() =>
                {
                    foreach (var policy in mixedExponentialCurve.PolicySet)
                    {
                        var policyLimit = policy.Limit;
                        var policySir = policy.Sir;
                        var policyWeight = policy.Amount;

                        if (policyWeight <= 0) continue;

                        var alaeToLoss = CalculateAlaeToLoss(policyLimit, policySir, selectedAlaeCurveCalculator(), mixedExponentialCurve.CurveParameters);

                        var weight = policyWeight * mixedExponentialCurve.Weight;
                        alae += 1 / limitToUnlimitedFactors.Custom * alaeToLoss.Custom / (1 + alaeToLoss.Custom) * weight;
                        benchmarkAlae += 1 / limitToUnlimitedFactors.Benchmark * alaeToLoss.Benchmark / (1 + alaeToLoss.Benchmark) * weight;
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


        private FactorPair CalculateAlaeToLossFactors(IEnumerable<TruncatedParetoCurve> tpCurves,
            FactorPair limitedOverUnlimitedFactorPair,
            Func<Curves.TruncatedParetos.ICalculator> limitedLossRatioCurveCalculator,
            Func<Curves.TruncatedParetos.ICalculator> unlimitedLossRatioCurveCalculator,
            Func<Curves.TruncatedParetos.ICalculator> selectedAlaeCurveCalculator)
        {
            if (tpCurves == null)
            {
                return new FactorPair();
            }

            var tasks = new List<Task<FactorPair>>();
            foreach (var truncatedParetoCurve in tpCurves)
            {
                var alae = 0d;
                var benchmarkAlae = 0d;
                

                var task = new Task<FactorPair>(() =>
                {
                    foreach (var policy in truncatedParetoCurve.PolicySet)
                    {
                        var policyLimit = policy.Limit;
                        var policySir = policy.Sir;
                        var policyWeight = policy.Amount;

                        if (policyWeight.IsEqualToZero()) continue;
                        
                        var result = CalculateAlaeResult(
                            truncatedParetoCurve.CurveParameters,
                            policyLimit,
                            policySir ?? 0,
                            limitedLossRatioCurveCalculator(),
                            unlimitedLossRatioCurveCalculator(),
                            selectedAlaeCurveCalculator());

                        var weight = policyWeight * truncatedParetoCurve.Weight;

                        alae += 1 / limitedOverUnlimitedFactorPair.Custom * result.AlaeToLoss.Custom / (1 + result.AlaeToLoss.Custom) * weight;

                        benchmarkAlae += 1 / limitedOverUnlimitedFactorPair.Benchmark *
                            result.AlaeToLoss.Benchmark / (1 + result.AlaeToLoss.Benchmark) * weight;
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
                Custom = overallAlae / (1 / limitedOverUnlimitedFactorPair.Custom - overallAlae),
                Benchmark = overallBenchmarkAlae / (1 / limitedOverUnlimitedFactorPair.Benchmark - overallBenchmarkAlae)
            };
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

        private LayerResultItem CalculateReinsuranceLayerResult(ReinsuranceParameters reinsurance,
            Curves.TruncatedParetos.Parameters parameters, 
            IPolicy policy,
            Curves.TruncatedParetos.ICalculator reinsCurveCalculator,
            Curves.TruncatedParetos.ICalculator policyCurveCalculator)
        {
            IHelper helper = new ReinsuranceLayerHelper(reinsurance,
                _policyAlaeTreatment,
                _sublineInput,
                policy.Limit, 
                policy.Sir ?? 0);

            var reinsuranceParameters = helper.GetNumeratorCurveInputs();
            var policyParameters = helper.GetDenominatorCurveInputs();

            var reinsuranceLimitedExpectedValue = GetLayerLimitedExpectedValue(
                parameters,
                policy.Limit,
                policy.Sir ?? 0,
                reinsuranceParameters,
                reinsCurveCalculator);

            var policyLimitedExpectedValue = GetLayerLimitedExpectedValue(
                parameters,
                policy.Limit,
                policy.Sir ?? 0,
                policyParameters,
                policyCurveCalculator);

            var exposureRate = reinsuranceLimitedExpectedValue.DivideByWithTrap(policyLimitedExpectedValue);

            var probabilityLessThanReinsuranceBottom = reinsCurveCalculator.GetCdfForLimit(
                reinsuranceParameters.BottomLimit,
                policy.Limit,
                policy.Sir ?? 0,
                reinsuranceParameters.AlaeAdjustmentFactor,
                reinsuranceParameters.ReinsurancePerspective,
                parameters);

            var averageSeverity = probabilityLessThanReinsuranceBottom < 1 
                ? reinsuranceLimitedExpectedValue / (1 - probabilityLessThanReinsuranceBottom) 
                : 0;

            return new LayerResultItem
            {
                LayerLevOverPolicyLev = exposureRate,
                Severity = averageSeverity,
                Weight = policy.Amount
            };

        }

        

        private GrossUpLossRatioFactors CalculateAlaeResult(Curves.TruncatedParetos.Parameters parameters, double policyLimit, double policySir,
            Curves.TruncatedParetos.ICalculator limitedLossRatioCurveCalculator,
            Curves.TruncatedParetos.ICalculator unlimitedLossRatioCurveCalculator, 
            Curves.TruncatedParetos.ICalculator selectedAlaeCalculator)
        {
            IHelper helper = new LossRatioHelper(_policyAlaeTreatment,
                _sublineInput, policyLimit, policySir, _reinsuranceParameters.PerspectiveHandler);

            var limitedLossRatioLev = GetLayerLimitedExpectedValue(
                parameters,
                policyLimit,
                policySir,
                helper.GetNumeratorCurveInputs(),
                limitedLossRatioCurveCalculator);

            var unlimitedLossRatioLev = GetLayerLimitedExpectedValue(
                parameters,
                policyLimit,
                policySir,
                helper.GetDenominatorCurveInputs(),
                unlimitedLossRatioCurveCalculator);

            // ReSharper disable once CompareOfFloatsByEqualityOperator
            var isNoAlaeAdjustment = _sublineInput.AlaeAdjustmentFactor == 1;
            var benchmarkLimitedLossRatioLev = isNoAlaeAdjustment
                ? limitedLossRatioLev
                : GetLayerLimitedExpectedValue(
                    parameters,
                    policyLimit,
                    policySir,
                    helper.GetNumeratorCurveBenchmarkInputs(),
                    limitedLossRatioCurveCalculator);

            var benchmarkUnlimitedLossRatioLev = isNoAlaeAdjustment
                ? unlimitedLossRatioLev
                : GetLayerLimitedExpectedValue(
                    parameters,
                    policyLimit,
                    policySir,
                    helper.GetDenominatorCurveBenchmarkInputs(),
                    unlimitedLossRatioCurveCalculator);

            var limToUnlim = limitedLossRatioLev.DivideByWithTrap(unlimitedLossRatioLev);
            var benchmarkLimToUnlim = benchmarkLimitedLossRatioLev.DivideByWithTrap(benchmarkUnlimitedLossRatioLev);

            if (_policyAlaeTreatment != PolicyAlaeTreatmentType.InAdditionToLimit)
            {
                var limitedToUnlimited = new LimitedToUnlimitedFactors
                {
                    Custom = limToUnlim,
                    Benchmark = benchmarkLimToUnlim,
                };

                var alaeToLoss = new AlaeToLossFactors
                {
                    Custom = 0,
                    Benchmark = 0,
                };
                return new GrossUpLossRatioFactors(limitedToUnlimited, alaeToLoss);
            }

            helper = new SelectedAlaeHelper(_sublineInput, policyLimit, policySir);

            var policyLev = GetLayerLimitedExpectedValue(
                parameters,
                policyLimit,
                policySir,
                helper.GetNumeratorCurveInputs(),
                selectedAlaeCalculator);

            var benchmarkPolicyLev = GetLayerLimitedExpectedValue(
                parameters,
                policyLimit,
                policySir,
                helper.GetNumeratorCurveBenchmarkInputs(),
                selectedAlaeCalculator);

            var policyLevUsingLossOnly = GetLayerLimitedExpectedValue(
                parameters,
                policyLimit,
                policySir,
                helper.GetDenominatorCurveInputs(),
                selectedAlaeCalculator);

            var alaeRatio = policyLev.DivideByWithTrap(policyLevUsingLossOnly, 1) - 1;
            var benchmarkAlaeRatio = benchmarkPolicyLev.DivideByWithTrap(policyLevUsingLossOnly,1) - 1;

            var limitedToUnlimited2 = new LimitedToUnlimitedFactors
            {
                Custom = limToUnlim,
                Benchmark = benchmarkLimToUnlim,
            };
         
            var alaeToLoss2  = new AlaeToLossFactors
            {
                Custom = alaeRatio,
                Benchmark = benchmarkAlaeRatio
            };

            return new GrossUpLossRatioFactors(limitedToUnlimited2, alaeToLoss2);
        }

        

        private static double GetLayerLimitedExpectedValue(Curves.TruncatedParetos.Parameters parameters, double policyLimit, double policySir,
            CurveInputs curveInputs, Curves.TruncatedParetos.ICalculator curveCalculator)
        {
            return curveCalculator.GetLayerLimitedExpectedValue(
                curveInputs.TopLimit,
                curveInputs.BottomLimit,
                policyLimit,
                policySir,
                curveInputs.AlaeAdjustmentFactor,
                curveInputs.ReinsuranceAlaeTreatment,
                curveInputs.ReinsurancePerspective,
                parameters);
        }

    }
}