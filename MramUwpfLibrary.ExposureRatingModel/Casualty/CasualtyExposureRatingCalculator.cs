using System;
using System.Collections.Generic;
using System.Linq;
using MramUwpfLibrary.Common.Extensions;
using MramUwpfLibrary.ExposureRatingModel.Input;
using MramUwpfLibrary.ExposureRatingModel.Input.Casualty;


namespace MramUwpfLibrary.ExposureRatingModel.Casualty
{
    public interface ICasualtyExposureRatingCalculator
    {
        IEnumerable<IExposureRatingResult> CalculateExposureRatingResults(ICasualtyExposureRatingInput input);
    }

    public class CasualtyExposureRatingCalculator : ICasualtyExposureRatingCalculator
    {
        private readonly ISeverityCurveProvider _curveProvider;

        public CasualtyExposureRatingCalculator(ISeverityCurveProvider curveProvider)
        {
            _curveProvider = curveProvider;
        }

        public IEnumerable<IExposureRatingResult> CalculateExposureRatingResults(ICasualtyExposureRatingInput input)
        {
            var inputClone = input;

            SetSublineNormalizedAllocations(inputClone);
            SetUmbrellaNormalizedAllocations(inputClone);

            var curveSets = FetchCurves(inputClone, _curveProvider).ToList();

            var lossRatioSets = new List<LossRatioResultSet>();
            foreach (var segmentInput in inputClone.SegmentInputs)
            {
                var reinsuranceParametersWithoutLayer = new ReinsuranceParameters(input.ReinsuranceAlaeTreatment, input.ReinsurancePerspective); 
                var segmentLossRatioSets = GrossUpSegmentLossRatios(reinsuranceParametersWithoutLayer, segmentInput, curveSets);
                lossRatioSets.AddRange(segmentLossRatioSets);
            }
            var aggregatedLossRatio = ExposureRatingCalculatorShared.AggregateLossRatioSets(inputClone.UseAlternative, lossRatioSets);


            var results = new List<IExposureRatingResult>();
            foreach (var reinsuranceLayer in inputClone.ReinsuranceLayers)
            {
                var reinsuranceParameters = new ReinsuranceParameters(reinsuranceLayer, inputClone.ReinsuranceAlaeTreatment, inputClone.ReinsurancePerspective);
                foreach (var segmentInput in inputClone.SegmentInputs)
                {
                    var segmentResultItems = SliceIntoLayer(segmentInput, curveSets, reinsuranceParameters,
                        inputClone.UseAlternative, lossRatioSets, aggregatedLossRatio);

                    var resultContainer = new ExposureRatingResult
                    {
                        LayerId = reinsuranceLayer.Id,
                        SubmissionSegmentId = segmentInput.Id,
                        Items = segmentResultItems
                    };
                    results.Add(resultContainer);
                }
            }

            return results;
        }

        
        private static IEnumerable<IExposureRatingResultItem> SliceIntoLayer(ISegmentInput segmentInput,
            IList<CasualtyCurveProviderResultSet> curveSets, ReinsuranceParameters reinsurance,
            bool useAlternative, IList<LossRatioResultSet> lossRatioInputSets, GrossUpLossRatioFactors grossUpLossRatio)
        {
            var containerResults = new List<IExposureRatingResultItem>();
            foreach (var sublineInput in segmentInput.SublineInputs)
            {
                var sublineCurveSet = curveSets.Single(curve => curve.SegmentId == segmentInput.Id && curve.SublineId == sublineInput.Id);
                var calculatorHelper = CalculatorHelperFactory.Create(sublineInput);
                var sliceResult = calculatorHelper.SliceIntoLayer(reinsurance, lossRatioInputSets, useAlternative, grossUpLossRatio, 
                    segmentInput.PolicyAlaeTreatment, sublineInput, new List<CasualtyCurveProviderResultSet> {sublineCurveSet});

                containerResults.Add(sliceResult);
            }

            return containerResults;
        }

        private static IEnumerable<LossRatioResultSet> GrossUpSegmentLossRatios(ReinsuranceParameters reinsurance,
            ISegmentInput segmentInput, List<CasualtyCurveProviderResultSet> curveSets)
        {
            var results = new List<LossRatioResultSet>();
            foreach (var sublineInput in segmentInput.SublineInputs)
            {
                var calculatorHelper = CalculatorHelperFactory.Create(sublineInput);

                var sublineCurveSet = curveSets.Single(curve => curve.SegmentId == segmentInput.Id && curve.SublineId == sublineInput.Id);
                var sublineGrossedUpLossRatio = calculatorHelper.GrossUpSublineLossRatio(reinsurance, segmentInput.PolicyAlaeTreatment,
                    sublineInput, new List<CasualtyCurveProviderResultSet> {sublineCurveSet});

                if (sublineGrossedUpLossRatio != null)
                {
                    results.Add(sublineGrossedUpLossRatio);
                }
            }
        
            return results;
        }

        
        private static void SetSublineNormalizedAllocations(ICasualtyExposureRatingInput exposureRatingInput)
        {
            foreach (var container in exposureRatingInput.SegmentInputs)
            {
                var groupIds = container.SublineInputs.Select(x => x.GroupId).Distinct();
                foreach (var groupId in groupIds)
                {
                    var sublineInputs = container.SublineInputs.Where(x => x.GroupId == groupId).ToList();
                    var allocationValueSum = sublineInputs.Select(x => x.Allocation).Sum();
                    foreach (var sublineInput in sublineInputs)
                    {
                        ((SublineExposureRatingInput) sublineInput).NormalizedAllocation =
                            sublineInput.Allocation.DivideByWithTrap(allocationValueSum);
                    }
                }
            }
        }

        private static void SetUmbrellaNormalizedAllocations(ICasualtyExposureRatingInput exposureRatingInput)
        {
            var atLeastOneUmbrella = exposureRatingInput.SegmentInputs.SelectMany(cont => cont.SublineInputs)
                .Any(si => si is CasualtyUmbrellaSublineExposureRatingInput);

            if (!atLeastOneUmbrella) return;

            var containersWithAnyUmbrella =
                exposureRatingInput.SegmentInputs.Where(cont =>
                    cont.SublineInputs.Any(si => si is CasualtyUmbrellaSublineExposureRatingInput));

            foreach (var container in containersWithAnyUmbrella)
            {
                var umbrellaSublineInputs = container.SublineInputs.Where(si => si is CasualtyUmbrellaSublineExposureRatingInput);
                foreach (var umbrellaSublineInput in umbrellaSublineInputs.Select(usi => (CasualtyUmbrellaSublineExposureRatingInput) usi))
                {
                    var umbrellaAllocationValueSum = umbrellaSublineInput.UmbrellaTypeAllocations
                        .Where(x => x.IsPersonal == umbrellaSublineInput.IsPersonal)
                        .Select(x => x.Allocation)
                        .Sum();

                    foreach (var umbrellaTypeAllocation in umbrellaSublineInput.UmbrellaTypeAllocations.Where(umbrellaTypeAllocation =>
                        umbrellaTypeAllocation.IsPersonal == umbrellaSublineInput.IsPersonal))
                    {
                        umbrellaTypeAllocation.NormalizedAllocation =
                            umbrellaTypeAllocation.Allocation.DivideByWithTrap(umbrellaAllocationValueSum);
                    }
                }
            }
        }

        private static IEnumerable<CasualtyCurveProviderResultSet> FetchCurves(ICasualtyExposureRatingInput input, ISeverityCurveProvider curveProvider)
        {
            var resultSets = new List<CasualtyCurveProviderResultSet>();
            foreach (var segmentInput in input.SegmentInputs)
            {
                var date = input.EffectiveDate;
                var segmentId = segmentInput.Id;

                foreach (var sublineInput in segmentInput.SublineInputs)
                {
                    switch (sublineInput)
                    {
                        case CasualtyPrimarySublineExposureRatingInput primarySublineInput:
                        {
                            var curveContainer = primarySublineInput.CasualtyCurveContainer;

                            var hazardAllocations = curveContainer.HazardAllocations.ToList();
                            var totalHazardWeight = hazardAllocations.Sum(alloc => alloc.Amount);
                            var sublineCurves = new List<SeverityCurveResult>();

                            foreach (var hazardAllocation in hazardAllocations)
                            {
                                var hazardWeight = hazardAllocation.Amount / totalHazardWeight;
                                if (hazardWeight <= 0) continue;

                                var stateAllocations = curveContainer.StateAllocations.Where(state => state.Amount > 0).ToList();
                                var curves = curveProvider.GetSeverityCurve(date, stateAllocations, hazardAllocation.Id,
                                    Convert.ToInt64(sublineInput.SublineCode));

                                var totalStateGroupWeight = curves.Sum(curve => curve.Weight);
                                curves.ForEach(curve => curve.Weight = curve.Weight / totalStateGroupWeight * hazardWeight);

                                sublineCurves.AddRange(curves);
                            }

                            resultSets.Add(new CasualtyCurveProviderResultSet
                            {
                                SegmentId = segmentId, SublineId = sublineInput.Id, Curves = sublineCurves
                            });
                            break;
                        }
                        case CasualtyUmbrellaSublineExposureRatingInput umbrellaSublineInput:
                        {
                            //hazard and state don't vary by umbrella type
                            var curveContainer = umbrellaSublineInput.UmbrellaTypeAllocations.First().CasualtyCurveSetContainer;

                            var hazardAllocations = curveContainer.HazardAllocations.ToList();
                            var totalHazardWeight = hazardAllocations.Sum(alloc => alloc.Amount);
                            var sublineCurves = new List<SeverityCurveResult>();

                            foreach (var hazardAllocation in hazardAllocations)
                            {
                                var hazardWeight = hazardAllocation.Amount / totalHazardWeight;
                                if (hazardWeight <= 0) continue;

                                var stateAllocations = curveContainer.StateAllocations.Where(state => state.Amount > 0).ToList();
                                var curves = curveProvider.GetSeverityCurve(date, stateAllocations, hazardAllocation.Id,
                                    Convert.ToInt64(sublineInput.SublineCode));

                                var totalStateGroupWeight = curves.Sum(curve => curve.Weight);
                                curves.ForEach(curve => curve.Weight = curve.Weight / totalStateGroupWeight * hazardWeight);

                                sublineCurves.AddRange(curves);
                            }

                            resultSets.Add(new CasualtyCurveProviderResultSet
                            {
                                SegmentId = segmentId, SublineId = sublineInput.Id, Curves = sublineCurves
                            });
                            break;
                        }
                    }
                }
            }

            return resultSets;
        }
    }

    public class CasualtyCurveProviderResultSet
    {
        public string SegmentId { get; set; }
        public string SublineId { get; set; }
        public IList<SeverityCurveResult> Curves { get; set; }
    }
}

