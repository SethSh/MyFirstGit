using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MramUwpfLibrary.Common.Extensions;
using MramUwpfLibrary.Common.Policies;
using MramUwpfLibrary.ExposureRatingModel.Casualty;
using MramUwpfLibrary.ExposureRatingModel.Input;
using MramUwpfLibrary.ExposureRatingModel.Input.WorkersComp;

namespace MramUwpfLibrary.ExposureRatingModel.WorkersCompensation
{
    public interface IWorkerCompExposureRatingCalculator
    {
        IEnumerable<IExposureRatingResult> Calculate(IWorkersCompExposureRatingInput input);
    }

    public class WorkerCompExposureRatingCalculator : IWorkerCompExposureRatingCalculator
    {
        private readonly ISeverityCurveProvider _curveProvider; 
        
        public WorkerCompExposureRatingCalculator(ISeverityCurveProvider curveProvider)
        {
            _curveProvider = curveProvider;
        }

        public IEnumerable<IExposureRatingResult> Calculate(IWorkersCompExposureRatingInput input)
        {
            NormalizeSublineAllocations(input);
            var curveSets = FetchCurves(input).ToList();
            var mixedExponentials = MapToMixedExponentials(input, curveSets);
            
            var lossRatioSets = new List<LossRatioResultSet>();
            foreach (var segmentInput in input.SegmentInputs)
            {
                var reinsuranceParametersWithoutLayer = new ReinsuranceParameters(input.ReinsuranceAlaeTreatment, input.ReinsurancePerspective);
                var thisMixedExponentials = mixedExponentials[segmentInput.Id];

                var segmentLossRatioSets = GrossUpSegmentLossRatios(reinsuranceParametersWithoutLayer, segmentInput, thisMixedExponentials);
                lossRatioSets.AddRange(segmentLossRatioSets);
            }
            var aggregatedLossRatio = ExposureRatingCalculatorShared.AggregateLossRatioSets(input.UseAlternative, lossRatioSets);

            var exposureRatingResults = new List<IExposureRatingResult>();
            foreach (var reinsuranceLayer in input.ReinsuranceLayers)
            {
                var reinsuranceParameters = new ReinsuranceParameters(reinsuranceLayer, input.ReinsuranceAlaeTreatment, input.ReinsurancePerspective);
                foreach (var segmentInput in input.SegmentInputs)
                {
                    var thisMixedExponentials = mixedExponentials[segmentInput.Id];

                    var segmentResultItems = SliceIntoLayer(segmentInput, reinsuranceParameters,
                        input.UseAlternative, lossRatioSets, aggregatedLossRatio, thisMixedExponentials);

                    var exposureRatingResult = new ExposureRatingResult
                    {
                        LayerId = reinsuranceLayer.Id,
                        SubmissionSegmentId = segmentInput.Id,
                        Items = segmentResultItems
                    };
                    exposureRatingResults.Add(exposureRatingResult);
                }
            }

            return exposureRatingResults;
        }

        private static IEnumerable<IExposureRatingResultItem> SliceIntoLayer(ISegmentInput segmentInput, ReinsuranceParameters reinsuranceParameters, 
            bool useAlternative, IList<LossRatioResultSet> lossRatioInputSets, GrossUpLossRatioFactors grossUpLossRatio, IList<MixedExponentialCurve> curves)
        {
            //no need to call this more than once
            var results = new List<IExposureRatingResultItem>();
            
            var isFirstResultCopyable = false;

            var firstExposureAmount = 0d;
            foreach (var sublineInput in segmentInput.SublineInputs)
            {
                IExposureRatingResultItem sliceResult;
                if (isFirstResultCopyable)
                {
                    var firstResult = results.First();
                    var factor = sublineInput.AllocatedExposureAmount / firstExposureAmount;
                    sliceResult = ExposureRatingResultItem.CopyAndAdjust(sublineInput.Id, factor, firstResult);
                }
                else
                {
                    if (sublineInput.AllocatedExposureAmount <= 0)
                    {
                        sliceResult = new ExposureRatingResultItem { SublineId = sublineInput.Id };
                    }
                    else
                    {
                        var calculatorHelper = new WorkersCompCalculatorHelper();

                        var lossRatioGoingForward = !useAlternative
                            ? grossUpLossRatio
                            : lossRatioInputSets.First(x => x.Id == sublineInput.Id).GrossUpLossRatio;

                        sliceResult = calculatorHelper.SliceIntoLayer(reinsuranceParameters,
                            lossRatioGoingForward, segmentInput.PolicyAlaeTreatment,
                            sublineInput, curves);

                        isFirstResultCopyable = true;
                        firstExposureAmount = sublineInput.AllocatedExposureAmount;
                    }
                }

                results.Add(sliceResult);
            }

            return results;
        }

        private static IEnumerable<LossRatioResultSet> GrossUpSegmentLossRatios(ReinsuranceParameters reinsurance, IWorkersCompSegmentInput segmentInput, 
            IList<MixedExponentialCurve> curves)
        {
            //workers comp doesn't need to gross up for each subline, only once
            var results = new List<LossRatioResultSet>();
            
            foreach (var sublineInput in segmentInput.SublineInputs.Where(subline => subline.AllocatedExposureAmount > 0))
            {
                LossRatioResultSet sublineGrossedUpLossRatio;
                if (!results.Any())
                {
                    var calculatorHelper = new WorkersCompCalculatorHelper();
                    var workersCompSublineInput = (WorkersCompSublineExposureRatingInput)sublineInput;

                    sublineGrossedUpLossRatio = calculatorHelper.GrossUpSublineLossRatio(reinsurance, segmentInput.PolicyAlaeTreatment,
                        workersCompSublineInput, curves);

                    results.Add(sublineGrossedUpLossRatio);
                }
                else
                {
                    var firstResult = results.First();
                    sublineGrossedUpLossRatio = LossRatioResultSet.CopyValues(sublineInput.Id, firstResult);
                }
                results.Add(sublineGrossedUpLossRatio);
            }

            return results;
        }

        private static void NormalizeSublineAllocations(IWorkersCompExposureRatingInput input)
        {
            foreach (var segmentInput in input.SegmentInputs)
            {
                var allocationTotal = segmentInput.SublineInputs.Sum(gs => gs.Allocation);
                foreach (var sublineInput in segmentInput.SublineInputs)
                {
                    sublineInput.Allocation = sublineInput.Allocation.DivideByWithTrap(allocationTotal, 0);
                    ((SublineExposureRatingInput)sublineInput).NormalizedAllocation = sublineInput.Allocation;
                }
            }
        }

        private IEnumerable<WorkersCompCurveProviderResultSetPlus> FetchCurves(IWorkersCompExposureRatingInput input)
        {
            var curveSets = new List<WorkersCompCurveProviderResultSetPlus>();
            foreach (var segmentInput in input.SegmentInputs)
            {
                var date = input.EffectiveDate;
                var segmentId = segmentInput.Id;

                var firstSubline = segmentInput.PrimaryInputs.First();
                var stateHazardAllocations = firstSubline.CurveContainer.StateHazardGroupAllocations.ToList();
                
                NormalizePremium(stateHazardAllocations);

                var curveSet = new List<SeverityCurveResult>();
                var uniqueHazardCodes = stateHazardAllocations.Select(alloc => alloc.HazardId).Distinct().ToList();

                var hazardGroupMapper = new Dictionary<long, long>();
                foreach (var hazardGroupCode in uniqueHazardCodes)
                {
                    var stateAllocations = stateHazardAllocations
                        .Where(alloc => alloc.HazardId == hazardGroupCode && alloc.Value > 0)
                        .Select(item => new ProfileAllocation {Id = item.StateId, Amount = item.Value})
                        .ToList();

                    if (!stateAllocations.Any()) continue;

                    var stateGroupCurves = _curveProvider.GetSeverityCurve(date, stateAllocations, hazardGroupCode, 4);
                    foreach (var severityCurveResult in stateGroupCurves)
                    {
                        hazardGroupMapper.Add(severityCurveResult.CurveId, hazardGroupCode);
                    }

                    curveSet.AddRange(stateGroupCurves);
                }
                
                curveSets.Add(new WorkersCompCurveProviderResultSetPlus
                {
                    SegmentId = segmentId,
                    Curves = curveSet,
                    HazardGroupMapper = hazardGroupMapper
                });
            }

            return curveSets;
        }

        private static IDictionary<string, IList<MixedExponentialCurve>> MapToMixedExponentials(IWorkersCompExposureRatingInput input,
            IList<WorkersCompCurveProviderResultSetPlus> curveSets)
        {
            const double unlimitedPolicyLimit = double.MaxValue; 
            
            var dict = new Dictionary<string, IList<MixedExponentialCurve>>();
            foreach (var segmentInput in input.SegmentInputs)
            {
                var segmentId = segmentInput.Id;
                var curveSet = curveSets.Single(item => item.SegmentId == segmentId);

                var stateHazardGroupAllocations = segmentInput.PrimaryInputs.First().CurveContainer.StateHazardGroupAllocations.ToList();
                var stateAttachmentAllocations = segmentInput.PrimaryInputs.First().CurveContainer.StateAttachmentAllocations.ToList();
                var minnesotaRetentionValue = segmentInput.PrimaryInputs.First().CurveContainer.MinnesotaRetentionValue;
            
                var curves = new List<MixedExponentialCurve>();
                foreach (var curve in curveSet.Curves.Where(item => item.Weight> 0))
                {
                    //in practice the wc mapping of state to state-group is a one-to-one relationship
                    //when the single state in the state-group is minnesota, then policy limit is different
                    if (curve.States.Count != 1) throw new InvalidDataException("WC can't map state to state-group must have on-to-one relationship");
                    var curveState = curve.States.First();
                    //todo - this should get moved to a provider
                    var isMinnesota = curveState == 25;
                    
                    var policySet = new PolicySet();
                    var hazardGroupId = curveSet.HazardGroupMapper[curve.CurveId];

                    var stateIdsWithPremium = stateHazardGroupAllocations
                        .Where(item => item.HazardId == hazardGroupId 
                                       && curve.States.Contains(item.StateId)
                                       && item.Value > 0)
                        .Select(item => item.StateId);

                    var premiumAllocations = stateHazardGroupAllocations.Where(item => stateIdsWithPremium.Contains(item.StateId) && item.HazardId == hazardGroupId).ToDictionary(item => item.StateId);
                    var totalPremiumAllocation = premiumAllocations.Values.Sum(item => item.Value);
                    var attachmentAllocations = stateAttachmentAllocations.Where(item => stateIdsWithPremium.Contains(item.StateId)).ToList();

                    if (attachmentAllocations.Any())
                    {
                        IList<AttachmentAndValue> attachments;
                        if (attachmentAllocations.Select(item => item.StateId).Distinct().Count() > 1)
                        {
                            var premiumWeightedAttachmentAllocations = new List<WorkersCompStateAttachmentAllocation>();
                            foreach (var attachmentAllocation in attachmentAllocations)
                            {
                                if (!premiumAllocations.ContainsKey(attachmentAllocation.StateId)) continue;

                                var premiumWeight = premiumAllocations[attachmentAllocation.StateId].Value.DivideByWithTrap(totalPremiumAllocation);
                                premiumWeightedAttachmentAllocations.Add(new WorkersCompStateAttachmentAllocation
                                {
                                    StateId = attachmentAllocation.StateId,
                                    Attachment = attachmentAllocation.Attachment,
                                    Value = attachmentAllocation.Value * premiumWeight
                                });
                            }

                            attachments = premiumWeightedAttachmentAllocations.GroupBy(item => item.Attachment).Select(item =>
                                new AttachmentAndValue
                                {
                                    Attachment = item.Key,
                                    Value = item.Sum(t => t.Value)
                                }).ToList();
                        }
                        else
                        {
                            attachments = attachmentAllocations.Select(item =>
                                new AttachmentAndValue
                                {
                                    Attachment = item.Attachment,
                                    Value = item.Value
                                }).ToList();
                        }

                        var excessPercent = attachments.Sum(item => item.Value);
                        if (excessPercent > 1) throw new InvalidDataException("State Attachment percents sum to greater than 1");

                        var primaryPercent = 1d - excessPercent;
                        if (primaryPercent > 0)
                        {
                            attachments.Add(new AttachmentAndValue {Attachment = 0, Value = primaryPercent});
                        }

                        foreach (var attachmentAndValue in attachments)
                        {
                            policySet.Add(isMinnesota
                                ? new Policy
                                {
                                    Limit = Math.Max(0, minnesotaRetentionValue - attachmentAndValue.Attachment),
                                    Sir = attachmentAndValue.Attachment,
                                    Amount = attachmentAndValue.Value
                                }
                                : new Policy
                                {
                                    Limit = unlimitedPolicyLimit,
                                    Sir = attachmentAndValue.Attachment,
                                    Amount = attachmentAndValue.Value
                                });
                        }
                    }
                    else
                    {
                        policySet.Add(isMinnesota
                        ? new Policy{ Limit = minnesotaRetentionValue, Sir = 0d, Amount = 1d }
                        : new Policy { Limit = unlimitedPolicyLimit, Sir = 0d, Amount = 1d });
                    }

                    curves.Add(new MixedExponentialCurve
                    {
                        Id = curve.CurveId.ToString(),
                        Weight = curve.Weight,
                        CurveParameters = curve.MixedExponentialCurveParameters,
                        PolicySet = policySet,
                        ForceWithinLimits = isMinnesota
                    });
                }
                dict.Add(segmentId, curves);
            }
            return dict;
        }

        private static void NormalizePremium(List<WorkersCompStateHazardAllocation> stateHazardAllocations)
        {
            var totalStateHazardValue = stateHazardAllocations.Sum(alloc => alloc.Value);
            foreach (var workersCompStateHazardAllocation in stateHazardAllocations)
            {
                workersCompStateHazardAllocation.Value = workersCompStateHazardAllocation.Value.DivideByWithTrap(totalStateHazardValue);
            }
        }

        private class AttachmentAndValue
        {
            public double Attachment { get; set; }
            public double Value { get; set; }
        }
    }

    public class WorkersCompCurveProviderResultSet : CasualtyCurveProviderResultSet
    {
        public IEnumerable<string> StateIds { get; set; }
    }

    public class WorkersCompCurveProviderResultSetPlus : WorkersCompCurveProviderResultSet
    {
        public IDictionary<long, long> HazardGroupMapper { get; set; }
    }
}