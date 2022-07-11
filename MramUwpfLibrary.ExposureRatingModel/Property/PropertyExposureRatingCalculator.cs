using System;
using System.Collections.Generic;
using System.Linq;
using MramUwpfLibrary.Common.Extensions;
using MramUwpfLibrary.ExposureRatingModel.Input;
using MramUwpfLibrary.ExposureRatingModel.Input.Property;

namespace MramUwpfLibrary.ExposureRatingModel.Property
{
    public interface IPropertyExposureRatingCalculator
    {
        IEnumerable<IExposureRatingResult> Calculate(IPropertyExposureRatingInput input);
    }

    public class PropertyExposureRatingCalculator : IPropertyExposureRatingCalculator
    {
        private readonly IPropertyCurveProvider _curveProvider;
        private readonly IPropertyCurveFactorProvider _curveCurveFactorProvider;
        private readonly IPropertyAlaeToLossProvider _propertyAlaeToLossProvider;

        public PropertyExposureRatingCalculator(IPropertyCurveProvider curveProvider,
            IPropertyCurveFactorProvider curveCurveFactorProvider,
            IPropertyAlaeToLossProvider propertyAlaeToLossProvider)
        {
            _curveProvider = curveProvider;
            _curveCurveFactorProvider = curveCurveFactorProvider;
            _propertyAlaeToLossProvider = propertyAlaeToLossProvider;
        }

        public IEnumerable<IExposureRatingResult> Calculate(IPropertyExposureRatingInput input)
        {
            NormalizeSublineAllocations(input);

            var curveSets = FetchCurves(input).ToList();
            var sublineAlaeToLossPercents = FetchAlaeToLossPercents(input);

            var lossRatioSets = new List<LossRatioResultSet>();
            foreach (var segmentInput in input.SegmentInputs)
            {
                var reinsuranceParametersWithoutLayer = new ReinsuranceParameters(input.ReinsuranceAlaeTreatment, input.ReinsurancePerspective);
                var segmentLossRatioSets = GrossUpSegmentLossRatios(reinsuranceParametersWithoutLayer, segmentInput, curveSets, sublineAlaeToLossPercents);
                lossRatioSets.AddRange(segmentLossRatioSets);
            }
            var aggregatedLossRatio = ExposureRatingCalculatorShared.AggregateLossRatioSets(input.UseAlternative, lossRatioSets);

            var results = new List<IExposureRatingResult>();
            foreach (var reinsuranceLayer in input.ReinsuranceLayers)
            {
                var reinsuranceParameters = new ReinsuranceParameters(reinsuranceLayer, input.ReinsuranceAlaeTreatment, input.ReinsurancePerspective);
                foreach (var segmentInput in input.SegmentInputs)
                {
                    var segmentResultItems = SliceIntoLayer(segmentInput, curveSets, sublineAlaeToLossPercents, reinsuranceParameters,
                        input.UseAlternative, lossRatioSets, aggregatedLossRatio);

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
            IList<PropertyCurveSet> curveSets,
            IDictionary<long, double> sublineAlaeToLossPercents, 
            ReinsuranceParameters reinsurance,
            bool useAlternative, IList<LossRatioResultSet> lossRatioResultSet, GrossUpLossRatioFactors grossUpLossRatio)
        {
            var results = new List<IExposureRatingResultItem>();
            foreach (var sublineInput in segmentInput.SublineInputs)
            {
                var totalInsuredValueItems = curveSets.Single(curve => curve.SegmentId == segmentInput.Id && curve.SublineId == sublineInput.Id).TotalInsuredValueItems;
                var sublineAlaeToLossPercent = sublineAlaeToLossPercents[Convert.ToInt64(sublineInput.SublineCode)];

                var lossRatioGoingForward = !useAlternative
                    ? grossUpLossRatio
                    : lossRatioResultSet.First(x => x.Id == sublineInput.Id).GrossUpLossRatio;
                
                var restatedLossRatioInput = PropertyCalculatorHelper.SliceIntoLayer(reinsurance,
                    lossRatioGoingForward, segmentInput.PolicyAlaeTreatment, sublineInput,
                    totalInsuredValueItems,
                    sublineAlaeToLossPercent);

                if (restatedLossRatioInput != null)
                {
                    results.Add(restatedLossRatioInput);
                }
            }

            return results;
        }

        private static IEnumerable<LossRatioResultSet> GrossUpSegmentLossRatios(ReinsuranceParameters reinsurance, ISegmentInput segmentInput, 
            IList<PropertyCurveSet> curveSets, IDictionary<long, double> sublineAlaeToLossPercents)
        {
            var results = new List<LossRatioResultSet>();
            foreach (var sublineInput in segmentInput.SublineInputs)
            {
                var totalInsuredValueItems = curveSets.Single(curve => curve.SegmentId == segmentInput.Id && curve.SublineId == sublineInput.Id).TotalInsuredValueItems;
                var sublineAlaeToLossPercent = sublineAlaeToLossPercents[Convert.ToInt64(sublineInput.SublineCode)];

                var sublineGrossedUpLossRatio = PropertyCalculatorHelper.GrossUpSublineLossRatio(reinsurance, segmentInput.PolicyAlaeTreatment,
                    sublineInput, totalInsuredValueItems, sublineAlaeToLossPercent);

                if (sublineGrossedUpLossRatio != null)
                {
                    results.Add(sublineGrossedUpLossRatio);
                }
            }

            return results;

        }

        private IDictionary<long, double> FetchAlaeToLossPercents(IPropertyExposureRatingInput input)
        {
            var dict = new Dictionary<long, double>();
            foreach (var segmentInput in input.SegmentInputs)
            {
                foreach (var sublineInput in segmentInput.PrimaryInputs)
                {
                    var sublineCode = Convert.ToInt64(sublineInput.SublineCode);
                    if (!dict.ContainsKey(sublineCode))
                    {
                        // ReSharper disable once PossibleInvalidOperationException
                        dict.Add(sublineCode, _propertyAlaeToLossProvider.GetPropertyAlaeToLossFactor(input.EffectiveDate, sublineCode).Value);
                    }
                }
            }

            return dict;
        }

        internal IEnumerable<PropertyCurveSet> FetchCurves(IPropertyExposureRatingInput input)
        {
            var curveSets = new List<PropertyCurveSet>();
            foreach (var segmentInput in input.SegmentInputs)
            {
                foreach (var sublineInput in segmentInput.PrimaryInputs)
                {
                    var curveSet = new PropertyCurveSet {SegmentId = segmentInput.Id, SublineId = sublineInput.Id};
                    var curveContainer = sublineInput.CurveContainer;

                    if (curveContainer.OccupancyTypeAllocations.Any())
                    {
                        curveSet.TotalInsuredValueItems = FetchCommercialCurves(input.EffectiveDate,
                            curveContainer.TotalInsuredValueAllocations.ToList(),
                            curveContainer.OccupancyTypeAllocations.ToList(),
                            curveContainer.ConstructionTypeAllocations.ToList(),
                            curveContainer.ProtectionClassAllocations.ToList());
                    }
                    else
                    {
                        curveSet.TotalInsuredValueItems = FetchPersonalCurves(input.EffectiveDate,
                            curveContainer.TotalInsuredValueAllocations.ToList(),
                            curveContainer.ConstructionTypeAllocations.ToList(),
                            curveContainer.ProtectionClassAllocations.ToList());
                    }

                    curveSets.Add(curveSet);
                }
            }

            return curveSets;
        }

        internal IList<ITotalInsuredValueItem> FetchPersonalCurves(DateTime date, 
            IList<ITotalInsuredValueAllocation> totalInsuredValueAllocations,
            IList<IProfileAllocation> constructionTypeAllocations,
            IList<IProfileAllocation> protectionClassAllocations)
        {
            var insuredValueTotal = totalInsuredValueAllocations.Sum(alloc => alloc.Amount);
            var constructionTotal = constructionTypeAllocations.Sum(alloc => alloc.Amount);
            var protectionTotal = protectionClassAllocations.Sum(alloc => alloc.Amount);
            var tivCap = 2.5;  // ToDo: This should be stored in DB and retrieved via a provider.

            var totalInsuredValueItems = new List<ITotalInsuredValueItem>();
            foreach (var totalInsuredValue in totalInsuredValueAllocations)
            {
                var totalInsuredValueItem = new TotalInsuredValueItem
                {
                    TotalInsuredValue = totalInsuredValue.TotalInsuredValue,
                    Share = totalInsuredValue.Share,
                    Limit = Math.Min(totalInsuredValue.Limit ?? double.MaxValue, totalInsuredValue.TotalInsuredValue * tivCap),
                    Attachment = totalInsuredValue.Attachment,
                    Weight = totalInsuredValue.Amount.DivideByWithTrap(insuredValueTotal, 0),
                };

                if (totalInsuredValueItem.Weight.IsEqualToZero()) continue;

                foreach (var construction in constructionTypeAllocations)
                {
                    var constructionWeight = construction.Amount.DivideByWithTrap(constructionTotal, 0);
                    if (constructionWeight.IsEqualToZero()) continue;

                    foreach (var protection in protectionClassAllocations)
                    {
                        var protectionWeight = protection.Amount.DivideByWithTrap(protectionTotal, 0);
                        if (protectionWeight.IsEqualToZero()) continue;

                        var curve = _curveProvider.GetPersonalPropertyCurve(date, totalInsuredValue.TotalInsuredValue,
                            construction.Id, protection.Id);
                        var processor = new PropertyCurveProcessor(curve)
                        {
                            Weight = constructionWeight * protectionWeight
                        };
                        totalInsuredValueItem.PropertyCurveProcessors.Add(processor);
                    }
                }
                totalInsuredValueItems.Add(totalInsuredValueItem);
            }
            return totalInsuredValueItems;
        }


        internal IList<ITotalInsuredValueItem> FetchCommercialCurves(DateTime date,
            IList<ITotalInsuredValueAllocation> totalInsuredValueAllocations,
            IList<IProfileAllocation> occupancyAllocations,
            IList<IProfileAllocation> constructionTypeAllocations,
            IList<IProfileAllocation> protectionClassAllocations)
        {
            var curveFactors = GetCurveFactors(date, constructionTypeAllocations, protectionClassAllocations);
            var tivCap = 1.5;  // ToDo: This should be stored in DB and retrieved via a provider.

            var insuredValueTotal = totalInsuredValueAllocations.Sum(alloc => alloc.Amount);
            var occupancyTotal = occupancyAllocations.Sum(alloc => alloc.Amount);

            var totalInsuredValueItems = new List<ITotalInsuredValueItem>();
            foreach (var totalInsuredValue in totalInsuredValueAllocations)
            {
                var totalInsuredValueItem = new TotalInsuredValueItem
                {
                    TotalInsuredValue = totalInsuredValue.TotalInsuredValue,
                    Share = totalInsuredValue.Share,
                    Limit = Math.Min(totalInsuredValue.Limit ?? double.MaxValue, totalInsuredValue.TotalInsuredValue * tivCap),
                    Attachment = totalInsuredValue.Attachment,
                    Weight = totalInsuredValue.Amount.DivideByWithTrap(insuredValueTotal, 0),
                };

                if (totalInsuredValueItem.Weight.IsEqualToZero()) continue;
                
                foreach (var occupancy in occupancyAllocations)
                {
                    var occupancyWeight = occupancy.Amount.DivideByWithTrap(occupancyTotal, 0);
                    if (occupancyWeight.IsEqualToZero()) continue;
                    
                    var curve = _curveProvider.GetCommercialPropertyCurve(date, 
                        totalInsuredValue.TotalInsuredValue, 
                        occupancy.Id);
                    
                    
                    foreach (var curveFactor in curveFactors)
                    {
                        var curveClone = (IPropertyCurve)curve.Clone();
                        curveClone.MixedExponentialCurveParameters.Means.MultiplyInPlace(curveFactor.MultiplicativeFactor);
                        
                        var processor = new PropertyCurveProcessor(curveClone)
                        {
                            Weight = occupancyWeight * curveFactor.Weight
                        };
                        totalInsuredValueItem.PropertyCurveProcessors.Add(processor);
                    }
                }
                totalInsuredValueItems.Add(totalInsuredValueItem);
            }
            return totalInsuredValueItems;
        }

        private List<CurveFactor> GetCurveFactors(DateTime date, IList<IProfileAllocation> constructionTypeAllocations, IList<IProfileAllocation> protectionClassAllocations)
        {
            var constructionTotal = constructionTypeAllocations.Sum(alloc => alloc.Amount);
            var protectionTotal = protectionClassAllocations.Sum(alloc => alloc.Amount);

            var curveFactors = new List<CurveFactor>();
            foreach (var construction in constructionTypeAllocations)
            {
                var constructionWeight = construction.Amount.DivideByWithTrap(constructionTotal, 0);
                if (constructionWeight.IsEqualToZero()) continue;

                foreach (var protection in protectionClassAllocations)
                {
                    var protectionWeight = protection.Amount.DivideByWithTrap(protectionTotal, 0);
                    if (protectionWeight.IsEqualToZero()) continue;

                    var propertyFactor = _curveCurveFactorProvider.GetPropertyFactor(date, protection.Id, construction.Id);
                    if (!propertyFactor.HasValue)
                    {
                        const string message = "Can't get property curve factor";
                        throw new ArgumentNullException(message);
                    }

                    curveFactors.Add(new CurveFactor
                    {
                        MultiplicativeFactor = propertyFactor.Value,
                        Weight = constructionWeight * protectionWeight
                    });
                }
            }

            return curveFactors;
        }
        
        private static void NormalizeSublineAllocations(IPropertyExposureRatingInput input)
        {
            foreach (var segmentInput in input.SegmentInputs)
            {
                var allocationTotal = segmentInput.SublineInputs.Sum(gs => gs.Allocation);
                foreach (var sublineInput in segmentInput.SublineInputs)
                {
                    sublineInput.Allocation = sublineInput.Allocation.DivideByWithTrap(allocationTotal, 0);
                    ((SublineExposureRatingInput) sublineInput).NormalizedAllocation = sublineInput.Allocation;
                }
            }
        }

        private class CurveFactor
        {
            public double MultiplicativeFactor { get; set; }
            public double Weight { get; set; }
        }
    }

    internal static class DoubleExtensions
    {
        public static void MultiplyInPlace(this double[] array, double multiplier)
        {
            for (var row = 0;  row < array.Length; row++)
            {
                array[row] *= multiplier;
            }
        }
    }
}
