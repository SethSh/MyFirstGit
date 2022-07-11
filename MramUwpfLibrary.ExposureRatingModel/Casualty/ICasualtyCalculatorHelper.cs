using System.Collections.Generic;
using System.Linq;
using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.Common.Extensions;
using MramUwpfLibrary.ExposureRatingModel.Casualty.ResultAggregators;
using MramUwpfLibrary.ExposureRatingModel.Input;
using MramUwpfLibrary.ExposureRatingModel.Input.Casualty;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty
{
    internal interface ICasualtyCalculatorHelper
    {
        LossRatioResultSet GrossUpSublineLossRatio(ReinsuranceParameters reinsurance,
            PolicyAlaeTreatmentType policyAlaeTreatmentType,
            ISublineExposureRatingInput sublineInput, 
            IList<CasualtyCurveProviderResultSet> curveResultSets);

        IExposureRatingResultItem SliceIntoLayer(ReinsuranceParameters reinsurance,
            IList<LossRatioResultSet> restatedLossRatioInputs,
            bool useAlternative,
            GrossUpLossRatioFactors grossUpLossRatio,
            PolicyAlaeTreatmentType policyAlaeTreatmentType, 
            ISublineExposureRatingInput sublineInput, 
            IList<CasualtyCurveProviderResultSet> curveResultSets);

        bool Handles(ISublineExposureRatingInput sublineInput);
    }

    internal static class CalculatorHelperFactory
    {
        internal static ICasualtyCalculatorHelper Create(ISublineExposureRatingInput sublineInput)
        {
            var helpers = new List<ICasualtyCalculatorHelper> {new PrimaryCasualtyCalculatorHelper(), new UmbrellaCasualtyCalculatorHelper()};
            return helpers.First(helper => helper.Handles(sublineInput));
        }
    }

    internal class PrimaryCasualtyCalculatorHelper : ICasualtyCalculatorHelper
    {
        public LossRatioResultSet GrossUpSublineLossRatio(ReinsuranceParameters reinsurance,
            PolicyAlaeTreatmentType policyAlaeTreatmentType, ISublineExposureRatingInput sublineInput,
            IList<CasualtyCurveProviderResultSet> curveResultSets)
        {
            var primarySublineInput = (CasualtyPrimarySublineExposureRatingInput)sublineInput;
            if (primarySublineInput.AllocatedExposureAmount <= 0) return null;

            var sublineCalculator = new CasualtySublineCalculator(
                reinsurance,
                policyAlaeTreatmentType, 
                primarySublineInput, 
                curveResultSets.First());
            sublineCalculator.SetCurves();

            var restateLossRatio = sublineCalculator.GrossUpLossRatio();
            return new LossRatioResultSet
            {
                Id = primarySublineInput.Id,
                SubjBase = primarySublineInput.AllocatedExposureAmount,
                LimitedLossRatio = primarySublineInput.LimitedLossRatio,
                GrossUpLossRatio = restateLossRatio
            };
        }

        public IExposureRatingResultItem SliceIntoLayer(ReinsuranceParameters reinsurance,
            IList<LossRatioResultSet> restatedLossRatioInputs,
            bool useAlternative,
            GrossUpLossRatioFactors grossUpLossRatio,
            PolicyAlaeTreatmentType policyAlaeTreatmentType,
            ISublineExposureRatingInput sublineInput, 
            IList<CasualtyCurveProviderResultSet> curveResultSets)
        {
            var primarySublineInput = (CasualtyPrimarySublineExposureRatingInput)sublineInput;
            if (primarySublineInput.AllocatedExposureAmount <= 0) return new ExposureRatingResultItem {SublineId = sublineInput.Id};
            
            var sublineCalculator = new CasualtySublineCalculator(reinsurance, policyAlaeTreatmentType,
                primarySublineInput, curveResultSets.First());
            sublineCalculator.SetCurves();

            return !useAlternative
                ? sublineCalculator.Calculate(grossUpLossRatio)
                : sublineCalculator.Calculate(restatedLossRatioInputs.First(x => x.Id == primarySublineInput.Id).GrossUpLossRatio);
        }

        public bool Handles(ISublineExposureRatingInput sublineInput)
        {
            return sublineInput is CasualtyPrimarySublineExposureRatingInput;
        }
    }

    internal class UmbrellaCasualtyCalculatorHelper : ICasualtyCalculatorHelper
    {
        public LossRatioResultSet GrossUpSublineLossRatio(ReinsuranceParameters reinsurance,
            PolicyAlaeTreatmentType policyAlaeTreatmentType, 
            ISublineExposureRatingInput sublineInput, 
            IList<CasualtyCurveProviderResultSet> curveResultSets)
        {
            var umbrellaSublineInput = (CasualtyUmbrellaSublineExposureRatingInput)sublineInput;
            if (umbrellaSublineInput.AllocatedExposureAmount <= 0) return null;

            var restatedLossRatioInputs = new List<LossRatioResultSet>();
            foreach (var umbrellaAlloc in umbrellaSublineInput.UmbrellaTypeAllocations
                .Where(umbrellaTypeAllocation => umbrellaTypeAllocation.IsPersonal == umbrellaSublineInput.IsPersonal))
            {
                if (umbrellaAlloc.Allocation.IsEqualToZero() || umbrellaSublineInput.NormalizedAllocation.IsEqualToZero()) continue;

                //calculator doesn't know umbrella allocation percent
                var primaryInput = MapToPrimaryInput(umbrellaSublineInput, umbrellaAlloc);
                var sublineCalculator = new CasualtySublineCalculator(reinsurance, policyAlaeTreatmentType,
                    primaryInput, curveResultSets.Single(curveSet => curveSet.SublineId == primaryInput.Id));
                sublineCalculator.SetCurves();

                var restatedLossRatio = sublineCalculator.GrossUpLossRatio();


                restatedLossRatioInputs.Add(new LossRatioResultSet
                {
                    //this is where umbrella allocation percent is used
                    SubjBase = primaryInput.AllocatedExposureAmount * umbrellaAlloc.NormalizedAllocation,
                    LimitedLossRatio = primaryInput.LimitedLossRatio,
                    GrossUpLossRatio = restatedLossRatio
                });
            }

            if (!restatedLossRatioInputs.Any()) return null;

            var aggregatedGrossUp = ExposureRatingCalculatorShared.AggregateLossRatioSets(restatedLossRatioInputs);
            return new LossRatioResultSet
            {
                Id = umbrellaSublineInput.Id,
                SubjBase = umbrellaSublineInput.AllocatedExposureAmount,
                LimitedLossRatio = sublineInput.LimitedLossRatio,
                GrossUpLossRatio = aggregatedGrossUp
            };
        }

        

        public IExposureRatingResultItem SliceIntoLayer(ReinsuranceParameters reinsurance,
            IList<LossRatioResultSet> restatedLossRatioInputs, 
            bool useAlternative,
            GrossUpLossRatioFactors grossUpLossRatio, PolicyAlaeTreatmentType policyAlaeTreatmentType,
            ISublineExposureRatingInput sublineInput, IList<CasualtyCurveProviderResultSet> curveResultSets)
        {
            var umbrellaSublineInput = (CasualtyUmbrellaSublineExposureRatingInput)sublineInput;

            var sublineResults = new List<ExposureRatingResultItem>();
            foreach (var umbrellaTypeAllocation in umbrellaSublineInput.UmbrellaTypeAllocations.Where(
                allocation => allocation.IsPersonal == umbrellaSublineInput.IsPersonal))
            {
                if (umbrellaTypeAllocation.Allocation.IsEqualToZero() || umbrellaSublineInput.NormalizedAllocation.IsEqualToZero())
                {
                    sublineResults.Add(new ExposureRatingResultItem {SublineId = umbrellaSublineInput.Id});
                    continue;
                }

                //calculator doesn't know umbrella allocation percent
                var primaryInput = MapToPrimaryInput(umbrellaSublineInput, umbrellaTypeAllocation);
                var sublineCalculator = new CasualtySublineCalculator(reinsurance, 
                    policyAlaeTreatmentType, 
                    primaryInput, 
                    curveResultSets.Single(curveSet => curveSet.SublineId == primaryInput.Id));

                sublineCalculator.SetCurves();

                var exposureRatingResult = !useAlternative
                    ? sublineCalculator.Calculate(grossUpLossRatio)
                    : sublineCalculator.Calculate(restatedLossRatioInputs.First(x => x.Id == umbrellaSublineInput.Id).GrossUpLossRatio);
                //this is where umbrella allocation percent is used
                exposureRatingResult.Weight = primaryInput.AllocatedExposureAmount * umbrellaTypeAllocation.NormalizedAllocation;

                sublineResults.Add(exposureRatingResult);
            }

            if (!sublineResults.Any()) return new ExposureRatingResultItem();

            if (sublineResults.Count > 1)
            {
                var aggregate = ExposureRatingResultAggregator.Aggregate(sublineResults);
                aggregate.SublineId = umbrellaSublineInput.Id;

                aggregate.CheckForNan();
                return aggregate;
            }

            sublineResults.First().CheckForNan();
            return sublineResults.First();
        }

        public bool Handles(ISublineExposureRatingInput sublineInput)
        {
            return sublineInput is CasualtyUmbrellaSublineExposureRatingInput;
        }

        private static ICasualtyPrimarySublineInput MapToPrimaryInput(ISublineExposureRatingInput sublineInput, UmbrellaTypeAllocation umbrellaTypeAllocation)
        {
            return new CasualtyPrimarySublineExposureRatingInput(sublineInput.Id)
            {
                GroupId = sublineInput.GroupId,
                GroupExposureAmount = sublineInput.GroupExposureAmount,
                NormalizedAllocation = ((SublineExposureRatingInput)sublineInput).NormalizedAllocation,
                AlaeAdjustmentFactor = sublineInput.AlaeAdjustmentFactor,
                LimitedLossRatio = sublineInput.LimitedLossRatio,
                LossRatioAlaeTreatment = sublineInput.LossRatioAlaeTreatment,
                LossRatioLimit = sublineInput.LossRatioLimit,

                CasualtyCurveContainer = new CasualtyCurveSetContainer
                {
                    PolicySet = umbrellaTypeAllocation.CasualtyCurveSetContainer.PolicySet
                }
            };
        }
        
    }
}
