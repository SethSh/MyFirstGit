using System.Collections.Generic;
using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.ExposureRatingModel.Input;

namespace MramUwpfLibrary.ExposureRatingModel.Property
{
    internal static class PropertyCalculatorHelper
    {
        public static LossRatioResultSet GrossUpSublineLossRatio(ReinsuranceParameters reinsurance,
            PolicyAlaeTreatmentType policyAlaeTreatmentType, ISublineExposureRatingInput sublineInput,
            IList<ITotalInsuredValueItem> totalInsuredValueItems, double alaeToLossPercent)
        {
            if (sublineInput.AllocatedExposureAmount <= 0) return null;

            var sublineCalculator = new PropertySublineCalculator(
                reinsurance,
                policyAlaeTreatmentType,
                sublineInput,
                totalInsuredValueItems, 
                alaeToLossPercent);
            
            var restatedLossRatio = sublineCalculator.GrossUpLossRatio();
            return new LossRatioResultSet
            {
                Id = sublineInput.Id,
                SubjBase = sublineInput.AllocatedExposureAmount,
                LimitedLossRatio = sublineInput.LimitedLossRatio,
                GrossUpLossRatio = restatedLossRatio
            };
        }

        public static IExposureRatingResultItem SliceIntoLayer(ReinsuranceParameters reinsurance,
            GrossUpLossRatioFactors grossUpLossRatio,
            PolicyAlaeTreatmentType policyAlaeTreatmentType,
            ISublineExposureRatingInput sublineInput,
            IList<ITotalInsuredValueItem> totalInsuredValueItems,
            double alaeToLossPercent)
        {
            if (sublineInput.AllocatedExposureAmount <= 0) return new ExposureRatingResultItem { SublineId = sublineInput.Id };

            var sublineCalculator = new PropertySublineCalculator(reinsurance, policyAlaeTreatmentType,
                sublineInput, totalInsuredValueItems, alaeToLossPercent);
            
            return sublineCalculator.Calculate(grossUpLossRatio);
        }
    }
}
