using System.Collections.Generic;
using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.ExposureRatingModel.Casualty;
using MramUwpfLibrary.ExposureRatingModel.Input;
using MramUwpfLibrary.ExposureRatingModel.Input.WorkersComp;

namespace MramUwpfLibrary.ExposureRatingModel.WorkersCompensation
{
    internal class WorkersCompCalculatorHelper
    {
        public LossRatioResultSet GrossUpSublineLossRatio(ReinsuranceParameters reinsuranceParameters,
            PolicyAlaeTreatmentType policyAlaeTreatmentType, WorkersCompSublineExposureRatingInput sublineInput,
            IList<MixedExponentialCurve> curves)
        {
            var sublineCalculator = new WorkersCompSublineCalculator(reinsuranceParameters, sublineInput, policyAlaeTreatmentType, curves);
            var grossUpLossRatio = sublineCalculator.GrossUpLossRatio();

            return new LossRatioResultSet
            {
                Id = sublineInput.Id,
                SubjBase = sublineInput.AllocatedExposureAmount,
                LimitedLossRatio = sublineInput.LimitedLossRatio,
                GrossUpLossRatio = grossUpLossRatio
            };
        }

        public IExposureRatingResultItem SliceIntoLayer(ReinsuranceParameters reinsuranceParameters,
            GrossUpLossRatioFactors grossUpLossRatio,
            PolicyAlaeTreatmentType policyAlaeTreatmentType,
            ISublineExposureRatingInput sublineInput,
            IList<MixedExponentialCurve> curves)
        {
            var sublineCalculator = new WorkersCompSublineCalculator(reinsuranceParameters, sublineInput, policyAlaeTreatmentType, curves);
            return sublineCalculator.Calculate(grossUpLossRatio);
        }

    }
}
