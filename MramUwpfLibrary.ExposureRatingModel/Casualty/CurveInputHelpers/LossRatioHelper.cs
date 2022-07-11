using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.Common.ReinsurancePerspectives;
using MramUwpfLibrary.ExposureRatingModel.Casualty.Curves;
using MramUwpfLibrary.ExposureRatingModel.Input;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.CurveInputHelpers
{
    internal class LossRatioHelper : IHelper
    {
        private readonly PolicyAlaeTreatmentType _policyAlaeTreatment;
        private readonly ISublineExposureRatingInput _sublineInput;
        private readonly double _policyLimit;
        private readonly double _policySir;
        private readonly IReinsurancePerspectiveHandler _reinsurancePerspective;

        public LossRatioHelper(PolicyAlaeTreatmentType policyAlaeTreatment, ISublineExposureRatingInput sublineInput, 
            double policyLimit, 
            double policySir,
            IReinsurancePerspectiveHandler reinsurancePerspective)
        {
            _policyAlaeTreatment = policyAlaeTreatment;
            _sublineInput = sublineInput;
            _policyLimit = policyLimit;
            _policySir = policySir;
            _reinsurancePerspective = reinsurancePerspective;
        }

        public CurveInputs GetNumeratorCurveInputs()
        {
            return new CurveInputs
            {
                TopLimit = _reinsurancePerspective.GetLossRatioTop(_policySir, _sublineInput.LossRatioLimit),
                BottomLimit = _reinsurancePerspective.GetLossRatioBottom(_policySir),
                ReinsurancePerspective = _reinsurancePerspective,
                ReinsuranceAlaeTreatment = _sublineInput.LossRatioAlaeTreatment.MapToReinsuranceAlaeTreatmentType(),
                PolicyAlaeTreatment = _policyAlaeTreatment,
                AlaeAdjustmentFactor = _sublineInput.AlaeAdjustmentFactor
            };
        }

        public CurveInputs GetNumeratorCurveBenchmarkInputs()
        {
            return new CurveInputs
            {
                TopLimit = _reinsurancePerspective.GetLossRatioTop(_policySir, _sublineInput.LossRatioLimit),
                BottomLimit = _reinsurancePerspective.GetLossRatioBottom(_policySir),
                ReinsurancePerspective = _reinsurancePerspective,
                ReinsuranceAlaeTreatment = _sublineInput.LossRatioAlaeTreatment.MapToReinsuranceAlaeTreatmentType(),
                PolicyAlaeTreatment = _policyAlaeTreatment,
                AlaeAdjustmentFactor = 1.0d
            };
        }

        public CurveInputs GetDenominatorCurveInputs()
        {
            return new CurveInputs
            {
                TopLimit = _policyLimit + _policySir,
                BottomLimit = _policySir,
                ReinsurancePerspective = new FromGroundHandler(),
                ReinsuranceAlaeTreatment = _policyAlaeTreatment.MapToReinsuranceAlaeTreatmentTypeForLossRatio(),
                PolicyAlaeTreatment = _policyAlaeTreatment,
                AlaeAdjustmentFactor = _sublineInput.AlaeAdjustmentFactor
            };
        }

        public CurveInputs GetDenominatorCurveBenchmarkInputs()
        {
            return new CurveInputs
            {
                TopLimit = _policyLimit + _policySir,
                BottomLimit = _policySir,
                ReinsurancePerspective = new FromGroundHandler(),
                ReinsuranceAlaeTreatment = _policyAlaeTreatment.MapToReinsuranceAlaeTreatmentTypeForLossRatio(),
                PolicyAlaeTreatment = _policyAlaeTreatment,
                AlaeAdjustmentFactor = 1.0d
            };
        }
    }
}