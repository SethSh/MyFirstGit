using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.Common.ReinsurancePerspectives;
using MramUwpfLibrary.ExposureRatingModel.Casualty.Curves;
using MramUwpfLibrary.ExposureRatingModel.Input;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.CurveInputHelpers
{
    class SelectedAlaeHelper : IHelper
    {
        private readonly ISublineExposureRatingInput _input;
        private readonly double _policyLimit;
        private readonly double _policySir;
        
        public SelectedAlaeHelper(ISublineExposureRatingInput input, double policyLimit, double policySir)
        {
            _input = input;
            _policyLimit = policyLimit;
            _policySir = policySir;
        }

        public CurveInputs GetNumeratorCurveInputs()
        {
            return new CurveInputs
            {
                TopLimit = _policyLimit + _policySir,
                BottomLimit = _policySir,
                ReinsurancePerspective = new FromGroundHandler(),
                ReinsuranceAlaeTreatment = ReinsuranceAlaeTreatmentType.ProRata,
                PolicyAlaeTreatment = PolicyAlaeTreatmentType.InAdditionToLimit,
                AlaeAdjustmentFactor = _input.AlaeAdjustmentFactor
            };
        }

        public CurveInputs GetNumeratorCurveBenchmarkInputs()
        {
            return new CurveInputs
            {
                TopLimit = _policyLimit + _policySir,
                BottomLimit = _policySir,
                ReinsurancePerspective = new FromGroundHandler(),
                ReinsuranceAlaeTreatment = ReinsuranceAlaeTreatmentType.ProRata,
                PolicyAlaeTreatment = PolicyAlaeTreatmentType.InAdditionToLimit,
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
                ReinsuranceAlaeTreatment = ReinsuranceAlaeTreatmentType.ProRata,
                PolicyAlaeTreatment = PolicyAlaeTreatmentType.InAdditionToLimit,
                AlaeAdjustmentFactor = 0
            };
        }

        public CurveInputs GetDenominatorCurveBenchmarkInputs()
        {
            throw new System.NotImplementedException();
        }
    }
}