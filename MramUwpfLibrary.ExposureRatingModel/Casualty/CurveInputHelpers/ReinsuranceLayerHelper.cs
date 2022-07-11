using System;
using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.Common.ReinsurancePerspectives;
using MramUwpfLibrary.ExposureRatingModel.Casualty.Curves;
using MramUwpfLibrary.ExposureRatingModel.Input;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.CurveInputHelpers
{
    internal class ReinsuranceLayerHelper : IHelper
    {
        private readonly ReinsuranceParameters _reinsurance;
        private readonly PolicyAlaeTreatmentType _policyAlaeTreatment;
        private readonly ISublineExposureRatingInput _input;
        private readonly double _policyLimit;
        private readonly double _policySir;
        

        public ReinsuranceLayerHelper(ReinsuranceParameters reinsurance, 
            PolicyAlaeTreatmentType policyAlaeTreatment,
            ISublineExposureRatingInput input, 
            double policyLimit, double policySir)
        {
            _reinsurance = reinsurance;
            _policyAlaeTreatment = policyAlaeTreatment;
            _input = input;
            _policyLimit = policyLimit;
            _policySir = policySir;
        }

        public CurveInputs GetNumeratorCurveInputs()
        {
            return new CurveInputs
            {
                TopLimit = _reinsurance.Layer.Limit + _reinsurance.Layer.Attachment,
                BottomLimit = _reinsurance.Layer.Attachment,
                ReinsurancePerspective = _reinsurance.PerspectiveHandler,
                ReinsuranceAlaeTreatment = _reinsurance.AlaeTreatment,
                PolicyAlaeTreatment = _policyAlaeTreatment,
                AlaeAdjustmentFactor = _input.AlaeAdjustmentFactor
            };
        }

        public CurveInputs GetNumeratorCurveBenchmarkInputs()
        {
            throw new NotImplementedException();
        }

        public CurveInputs GetDenominatorCurveInputs()
        {
            var reinAlaeTreatment = _policyAlaeTreatment.MapToReinsuranceAlaeTreatmentTypeForLayer(_reinsurance.AlaeTreatment);
          
            return new CurveInputs
            {
                TopLimit = _policyLimit + _policySir,
                BottomLimit = _policySir,
                ReinsurancePerspective = new FromGroundHandler(),
                ReinsuranceAlaeTreatment = reinAlaeTreatment,
                PolicyAlaeTreatment = _policyAlaeTreatment,
                AlaeAdjustmentFactor = _input.AlaeAdjustmentFactor
            };
        }

        public CurveInputs GetDenominatorCurveBenchmarkInputs()
        {
            throw new NotImplementedException();
        }
    }
}