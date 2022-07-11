using MramUwpfLibrary.Common.ReinsurancePerspectives;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.TruncatedParetos
{
    public class AlaeExcludedAndInAdditionToLimit : BaseCurve
    {
        public override ICalculator CreateNew()
        {
            return new AlaeExcludedAndInAdditionToLimit();
        }
        
        public override double GetAppropriateAlaeAmount(double alaeAmount)
        {
            return 0;
        }

        public override double GetEffectiveLimit(double limit, double policyLimit, double policySir, IReinsurancePerspectiveHandler reinsurancePerspective, double variableAlae)
        {
            return reinsurancePerspective.GetEffectiveLimit(limit, policyLimit, policySir);
        }

        protected override double FilterAlaeAdjustmentFactorThroughReinsuranceAlaeTreatment(double alaeAdjustmentFactor)
        {
            // the only reinsurance alae treatments that overrides
            return 1;
        }
    }
}
