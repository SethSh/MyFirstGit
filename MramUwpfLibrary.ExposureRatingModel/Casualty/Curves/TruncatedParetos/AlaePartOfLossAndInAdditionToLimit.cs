using MramUwpfLibrary.Common.ReinsurancePerspectives;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.TruncatedParetos
{
    public class AlaePartOfLossAndInAdditionToLimit : BaseCurve
    {
        public override ICalculator CreateNew()
        {
            return new AlaePartOfLossAndInAdditionToLimit();
        }

        public override double GetEffectiveLimit(double limit, double policyLimit, double policySir,
            IReinsurancePerspectiveHandler reinsurancePerspective, double variableAlae)
        {
            return reinsurancePerspective.GetEffectiveLimit(limit/(1d + variableAlae), policyLimit, policySir);
        }
    }
}
