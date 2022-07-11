using MramUwpfLibrary.Common.ReinsurancePerspectives;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.TruncatedParetos
{
    public class AlaeProRataAndInAdditionToLimit : BaseCurve
    {
        public override ICalculator CreateNew()
        {
            return new AlaeProRataAndInAdditionToLimit();
        }

        public override double GetEffectiveLimit(double limit, double policyLimit, double policySir, IReinsurancePerspectiveHandler reinsurancePerspective, double variableAlae)
        {
            return reinsurancePerspective.GetEffectiveLimit(limit, policyLimit, policySir);
        }
    }
}
