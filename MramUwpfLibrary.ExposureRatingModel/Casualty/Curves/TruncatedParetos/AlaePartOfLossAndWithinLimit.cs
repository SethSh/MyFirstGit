using MramUwpfLibrary.Common.ReinsurancePerspectives;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.TruncatedParetos
{
    public class AlaePartOfLossAndWithinLimit : BaseCurve
    {
        public override ICalculator CreateNew()
        {
            return new AlaePartOfLossAndWithinLimit();
        }

        public override double GetEffectiveLimit(double limit, double policyLimit, double policySir,
            IReinsurancePerspectiveHandler reinsurancePerspective, double variableAlae)
        {
            var filteredLimit = reinsurancePerspective.GetEffectiveLimit(limit, policyLimit, policySir);
            filteredLimit /= (1d + variableAlae);
            return filteredLimit;
        }
    }
}
