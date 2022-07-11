using MramUwpfLibrary.Common.ReinsurancePerspectives;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials
{
    public interface ICalculator
    {
        double GetLayerLimitedExpectedValue(double limitTop,
            double limitBottom,
            double policyLimit,
            double policySir,
            double alaeAdjustmentFactor,
            IReinsurancePerspectiveHandler reinsurancePerspective,
            Parameters curveParameters);

        double GetLimitedExpectedValue(double limit,
            double policyLimit,
            double policySir,
            double alaeAdjustmentFactor,
            IReinsurancePerspectiveHandler reinsurancePerspective,
            Parameters curveParameters);

        double GetProbabilityLessThanLimit(
            double limit,
            double policyLimit,
            double policySir,
            double alaeAdjustmentFactor,
            IReinsurancePerspectiveHandler reinsurancePerspective,
            Parameters curveParameters);

        ICalculator CreateNew();
    }
}
