using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.Common.ReinsurancePerspectives;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.TruncatedParetos
{
    public interface ICalculator
    {
        double GetLayerLimitedExpectedValue(
            double limitTop,
            double limitBottom,
            double policyLimit,
            double policySir,
            double alaeAdjustmentFactor,
            ReinsuranceAlaeTreatmentType reinsuranceAlaeTreatment,
            IReinsurancePerspectiveHandler reinsurancePerspective,
            IParameters curveParameters);

        double GetLev(double limit, double policyLimit, double policySir, double alaeAdjustmentFactor, 
            IReinsurancePerspectiveHandler reinsurancePerspective, IParameters curveParameters);

        double GetCdfForLimit(double limit, double policyLimit, double policySir, double alaeAdjustmentFactor, 
            IReinsurancePerspectiveHandler reinsurancePerspective, IParameters curveParameters);
        
        ICalculator CreateNew();
    }
}
