using System;
using MramUwpfLibrary.Common.Extensions;
using MramUwpfLibrary.Common.ReinsurancePerspectives;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials
{
    public class AlaePartOfLossAndInAdditionToLimit : BaseCurve
    {
        public override ICalculator CreateNew()
        {
            return new AlaePartOfLossAndInAdditionToLimit();
        }

        protected override double GetAlaeLimitedExpectedValueForClaimsWithoutPay(double limit, double filteredLimit, double alaeForClaimsWithoutPay)
        {
            if (alaeForClaimsWithoutPay > 0 && PolicyLimit > 0 && PolicySir.IsEpsilonEqualToZero())
            {
                return alaeForClaimsWithoutPay * (1 - Math.Exp(-limit / alaeForClaimsWithoutPay));
            }
            return 0;
        }
        
        public override double GetEffectiveLimit(double limit, IReinsurancePerspectiveHandler reinsurancePerspective, double alaePercent)
        {
            return reinsurancePerspective.GetEffectiveLimit(limit/(1 + alaePercent), PolicyLimit, PolicySir);
        }

        public override double GetAlaeCdfForClaimsWithoutPay(double effectiveLimit, double limit, double alaeForClaimsWithoutPay)
        {
            if (alaeForClaimsWithoutPay > 0 && PolicyLimit > 0 && PolicySir.IsEpsilonEqualToZero())
            {
                return 1 - Math.Exp(-1 * limit / alaeForClaimsWithoutPay);
            }
            return 1;
        }

        protected override double GetLimitedExpectedValueComponent(double limit, double mean, double alaePercent)
        {
            var alaeFactor = 1 + alaePercent; 
            return GetLimitedExpectedValue(limit, mean) * alaeFactor;
        }

        

        

        protected override bool GetAlaeProbabilityLessThanLimitIndicator(double alaeForClaimsWithoutPay)
        {
            return alaeForClaimsWithoutPay > 0 && PolicyLimit > 0 && PolicySir.IsEpsilonEqualToZero();
        }
    }
}
