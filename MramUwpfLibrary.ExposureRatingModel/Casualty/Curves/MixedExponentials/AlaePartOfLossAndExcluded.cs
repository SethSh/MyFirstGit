using System;
using MramUwpfLibrary.Common.Extensions;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials
{
    // 99% sure this will never be called because underlying policy alae is never excluded
    internal class AlaePartOfLossAndExcluded : BaseCurve
    {
        public override ICalculator CreateNew()
        {
            return new AlaePartOfLossAndExcluded();
        }

        protected override double GetAlaeLimitedExpectedValueForClaimsWithoutPay(double limit, double forClaimsWithoutPay, double alaeForClaimsWithoutPay)
        {
            if (alaeForClaimsWithoutPay > 0 && PolicyLimit > 0 && PolicySir.IsEpsilonEqualToZero())
            {
                return alaeForClaimsWithoutPay * (1 - Math.Exp(-limit / alaeForClaimsWithoutPay));
            }
            return 0;
        }

        protected override double GetLimitedExpectedValueComponent(double limit, double mean, double alaePercent)
        {
            return GetLimitedExpectedValue(limit, mean);
        }

        public override double GetAlaeCdfForClaimsWithoutPay(double effectiveLimit, double limit, double alaeForClaimsWithoutPay)
        {
            return 1;
        }
    }
}
