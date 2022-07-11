using MramUwpfLibrary.Common.Extensions;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials
{
    public class AlaeProRataAndInAdditionToLimit : BaseCurve
    {
        public override ICalculator CreateNew()
        {
            return new AlaeProRataAndInAdditionToLimit();
        }
        
        public override double GetAlaeCdfForClaimsWithoutPay(double effectiveLimit, double policyLimit, double alaeLevForClaimsWithoutPay)
        {
            return 1;
        }

        protected override double GetAlaeLimitedExpectedValueForClaimsWithoutPay(double limit, double filteredLimit, double alaeForClaimsWithoutPay)
        {
            //changed to match legacy Cas ERM
            if (alaeForClaimsWithoutPay > 0 && PolicyLimit > 0 && PolicySir.IsEpsilonEqualToZero() && limit > 0)
            {
                return alaeForClaimsWithoutPay;
            }
            return 0;
        }

        protected override double GetLimitedExpectedValueComponent(double limit, double mean, double alaePercent)
        {
            var alaeFactor = 1 + alaePercent;
            return GetLimitedExpectedValue(limit, mean) * alaeFactor;
        }
    }
}
