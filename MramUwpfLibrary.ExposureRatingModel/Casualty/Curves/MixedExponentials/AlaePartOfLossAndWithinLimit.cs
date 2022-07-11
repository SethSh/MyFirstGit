using System;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials
{
    public class AlaePartOfLossAndWithinLimit : BaseCurve
    {
        public override ICalculator CreateNew()
        {
            return new AlaePartOfLossAndWithinLimit();
        }

        protected override double GetAlaeLimitedExpectedValueForClaimsWithoutPay(double limit, double filteredLimit, double alaeForClaimsWithoutPay)
        {
            if (alaeForClaimsWithoutPay > 0 && PolicyLimit > 0)
            {
                return alaeForClaimsWithoutPay * (1 - Math.Exp(-filteredLimit / alaeForClaimsWithoutPay));
            }
            return 0;
        }
        
        public override double GetAlaeCdfForClaimsWithoutPay(double effectiveLimit, double limit, double alaeForClaimsWithoutPay)
        {
            if (alaeForClaimsWithoutPay > 0)
            {
                return 1 - Math.Exp(-1 * effectiveLimit / alaeForClaimsWithoutPay);
            }
            return 1;
        }

        public override double AdjustMeanWithAlaeWhenNecessary(double mean, double alaePercent)
        {
            return mean*(1 + alaePercent);
        }

        protected override double GetLimitedExpectedValueComponent(double limit, double mean, double alaePercent)
        {
            var alaeFactor = 1 + alaePercent;
            var meanWithAlae = mean*alaeFactor;
            return GetLimitedExpectedValue(limit, meanWithAlae);
        }
        
        public override double GetProbabilityLessThanLimitComponent(double limit, double mean, double alaePercent)
        {
            var alaeFactor = 1 + alaePercent;
            var meanWithAlae = mean*alaeFactor;
            return Math.Exp(-limit/meanWithAlae);
        }

        protected override bool GetAlaeProbabilityLessThanLimitIndicator(double alaeForClaimsWithoutPay)
        {
            return alaeForClaimsWithoutPay > 0 && PolicyLimit > 0;
        }
    }
}
