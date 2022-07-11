namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials
{
    public class AlaeExcludedAndInAdditionToLimit : BaseCurve
    {
        public override ICalculator CreateNew()
        {
            return new AlaeExcludedAndInAdditionToLimit();
        }

        protected override double GetAlaeLimitedExpectedValueForClaimsWithoutPay(double limit, double filteredLimit, double alaeForClaimsWithoutPay)
        {
            return 0;
        }
        
        public override double GetAlaeCdfForClaimsWithoutPay(
            double effectiveLimit,
            double limit,
            double alaeForClaimsWithoutPay)
        {
            return 1;
        }
        
        protected override double GetLimitedExpectedValueComponent(double limit, double mean, double alaePercent)
        {
            return GetLimitedExpectedValue(limit, mean);
        }
        
        protected override double FilterAlaeAdjustmentFactorThroughReinsuranceAlaeTreatment(double alaeAdjustmentFactor)
        {
            // the only reinsurance alae treatments that overrides
            return 1;
        }
    }
}
