using System;
using System.Linq;
using MramUwpfLibrary.Common.Extensions;
using MramUwpfLibrary.Common.ReinsurancePerspectives;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials
{
    public abstract class BaseCurve : ICalculator
    {
        protected double PolicyLimit { get; set; }
        protected double PolicySir { get; set; }

        public double GetLayerLimitedExpectedValue(
            double limitTop,
            double limitBottom,
            double policyLimit,
            double policySir,
            double alaeAdjustmentFactor,
            IReinsurancePerspectiveHandler reinsurancePerspective,
            Parameters curveParameters)
        {
            var limitedExpectedValueTop = GetLimitedExpectedValue(limitTop, policyLimit, policySir, alaeAdjustmentFactor,
                reinsurancePerspective, curveParameters);

            var limitedExpectedValueBottom = GetLimitedExpectedValue(limitBottom, policyLimit, policySir, alaeAdjustmentFactor,
                reinsurancePerspective, curveParameters);

            return limitedExpectedValueTop - limitedExpectedValueBottom;
        }

        public double GetLimitedExpectedValue(
            double limit,
            double policyLimit,
            double policySir,
            double alaeAdjustmentFactor,
            IReinsurancePerspectiveHandler reinsurancePerspective,
            Parameters curveParameters)
        {
            PolicyLimit = policyLimit;
            PolicySir = policySir;

            var alaeForClaimsWithoutPay = curveParameters.AlaeForClaimsWithoutPay * alaeAdjustmentFactor;
            var alaePercents = curveParameters.AlaePercents.Select(x => x * alaeAdjustmentFactor).ToArray();

            var limitedExpectedValue = 0d;
            var effectiveLimit = 0d;
            var parameterSetCount = curveParameters.Means.Length;
            for (var i = 0; i < parameterSetCount; i++)
            {
                var mean = curveParameters.Means[i];
                var weight = curveParameters.Weights[i];
                var alaePercent = alaePercents[i];

                effectiveLimit = GetEffectiveLimit(limit, reinsurancePerspective, alaePercent);
                if (effectiveLimit.IsEpsilonEqualToZero()) continue;

                var thisLimitedExpectedValue = GetLimitedExpectedValueComponent(effectiveLimit, mean, alaePercent);
                limitedExpectedValue += thisLimitedExpectedValue*weight;
            }

            var alaeLevForClaimsWithoutPay =
                GetAlaeLimitedExpectedValueForClaimsWithoutPay(limit, effectiveLimit, alaeForClaimsWithoutPay);

            return alaeLevForClaimsWithoutPay*curveParameters.ProbabilityOfNoLoss
                   + limitedExpectedValue*(1 - curveParameters.ProbabilityOfNoLoss);
        }
        
        public double GetProbabilityLessThanLimit(
            double limit,
            double policyLimit,
            double policySir,
            double alaeAdjustmentFactor,
            IReinsurancePerspectiveHandler reinsurancePerspective,
            Parameters curveParameters)
        {
            PolicyLimit = policyLimit;
            PolicySir = policySir;

            var parameters = curveParameters;
            alaeAdjustmentFactor = FilterAlaeAdjustmentFactorThroughReinsuranceAlaeTreatment(alaeAdjustmentFactor);

            var alaeForClaimsWithoutPay = curveParameters.AlaeForClaimsWithoutPay * alaeAdjustmentFactor;
            var alaePercents = curveParameters.AlaePercents.Select(x => x * alaeAdjustmentFactor).ToArray();

            var probabilityLessThanLimit = 0d;
            var filteredLimit = 0d;
            var parameterSetCount = parameters.Means.Length;
            for (var i = 0; i < parameterSetCount; i++)
            {
                var mean = parameters.Means[i];
                var weight = parameters.Weights[i];
                var alaePercent = alaePercents[i];

                filteredLimit = GetEffectiveLimit(limit, reinsurancePerspective, alaePercent);
                var thisProbabilityLessThanLimit = GetProbabilityLessThanLimitComponent(filteredLimit, mean, alaePercent);
                probabilityLessThanLimit += thisProbabilityLessThanLimit * weight;
            }
            
            var alaeCdfForClaimsWithoutPay = 1d;
            if (GetAlaeProbabilityLessThanLimitIndicator(alaeForClaimsWithoutPay))
            {
                alaeCdfForClaimsWithoutPay = GetAlaeCdfForClaimsWithoutPay(filteredLimit,limit,alaeForClaimsWithoutPay);
            }

            return alaeCdfForClaimsWithoutPay * parameters.ProbabilityOfNoLoss
                   + (1 - probabilityLessThanLimit) * (1 - parameters.ProbabilityOfNoLoss);
        }
        
        public abstract ICalculator CreateNew();

        protected abstract double GetAlaeLimitedExpectedValueForClaimsWithoutPay(double limit, double forClaimsWithoutPay, double alaeForClaimsWithoutPay);

        protected virtual bool GetAlaeProbabilityLessThanLimitIndicator(double alaeForClaimsWithoutPay)
        {
            return false;
        }

        protected abstract double GetLimitedExpectedValueComponent(double limit, double mean, double alaePercent);
        
        protected double GetLimitedExpectedValue(double limit, double mean)
        {
            return mean*(1 - Math.Exp(-1*limit/mean));
        }

        public static double GetLevIgnoringAlae(double limit, double[] means, double[] probabilities)
        {
            //wanted to use GetLimitedExpectedValue from above, but couldn't access non-static class
            var counter = 0;
            var lev = 0d;

            foreach (var mean in means)
            {
                var levForThisMean = mean * (1 - Math.Exp(-1 * limit / mean));
                var probabilityForThisMean = probabilities[counter++];
                lev += levForThisMean * probabilityForThisMean;
            }

            return lev;
        }

        public virtual double GetEffectiveLimit(double limit, IReinsurancePerspectiveHandler reinsurancePerspective, double alaePercent)
        {
            return reinsurancePerspective.GetEffectiveLimit(limit, PolicyLimit, PolicySir);
        }

        public abstract double GetAlaeCdfForClaimsWithoutPay(double effectiveLimit,double limit,double alaeForClaimsWithoutPay);

        public virtual double GetAppropriateAlaeAmountForFiltering(double alaeAmount)
        {
            return 0;
        }

        public virtual double AdjustMeanWithAlaeWhenNecessary(double mean, double alaePercent)
        {
            return mean;
        }

        public virtual double GetProbabilityLessThanLimitComponent(double limit, double mean, double alaePercent)
        {
            return Math.Exp(-limit/mean);
        }

        protected virtual double FilterAlaeAdjustmentFactorThroughReinsuranceAlaeTreatment(double alaeAdjustmentFactor)
        {
            // for most reinsurance alae treatments, leave alae adjustment factor alone
            return alaeAdjustmentFactor;
        }
    }
}
