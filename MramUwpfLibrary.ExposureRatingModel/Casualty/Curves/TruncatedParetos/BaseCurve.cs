using System;
using MramUwpfLibrary.Common.Enums;
using MramUwpfLibrary.Common.Extensions;
using MramUwpfLibrary.Common.Layers;
using MramUwpfLibrary.Common.ReinsurancePerspectives;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.TruncatedParetos
{
    public abstract class BaseCurve : ICalculator
    {
        public double GetLayerLimitedExpectedValue(
            double limitTop,
            double limitBottom,
            double policyLimit,
            double policySir,
            double alaeAdjustmentFactor,
            ReinsuranceAlaeTreatmentType reinsuranceAlaeTreatment,
            IReinsurancePerspectiveHandler reinsurancePerspective,
            IParameters curveParameters)
        {
            var limitedExpectedValueTop = GetLev(limitTop, policyLimit, policySir, alaeAdjustmentFactor, reinsurancePerspective, curveParameters);
            var limitedExpectedValueBottom = GetLev(limitBottom, policyLimit, policySir, alaeAdjustmentFactor, reinsurancePerspective, curveParameters);
            return limitedExpectedValueTop - limitedExpectedValueBottom;
        }
        
        public static double GetLayerVarianceIgnoringAlae(ILayer layer, IParameters curveParameters)
        {
            var secMomentTop = GetLimitedSecondMomentIgnoringAlae(layer.Attachment + layer.Limit, curveParameters);
            var fstMomentTop = GetLimitedFirstMoment(layer.Limit + layer.Attachment, curveParameters);
            var cdfTop = GetCdfIgnoringAlae(layer.Limit, curveParameters);

            var secMomentBottom = 0d;
            var fstMomentBottom = 0d;
            var cdfBottom = 0d;
            if (layer.Attachment > 0)
            {
                secMomentBottom = GetLimitedSecondMomentIgnoringAlae(layer.Attachment,
                    curveParameters);
                fstMomentBottom = GetLimitedFirstMoment(layer.Attachment, curveParameters);
                cdfBottom = GetCdfIgnoringAlae(layer.Attachment,
                    curveParameters);
            }

            var layerEx = (fstMomentTop - fstMomentBottom) - layer.Attachment * (cdfTop - cdfBottom) + layer.Limit * (1 - cdfTop);
            return (secMomentTop - secMomentBottom) - 2 * layer.Attachment * (fstMomentTop - fstMomentBottom) +
                   layer.Attachment.Squared() * (cdfTop - cdfBottom) + layer.Limit.Squared() * (1 - cdfTop) -
                   layerEx.Squared();
        }
        public double GetLev(
            double limit, 
            double policyLimit, 
            double policySir, 
            double alaeAdjustmentFactor,
            IReinsurancePerspectiveHandler reinsurancePerspective,
            IParameters curveParameters
            )
        {
            var variableAlaeForSmallLoss = GetAppropriateAlaeAmount(curveParameters.VariableAlaeForSmallLoss)*alaeAdjustmentFactor;
            var variableAlaeForLargeLoss = GetAppropriateAlaeAmount(curveParameters.VariableAlaeForLargeLoss)*alaeAdjustmentFactor;
            var flatAlaeForLargeLoss = GetAppropriateAlaeAmount(curveParameters.FlatAlaeForLargeLoss) *alaeAdjustmentFactor;

            var smallFilteredLimit = GetEffectiveLimit(limit, policyLimit, policySir, reinsurancePerspective, variableAlaeForSmallLoss);
            var largeFilteredLimit = GetEffectiveLimit(limit, policyLimit, policySir, reinsurancePerspective, variableAlaeForLargeLoss);

            if (smallFilteredLimit.IsEpsilonEqualToZero() && largeFilteredLimit.IsEpsilonEqualToZero())
            {
                return 0;
            }

            var levForSmallLoss = GetLevForSmallLoss(smallFilteredLimit, curveParameters.Truncation, curveParameters.AverageSeverityForSmallLoss);
            var levForLargeLoss = GetLevForLargeLoss(largeFilteredLimit, curveParameters.Truncation, curveParameters.Scale, curveParameters.Shape);

            var truncationPlusScale = curveParameters.Truncation + curveParameters.Scale;
            var adjustedLimitPlusScale = Math.Max(largeFilteredLimit, curveParameters.Truncation) + curveParameters.Scale;
            var probabilityGreaterThanTruncation = 1d - curveParameters.ProbabilityLessThanTruncation;
            var onePlusVariableAlaeForLargeLosses = 1d + variableAlaeForLargeLoss;
            var onePlusVariableAlaeForSmallLosses = 1d + variableAlaeForSmallLoss;

            return levForSmallLoss * curveParameters.ProbabilityLessThanTruncation * onePlusVariableAlaeForSmallLosses
                   +
                   probabilityGreaterThanTruncation*levForLargeLoss * onePlusVariableAlaeForLargeLosses
                   + probabilityGreaterThanTruncation*flatAlaeForLargeLoss
                   *(1d - Math.Pow(truncationPlusScale/adjustedLimitPlusScale, curveParameters.Shape));
        }

        public double GetCdfForLimit(
            double limit,
            double policyLimit,
            double policySir,
            double alaeAdjustmentFactor,
            IReinsurancePerspectiveHandler reinsurancePerspective,
            IParameters curveParameters)
        {
            alaeAdjustmentFactor = FilterAlaeAdjustmentFactorThroughReinsuranceAlaeTreatment(alaeAdjustmentFactor);
            
            var variableAlaeForSmallLoss = GetAppropriateAlaeAmount(curveParameters.VariableAlaeForSmallLoss) *alaeAdjustmentFactor;
            var variableAlaeForLargeLoss = GetAppropriateAlaeAmount(curveParameters.VariableAlaeForLargeLoss) *alaeAdjustmentFactor;

            var smallFilteredLimit = GetEffectiveLimit(limit, policyLimit, policySir, reinsurancePerspective, variableAlaeForSmallLoss);
            var probabilityLessThanLimitForSmallLoss = GetCdfForSmallLoss(smallFilteredLimit,
                curveParameters.Truncation, curveParameters.AverageSeverityForSmallLoss);

            var largeFilteredLimit = GetEffectiveLimit(limit, policyLimit, policySir, reinsurancePerspective, variableAlaeForLargeLoss);
            var probabilityLessThanLimitForLargeLoss = GetCdfForLargeLoss(largeFilteredLimit,
                curveParameters.Truncation, curveParameters.Scale, curveParameters.Shape);

            var probabilityGreaterThanTruncation = 1d - curveParameters.ProbabilityLessThanTruncation;

            return probabilityLessThanLimitForSmallLoss* curveParameters.ProbabilityLessThanTruncation
                   + probabilityLessThanLimitForLargeLoss*probabilityGreaterThanTruncation;
        }


        public abstract ICalculator CreateNew();

        protected virtual double FilterAlaeAdjustmentFactorThroughReinsuranceAlaeTreatment(double alaeAdjustmentFactor)
        {
            // for most reinsurance alae treatments, leave alae adjustment factor alone
            return alaeAdjustmentFactor;
        }
        
        public abstract double GetEffectiveLimit(double limit, double  policyLimit, double  policySir, IReinsurancePerspectiveHandler reinsurancePerspective, double variableAlae);

        public virtual double GetAppropriateAlaeAmount(double alaeAmount)
        {
            return alaeAmount;
        }

        public static double GetLimitedSecondMomentIgnoringAlae(double limit, IParameters parameters)
        {
            var small = GetLimitedSecondMomentForSmallLoss(limit, parameters.Truncation, parameters.AverageSeverityForSmallLoss);
            var large = GetLimitedSecondMomentForLargeLoss(limit, parameters.Truncation, parameters.Scale, parameters.Shape);
            return parameters.ProbabilityLessThanTruncation * small + (1 - parameters.ProbabilityLessThanTruncation) * large;
        }
        public static double GetLevIgnoringAlae(double limit, IParameters parameters)
        {
            var small = GetLevForSmallLoss(limit, parameters.Truncation, parameters.AverageSeverityForSmallLoss);
            var large = GetLevForLargeLoss(limit, parameters.Truncation, parameters.Scale, parameters.Shape);
            return parameters.ProbabilityLessThanTruncation *  small + (1 - parameters.ProbabilityLessThanTruncation) * large;
        }

        public static double GetCdfIgnoringAlae(double limit, IParameters parameters)
        {
            var small = GetCdfForSmallLoss(limit, parameters.Truncation, parameters.AverageSeverityForSmallLoss);
            var large = GetCdfForLargeLoss(limit, parameters.Truncation, parameters.Scale, parameters.Shape);
            return parameters.ProbabilityLessThanTruncation * small + (1 - parameters.ProbabilityLessThanTruncation) * large;
        }

        private static double GetLevForSmallLoss(double limit, double truncation, double averageSeverity)
        {
            if (limit < averageSeverity)
            {
                return limit - (truncation - averageSeverity) * limit.Squared() / (truncation * averageSeverity * 2);
            }

            if (limit < truncation)
            {
                var truncationMinusAverageSeverity = truncation - averageSeverity;
                return limit +
                       (averageSeverity / 2 - limit) * truncationMinusAverageSeverity /
                       truncation -
                       averageSeverity * (limit - averageSeverity).Squared() /
                       (2 * truncation * truncationMinusAverageSeverity);
            }

            return averageSeverity;
        }

        private static double GetCdfForSmallLoss(double limit, double truncation, double averageSeverity)
        {
            if (limit < averageSeverity)
            {
                return (truncation - averageSeverity) * limit / (truncation * averageSeverity);
            }

            if (limit < truncation)
            {
                var truncationMinusAverageSeverity = truncation - averageSeverity;
                return truncationMinusAverageSeverity / truncation +
                       averageSeverity * (limit - averageSeverity) /
                       (truncation * truncationMinusAverageSeverity);
            }

            return 1d;
        }

        private static double GetLevForLargeLoss(double limit, double truncation, double scale, double shape)
        {
            if (limit < truncation)
            {
                return limit;
            }

            var truncationPlusScale = truncation + scale;
            // ReSharper disable once CompareOfFloatsByEqualityOperator
            if (shape == 1)
            {
                return truncation + truncationPlusScale * Math.Log10((limit + scale) / truncationPlusScale);
            }

            var shapeMinusOne = shape - 1;
            return truncation +
                   truncationPlusScale / shapeMinusOne *
                   (1 - Math.Pow(truncationPlusScale / (limit + scale), shapeMinusOne));
        }

        private static double GetCdfForLargeLoss(double limit, double truncation, double scale, double shape)
        {
            if (limit < truncation)
            {
                return 0d;
            }

            var truncationPlusScale = truncation + scale;
            return 1d - Math.Pow(truncationPlusScale / (limit + scale), shape);
        }

        private static double GetLimitedSecondMomentForLargeLoss(double limit, double truncation, double scale,
            double shape)
        {
            if (limit < truncation)
            {
                return limit.Squared();
            }

            var part1 = 2 - 3 * shape + Math.Pow(shape, 2);
            var part4T1 = Math.Pow(scale + truncation, shape);
            var part4L1 = Math.Pow(scale + limit, shape);
            var part4T3 = 2 * Math.Pow(scale, 2) + (2 * shape * scale * truncation) - shape * Math.Pow(truncation, 2) +
                             Math.Pow(shape * truncation, 2);
            var part4L3 = 2 * Math.Pow(scale, 2) + (2 * shape * scale * limit) - shape * Math.Pow(limit, 2) +
                             Math.Pow(limit * shape, 2);
            return part4T3 / part1 - part4L3 * part4T1 / (part4L1 * part1);
        }

        private static double GetLimitedSecondMomentForSmallLoss(double limit, double truncation,
            double averageSeverity)
        {

            if (limit < averageSeverity)
            {
                return limit.Squared() -  (truncation - averageSeverity) * Math.Pow(limit, 3) / (3 * averageSeverity * truncation);
            }

            if (limit < truncation)
            {
                return ((Math.Pow(limit, 3) - 2 * truncation * Math.Pow(averageSeverity, 2) +
                         averageSeverity * Math.Pow(truncation, 2)) * averageSeverity) /
                       (2 * truncation * (truncation - averageSeverity));
            }
            return (averageSeverity * (2 * averageSeverity + truncation))/3;
        }

        private static double GetLimitedFirstMoment(double amount, IParameters curveParameters)
        {
            var lev = GetLevIgnoringAlae(amount, curveParameters);
            var backedOutAmount = amount * (1 - curveParameters.ProbabilityLessThanTruncation) *
                                  Math.Pow(
                                      (curveParameters.Scale + curveParameters.Truncation) /
                                      (curveParameters.Scale + amount), curveParameters.Shape);
            return lev - backedOutAmount;
        }
    }
}
