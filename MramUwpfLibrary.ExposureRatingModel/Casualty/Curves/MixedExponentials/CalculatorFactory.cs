using System;
using System.Linq;
using System.Reflection;
using MramUwpfLibrary.Common.Enums;

namespace MramUwpfLibrary.ExposureRatingModel.Casualty.Curves.MixedExponentials
{
    public class CalculatorFactory
    {
        public static ICalculator Create(ReinsuranceAlaeTreatmentType reinsuranceAlaeTreatment, PolicyAlaeTreatmentType policyAlaeTreatment)
        {
            var lookupType = typeof(IMixedExponentialCurveHandler);
            var converters = Assembly.GetExecutingAssembly().GetTypes()
                .Where(t => lookupType.IsAssignableFrom(t) && !t.IsInterface && !t.IsAbstract).ToList();
            var handlers = converters.Select(x => (IMixedExponentialCurveHandler)Activator.CreateInstance(x));

            var handler = handlers.FirstOrDefault(h => h.Handles(reinsuranceAlaeTreatment, policyAlaeTreatment));
            if (handler == null)
                throw new ArgumentOutOfRangeException();

            return handler.CurveCalculator;
        }
    }


    public interface IMixedExponentialCurveHandler
    {
        ICalculator CurveCalculator { get; }
        bool Handles(ReinsuranceAlaeTreatmentType reinsuranceAlaeTreatment, PolicyAlaeTreatmentType policyAlaeTreatment);
    }

    public class MixedExponentialCurveWhenAlaePartOfLossAndInAdditionToLimitHandler : IMixedExponentialCurveHandler
    {
        public ICalculator CurveCalculator => new AlaePartOfLossAndInAdditionToLimit();
        public bool Handles(ReinsuranceAlaeTreatmentType reinsuranceAlaeTreatment, PolicyAlaeTreatmentType policyAlaeTreatment)
        {
            return reinsuranceAlaeTreatment == ReinsuranceAlaeTreatmentType.PartOfLoss
                   && policyAlaeTreatment == PolicyAlaeTreatmentType.InAdditionToLimit;
        }
    }

    public class MixedExponentialCurveWhenAlaeProRataAndInAdditionToLimitHandler : IMixedExponentialCurveHandler
    {
        public ICalculator CurveCalculator => new AlaeProRataAndInAdditionToLimit();
        public bool Handles(ReinsuranceAlaeTreatmentType reinsuranceAlaeTreatment, PolicyAlaeTreatmentType policyAlaeTreatment)
        {
            return reinsuranceAlaeTreatment == ReinsuranceAlaeTreatmentType.ProRata
                   && policyAlaeTreatment == PolicyAlaeTreatmentType.InAdditionToLimit;
        }
    }

    public class MixedExponentialCurveWhenWhenAlaeExcludedAndInAdditionToLimitHandler : IMixedExponentialCurveHandler
    {
        public ICalculator CurveCalculator => new AlaeExcludedAndInAdditionToLimit();
        public bool Handles(ReinsuranceAlaeTreatmentType reinsuranceAlaeTreatment, PolicyAlaeTreatmentType policyAlaeTreatment)
        {
            return reinsuranceAlaeTreatment == ReinsuranceAlaeTreatmentType.Excluded
                   && policyAlaeTreatment == PolicyAlaeTreatmentType.InAdditionToLimit;
        }
    }

    public class MixedExponentialCurveWhenWhenAlaeAlaeAnyAndWithinLimitHandler : IMixedExponentialCurveHandler
    {
        public ICalculator CurveCalculator => new AlaePartOfLossAndWithinLimit();
        public bool Handles(ReinsuranceAlaeTreatmentType reinsuranceAlaeTreatment, PolicyAlaeTreatmentType policyAlaeTreatment)
        {
            return policyAlaeTreatment == PolicyAlaeTreatmentType.WithinLimit;
        }
    }
}
