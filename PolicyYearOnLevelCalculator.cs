using System;
using System.Collections.Generic;
using System.Linq;
using MramUwpfLibrary.OnLevel.PolicyFolder;

namespace MramUwpfLibrary.OnLevel.YearTypes
{
    public class PolicyYearOnLevelCalculator : BaseOnLevelCalculator
    {
        public override DateTime GetFirstPolicyStart(IEnumerable<IPeriod> historicalPeriods)
        {
            return historicalPeriods.Min(p => p.Start);
        }

        public override IEnumerable<VirtualPolicy> FilterOnPoliciesThatImpactTreaty(IEnumerable<VirtualPolicy> policies, IPeriod treatyPeriod)
        {
            return policies.Where(x => x.Start.IsWithin(treatyPeriod.Start, treatyPeriod.End));
        }

        public override double GetTreatyPeriodLevel(IEnumerable<VirtualPolicy> policies, IPeriod treatyPeriod)
        {
            var daysRunningTotal = 0d;
            var factorRunningTotal = 0d;
            foreach (var policy in policies)
            {
                var dayCount = policy.Start.GetLengthInDays(policy.End, true);
                daysRunningTotal += dayCount;
                factorRunningTotal += dayCount * policy.CumulativeRateFactor;
            }
            return GetLevel(daysRunningTotal, factorRunningTotal);
        }
    }
}
