using System;
using System.Collections.Generic;
using System.Linq;
using MramUwpfLibrary.OnLevel.PolicyFolder;

namespace MramUwpfLibrary.OnLevel.YearTypes
{
    public class CalendarAccidentYearOnLevelCalculator : BaseOnLevelCalculator
    {
        public override DateTime GetFirstPolicyStart(IEnumerable<IPeriod> historicalPeriods)
        {
            const int policyLengthInMonths = 12;
            var firstHistoricalPeriodStart = historicalPeriods.Min(p => p.Start);
            return firstHistoricalPeriodStart.SubtractMonths(policyLengthInMonths);
        }

        public override IEnumerable<VirtualPolicy> FilterOnPoliciesThatImpactTreaty(IEnumerable<VirtualPolicy> policies, IPeriod treatyPeriod)
        {
            return policies.Where(policy => policy.Start <= treatyPeriod.End && policy.End >= treatyPeriod.Start);
        }

        public override double GetTreatyPeriodLevel(IEnumerable<VirtualPolicy> policies, IPeriod treatyPeriod)
        {
            var daysRunningTotal = 0d;
            var factorRunningTotal = 0d;
            foreach (var policy in policies)
            {
                var end = treatyPeriod.End.GetMinimum(policy.End);
                var start = treatyPeriod.Start.GetMaximum(policy.Start);
                var dayCount = start.GetLengthInDays(end, true);

                daysRunningTotal += dayCount;
                factorRunningTotal += dayCount * policy.CumulativeRateFactor;
            }
            return GetLevel(daysRunningTotal, factorRunningTotal);
        }
    }
}