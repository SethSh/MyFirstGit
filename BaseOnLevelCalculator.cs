using System;
using System.Collections.Generic;
using MramUwpfLibrary.Common.Extensions;
using MramUwpfLibrary.OnLevel.PolicyFolder;

namespace MramUwpfLibrary.OnLevel.YearTypes
{
    public abstract class BaseOnLevelCalculator : IOnLevelCalculator
    {
        public abstract DateTime GetFirstPolicyStart(IEnumerable<IPeriod> historicalPeriods);
        public abstract IEnumerable<VirtualPolicy> FilterOnPoliciesThatImpactTreaty(IEnumerable<VirtualPolicy> policies, IPeriod treatyPeriod);
        public abstract double GetTreatyPeriodLevel(IEnumerable<VirtualPolicy> policies, IPeriod treatyPeriod);

        protected double GetLevel(double daysRunningTotal, double factorRunningTotal)
        {
            return factorRunningTotal.DivideByWithTrap(daysRunningTotal, 1);
        }
    }
}
