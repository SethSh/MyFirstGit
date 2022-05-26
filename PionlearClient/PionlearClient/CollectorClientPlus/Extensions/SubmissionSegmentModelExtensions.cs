using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus.Extensions
{
    public static class SubmissionSegmentModelExtensions 
    {
        public static bool IsEqualTo(this SubmissionSegmentModel model, SubmissionSegmentModel otherModel)
        {
            if (model.Name != otherModel.Name) return false;
            if (model.HistoricalSubjectBaseUnitOfMeasure != otherModel.HistoricalSubjectBaseUnitOfMeasure) return false;
            if (model.IncludePaidInAggregateLoss != otherModel.IncludePaidInAggregateLoss) return false;
            if (model.IncludePaidInIndividualLoss != otherModel.IncludePaidInIndividualLoss) return false;
            if (model.IsUmbrella != otherModel.IsUmbrella) return false;
            if (model.IsAllocationPercent != otherModel.IsAllocationPercent) return false;
            if (model.PeriodType != otherModel.PeriodType) return false;
            if (model.SubjectBaseUnitOfMeasure != otherModel.SubjectBaseUnitOfMeasure) return false;
            if (model.SubjectPolicyAlaeTreatment != otherModel.SubjectPolicyAlaeTreatment) return false;
            if (!model.SubjectBaseAmount.IsEpsilonEqual(otherModel.SubjectBaseAmount)) return false;
            if (!model.ValidSublines.IsEqualsTo(otherModel.ValidSublines)) return false;
            if (model.IsUmbrella && otherModel.IsUmbrella)
            {
                if (!model.UmbrellaTypeAllocations.IsEqualsTo(otherModel.UmbrellaTypeAllocations)) return false;
            }
            if (!model.ValidPeriods.IsEqualsTo(otherModel.ValidPeriods)) return false;
            if (!model.MinnesotaRetentionId.IsEqualTo(otherModel.MinnesotaRetentionId)) return false;
            
            return true;
        }
    
        internal static bool IsEqualsTo(this ICollection<Allocation> allocations,
            ICollection<Allocation> otherAllocations)
        {
            if (allocations.Count != otherAllocations.Count) return false;
            return allocations.All(allocation => otherAllocations.Contains(allocation, new AllocationItemComparer()));
        }

        internal static bool IsEqualsTo(this ICollection<SubmissionSegmentPeriod> periods,
            ICollection<SubmissionSegmentPeriod> otherPeriods)
        {
            if (periods.Count != otherPeriods.Count) return false;
            return periods.All(period => otherPeriods.Contains(period, new PeriodItemComparer()));
        }

        internal class AllocationItemComparer : IEqualityComparer<Allocation>
        {
            public bool Equals(Allocation x, Allocation y)
            {
                Debug.Assert(x != null, "x != null");
                Debug.Assert(y != null, "y != null");
                return x.Id == y.Id && x.Value.IsEpsilonEqual(y.Value);
            }

            public int GetHashCode(Allocation obj)
            {
                return $"{obj.Id}".GetHashCode();
            }
        }

        internal class PeriodItemComparer : EqualityComparer<SubmissionSegmentPeriod>
        {
            public override bool Equals(SubmissionSegmentPeriod period, SubmissionSegmentPeriod otherPeriod)
            {
                Debug.Assert(period != null, "period != null");
                Debug.Assert(otherPeriod != null, "other period != null");
                return period.StartDate.Date.CompareTo(otherPeriod.StartDate.Date) == 0 &&
                       period.EndDate.Date.CompareTo(otherPeriod.EndDate.Date) == 0 &&
                       period.EvaluationDate.Date.CompareTo(otherPeriod.EvaluationDate.Date) == 0;
            }

            public override int GetHashCode(SubmissionSegmentPeriod period)
            {
                Debug.Assert(period != null, "period != null");
                return (period.StartDate.DateTime.ToShortDateString() +
                        period.EndDate.DateTime.ToShortDateString() +
                        period.EvaluationDate.Date.ToShortDateString()).GetHashCode();
            }
        }
    }

}
