using System.Collections.Generic;
using System.Linq;
using MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient;
using PionlearClient.CollectorClientPlus;
using PionlearClient.Extensions;

namespace SubmissionCollector
{
    public static class AllocationExtensions
    {
        public static void Normalize(this IList<Allocation> allocations)
        {
            var sum = allocations.Sum(a => a.Value);
            if (!sum.IsEqual(0) && !sum.IsEqual(1))
            {
                allocations.ForEach(a => a.Value /= sum);
            }
        }

        public static void Normalize(this IList<PolicyDistributionItemPlus> items)
        {
            var sum = items.Sum(a => a.Value);
            if (sum > 0)
            {
                items.ForEach(a => a.Value /= sum);
            }
        }

        public static void Normalize(this IList<HazardDistributionItemPlus> items)
        {
            var sum = items.Sum(a => a.Value);
            if (sum > 0)
            {
                items.ForEach(a => a.Value /= sum);
            }
        }

        public static void Normalize(this IList<ConstructionTypeDistributionItemPlus> items)
        {
            var sum = items.Sum(a => a.Weight);
            if (sum > 0)
            {
                items.ForEach(a => a.Weight /= sum);
            }
        }

        public static void Normalize(this IList<OccupancyTypeDistributionItemPlus> items)
        {
            var sum = items.Sum(a => a.Weight);
            if (sum > 0)
            {
                items.ForEach(a => a.Weight /= sum);
            }
        }

        public static void Normalize(this IList<ProtectionClassDistributionItemPlus> items)
        {
            var sum = items.Sum(a => a.Weight);
            if (sum > 0)
            {
                items.ForEach(a => a.Weight /= sum);
            }
        }

        public static void Normalize(this IList<StateDistributionItemPlus> items)
        {
            var sum = items.Sum(a => a.Value);
            if (sum > 0)
            {
                items.ForEach(a => a.Value /= sum);
            }
        }
    }
}
