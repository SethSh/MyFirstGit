using System.Collections.Generic;
using System.Diagnostics;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus
{
    public class HazardDistributionItemPlus : HazardDistributionItem, IProvidesLocation
    {
        public string Location { get; set; }
    }

    internal class HazardDistributionItemComparer : IEqualityComparer<HazardDistributionItem>
    {
        public bool Equals(HazardDistributionItem hazard, HazardDistributionItem otherHazard)
        {
            Debug.Assert(hazard != null, "hazard != null");
            Debug.Assert(otherHazard != null, "otherHazard != null");
            return hazard.HazardId == otherHazard.HazardId && hazard.Value.IsEpsilonEqual(otherHazard.Value);
        }

        public int GetHashCode(HazardDistributionItem hazard)
        {
            return $"{hazard.HazardId}".GetHashCode();
        }
    }
}
