using System.Collections.Generic;
using System.Diagnostics;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus
{
    public class OccupancyTypeDistributionItemPlus : OccupancyTypeAndWeight, IProvidesLocation
    {
        public string Location { get; set; }
    }


    internal class OccupancyTypeDistributionItemComparer : IEqualityComparer<OccupancyTypeAndWeight>
    {
        public bool Equals(OccupancyTypeAndWeight occupancyType,
            OccupancyTypeAndWeight otherOccupancyType)
        {
            Debug.Assert(occupancyType != null, "Occupancy Type != null");
            Debug.Assert(otherOccupancyType != null, "other Occupancy Type != null");
            return occupancyType.OccupancyTypeId == otherOccupancyType.OccupancyTypeId
                   && occupancyType.Weight.IsEpsilonEqual(otherOccupancyType.Weight);
        }

        public int GetHashCode(OccupancyTypeAndWeight occupancyType)
        {
            return $"{occupancyType.OccupancyTypeId}".GetHashCode();
        }
    }
}