using System.Collections.Generic;
using System.Diagnostics;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus
{
    public class WorkersCompStateHazardGroupAndWeightPlus : WorkersCompStateHazardGroupAndWeight, IProvidesLocation
    {
        public string Location { get; set; }
    }

    internal class WorkersCompStateHazardDistributionComparer : IEqualityComparer<WorkersCompStateHazardGroupAndWeight>
    {
        public bool Equals(WorkersCompStateHazardGroupAndWeight item, WorkersCompStateHazardGroupAndWeight otherItem)
        {
            Debug.Assert(item != null, "State Hazard != null");
            Debug.Assert(otherItem != null, "Other State Hazard != null");
            return item.WorkersCompStateId == otherItem.WorkersCompStateId
                   && item.WorkersCompHazardGroupId == otherItem.WorkersCompHazardGroupId
                   && item.Value.IsEpsilonEqual(otherItem.Value);
        }

        public int GetHashCode(WorkersCompStateHazardGroupAndWeight obj)
        {
            return $"S{obj.WorkersCompStateId}H{obj.WorkersCompHazardGroupId}V{obj.Value}".GetHashCode();
        }
    }
}