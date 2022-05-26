using System.Collections.Generic;
using System.Diagnostics;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus
{
    public class WorkersCompStateClassCodeAndValuePlus : WorkersCompStateClassCodeAndValue, IProvidesLocation
    {
        public string Location { get; set; }
        public bool HasHazard { get; set; }
    }

    internal class WorkersCompStateClassCodeDistributionComparer : IEqualityComparer<WorkersCompStateClassCodeAndValue>
    {
        public bool Equals(WorkersCompStateClassCodeAndValue item, WorkersCompStateClassCodeAndValue otherItem)
        {
            Debug.Assert(item != null, "state attachment != null");
            Debug.Assert(otherItem != null, "other state attachment != null");
            return item.WorkersCompStateId == otherItem.WorkersCompStateId 
                   && item.ClassCodeId == otherItem.ClassCodeId
                   && item.Value.IsEpsilonEqual(otherItem.Value);
        }

        public int GetHashCode(WorkersCompStateClassCodeAndValue obj)
        {
            return $"S{obj.WorkersCompStateId}C{obj.ClassCodeId}V{obj.Value}".GetHashCode();
        }
    }
}