using System.Collections.Generic;
using System.Diagnostics;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus
{
    public class ProtectionClassDistributionItemPlus : ProtectionClassAndWeight, IProvidesLocation
    {
        public string Location { get; set; }
    }

    internal class ProtectionClassDistributionItemComparer : IEqualityComparer<ProtectionClassAndWeight>
    {
        public bool Equals(ProtectionClassAndWeight protectionClass,
            ProtectionClassAndWeight otherProtectionClass)
        {
            Debug.Assert(protectionClass != null, "protection class != null");
            Debug.Assert(otherProtectionClass != null, "other protection class != null");
            return protectionClass.ProtectionClassId == otherProtectionClass.ProtectionClassId
                   && protectionClass.Weight.IsEpsilonEqual(otherProtectionClass.Weight);
        }

        public int GetHashCode(ProtectionClassAndWeight protectionClass)
        {
            return $"{protectionClass.ProtectionClassId}".GetHashCode();
        }
    }
}