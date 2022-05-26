using System.Collections.Generic;
using System.Diagnostics;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus
{
    public class PolicyDistributionItemPlus : PolicyDistributionItem, IProvidesLocation
    {
        public string Location { get; set; }
    }

    internal class PolicyDistributionItemComparer : IEqualityComparer<PolicyDistributionItem>
    {
        public bool Equals(PolicyDistributionItem policy, PolicyDistributionItem otherPolicy)
        {
            Debug.Assert(policy != null, "policy != null");
            Debug.Assert(otherPolicy != null, "otherPolicy != null");
            return policy.Limit.IsEpsilonEqual(otherPolicy.Limit) && policy.Attachment.IsEpsilonEqualIncludingNullAndNaN(otherPolicy.Attachment) && policy.Value.IsEpsilonEqual(otherPolicy.Value);
        }

        public int GetHashCode(PolicyDistributionItem obj)
        {
            return $"L{obj.Limit}A{obj.Attachment}".GetHashCode();
        }
    }

}
