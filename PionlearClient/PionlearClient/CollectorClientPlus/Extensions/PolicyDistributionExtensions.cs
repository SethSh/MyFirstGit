using System.Collections.Generic;
using System.Linq;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus.Extensions
{
    public static class PolicyDistributionExtensions
    {
        public static bool IsEqualTo(this PolicyDistributionModel model, PolicyDistributionModel otherModel)
        {
            //not checking name bc it's system generated and has changed in the past
            if (!model.Items.IsEqualsTo(otherModel.Items)) return false;
            if (!model.SublineIds.IsEqualsTo(otherModel.SublineIds)) return false;
            if (model.UmbrellaTypeId.GetValueOrDefault(0) != otherModel.UmbrellaTypeId.GetValueOrDefault(0)) return false;

            return true;
        }

        internal static bool IsEqualsTo(this ICollection<PolicyDistributionItem> policies, ICollection<PolicyDistributionItem> otherPolicies)
        {
            return policies.Count == otherPolicies.Count &&
                   policies.All(policy => otherPolicies.Contains(policy, new PolicyDistributionItemComparer()));
        }
    }
}