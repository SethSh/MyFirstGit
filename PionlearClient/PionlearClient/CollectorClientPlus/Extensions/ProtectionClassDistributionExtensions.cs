using System.Collections.Generic;
using System.Linq;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus.Extensions
{
    public static class ProtectionClassDistributionExtensions
    {
        public static bool IsEqualTo(this ProtectionClassDistributionModel model, ProtectionClassDistributionModel otherModel)
        {
            //not checking name bc it's system generated and has changed in the past
            if (!model.Items.IsEqualsTo(otherModel.Items)) return false;
            if (!model.SublineIds.IsEqualsTo(otherModel.SublineIds)) return false;

            return true;
        }

        internal static bool IsEqualsTo(this ICollection<ProtectionClassAndWeight> items, ICollection<ProtectionClassAndWeight> otherItems)
        {
            return items.Count == otherItems.Count &&
                   items.All(ct => otherItems.Contains(ct, new ProtectionClassDistributionItemComparer()));
        }
    }
}