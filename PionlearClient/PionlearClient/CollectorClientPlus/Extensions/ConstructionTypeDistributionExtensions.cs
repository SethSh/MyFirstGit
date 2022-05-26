using System.Collections.Generic;
using System.Linq;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus.Extensions
{
    public static class ConstructionTypeDistributionExtensions
    {
        public static bool IsEqualTo(this ConstructionTypeDistributionModel model, ConstructionTypeDistributionModel otherModel)
        {
            //not checking name bc it's system generated and has changed in the last release
            if (!model.Items.IsEqualsTo(otherModel.Items)) return false;
            if (!model.SublineIds.IsEqualsTo(otherModel.SublineIds)) return false;

            return true;
        }

        internal static bool IsEqualsTo(this ICollection<ConstructionTypeAndWeight> items,
            ICollection<ConstructionTypeAndWeight> otherItems)
        {
            return items.Count == otherItems.Count &&
                   items.All(ct => otherItems.Contains(ct, new ConstructionTypeDistributionItemComparer()));
        }
    }
}