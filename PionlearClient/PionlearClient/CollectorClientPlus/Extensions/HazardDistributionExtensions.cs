using System.Collections.Generic;
using System.Linq;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus.Extensions
{
    public static class HazardDistributionExtensions
    {
        public static bool IsEqualTo(this HazardDistributionModel model, HazardDistributionModel otherModel)
        {
            //not checking name bc it's system generated and has changed in the past
            if (!model.Items.IsEqualsTo(otherModel.Items)) return false;
            if (!model.SublineIds.IsEqualsTo(otherModel.SublineIds)) return false;

            return true;
        }

        public static bool IsEqualsTo(this ICollection<HazardDistributionItem> hazards, ICollection<HazardDistributionItem> otherHazards)
        {
            return hazards.Count == otherHazards.Count &&
                   hazards.All(item => otherHazards.Contains(item, new HazardDistributionItemComparer()));
        }
    }
}