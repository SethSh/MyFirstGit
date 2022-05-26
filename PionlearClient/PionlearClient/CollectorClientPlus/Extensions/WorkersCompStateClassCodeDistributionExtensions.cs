using System.Collections.Generic;
using System.Linq;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus.Extensions
{
    public static class WorkersCompStateClassCodeDistributionExtensions
    {
        public static bool IsEqualTo(this WorkersCompStateClassCodeDistributionModel model, WorkersCompStateClassCodeDistributionModel otherModel)
        {
            //not checking name bc it's system generated and has changed in the past
            if (!model.Items.IsEqualsTo(otherModel.Items)) return false;
            if (!model.SublineIds.IsEqualsTo(otherModel.SublineIds)) return false;

            return true;
        }
        
        
        internal static bool IsEqualsTo(this ICollection<WorkersCompStateClassCodeAndValue> items,
            ICollection<WorkersCompStateClassCodeAndValue> otherItems)
        {
            return items.Count == otherItems.Count &&
                   items.All(item => otherItems.Contains(item, new WorkersCompStateClassCodeDistributionComparer()));
        }
    }

}