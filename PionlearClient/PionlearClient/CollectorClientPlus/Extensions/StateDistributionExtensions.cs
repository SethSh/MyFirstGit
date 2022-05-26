using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus.Extensions
{
    public static class StateDistributionExtensions
    {
        public static bool IsEqualTo(this StateDistributionModel model, StateDistributionModel otherModel)
        {
            //not checking name bc it's system generated and has changed in the past
            if (!model.Items.IsEqualsTo(otherModel.Items)) return false;
            if (!model.SublineIds.IsEqualsTo(otherModel.SublineIds)) return false;

            return true;
        }

        internal static bool IsEqualsTo(this ICollection<StateDistributionItem> items, ICollection<StateDistributionItem> otherItems)
        {
            return items.Count == otherItems.Count &&
                   items.All(state => otherItems.Contains(state, new StateDistributionItemComparer()));
        }

        internal class StateDistributionItemComparer : IEqualityComparer<StateDistributionItem>
        {
            public bool Equals(StateDistributionItem state, StateDistributionItem otherState)
            {
                Debug.Assert(state != null, "state != null");
                Debug.Assert(otherState != null, "otherState != null");
                return state.StateCode == otherState.StateCode && state.Value.IsEpsilonEqual(otherState.Value);
            }

            public int GetHashCode(StateDistributionItem state)
            {
                return $"{state.StateCode}".GetHashCode();
            }
        }
    }
}