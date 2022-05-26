using System.Collections.Generic;
using System.Diagnostics;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus
{
    public class ConstructionTypeDistributionItemPlus : ConstructionTypeAndWeight, IProvidesLocation
    {
        public string Location { get; set; }
    }

    internal class ConstructionTypeDistributionItemComparer : IEqualityComparer<ConstructionTypeAndWeight>
    {
        public bool Equals(ConstructionTypeAndWeight constructionType, ConstructionTypeAndWeight otherConstructionType)
        {
            Debug.Assert(constructionType != null, "Construction Type != null");
            Debug.Assert(otherConstructionType != null, "other Construction Type != null");
            return constructionType.ConstructionTypeId == otherConstructionType.ConstructionTypeId
                   && constructionType.Weight.IsEpsilonEqual(otherConstructionType.Weight);
        }

        public int GetHashCode(ConstructionTypeAndWeight constructionType)
        {
            return $"{constructionType.ConstructionTypeId}".GetHashCode();
        }
    }
}