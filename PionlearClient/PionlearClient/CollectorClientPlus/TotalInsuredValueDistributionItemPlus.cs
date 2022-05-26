using System.Collections.Generic;
using System.Diagnostics;
using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus
{
    public class TotalInsuredValueDistributionItemPlus : TotalInsuredValueAndWeight, IProvidesLocation
    {
        public string Location { get; set; }
    }

    internal class TotalInsuredValueDistributionItemComparer : IEqualityComparer<TotalInsuredValueAndWeight>
    {
        public bool Equals(TotalInsuredValueAndWeight tiv, TotalInsuredValueAndWeight otherTiv)
        {
            Debug.Assert(tiv != null, "policy != null");
            Debug.Assert(otherTiv != null, "otherPolicy != null");
            return tiv.TotalInsuredValue.IsEpsilonEqual(otherTiv.TotalInsuredValue)
                   && tiv.Limit.IsEpsilonEqualIncludingNullAndNaN(otherTiv.Limit)
                   && tiv.Attachment.IsEpsilonEqualIncludingNullAndNaN(otherTiv.Attachment)
                   && tiv.Share.IsEpsilonEqualIncludingNullAndNaN(otherTiv.Share)
                   && tiv.Weight.IsEpsilonEqual(otherTiv.Weight);
        }

        public int GetHashCode(TotalInsuredValueAndWeight obj)
        {
            return $"T{obj.TotalInsuredValue}L{obj.Limit}A{obj.Attachment}S{obj.Share}V{obj.Weight}".GetHashCode();
        }
    }

}