using System.Collections.Generic;
using MunichRe.Bex.ApiClient.ClientApi;

namespace SubmissionCollector.Models.Comparers
{
    internal class UmbrellaTypeComparer : IEqualityComparer<UmbrellaTypeViewModel>
    {
        public bool Equals(UmbrellaTypeViewModel x, UmbrellaTypeViewModel y)
        {
            return y != null && x != null && x.UmbrellaTypeCode == y.UmbrellaTypeCode;
        }

        public int GetHashCode(UmbrellaTypeViewModel obj)
        {
            return obj.UmbrellaTypeCode.GetHashCode();
        }
    }
}
