using PionlearClient.CollectorClientPlus.Extensions;
using CollectorApi = MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus
{
    public class AggregateLossSetModelPlus : CollectorApi.AggregateLossSetModel
    {
        public bool IsEqualTo(CollectorApi.AggregateLossSetModel other)
        {
            //not checking name bc it's system generated and has changed in the past
            if (CombinedLossAndAlae != other.CombinedLossAndAlae) return false;
            if (!Sublines.IsEqualsTo(other.Sublines)) return false;
            return true;
        }
    }
}
