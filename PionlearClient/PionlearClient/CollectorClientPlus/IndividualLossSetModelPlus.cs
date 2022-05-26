using PionlearClient.CollectorClientPlus.Extensions;
using CollectorApi = MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus
{
    public class IndividualLossSetModelPlus : CollectorApi.IndividualLossSetModel
    {
        public bool IsEqualTo(CollectorApi.IndividualLossSetModel other)
        {
            //don't compare DefaultEvaluationDate
            //not checking name bc it's system generated and has changed in the past
            if (Threshold != other.Threshold) return false;
            if (CombinedLossAndAlae != other.CombinedLossAndAlae) return false;
            if (IncludedPolicyProperties != other.IncludedPolicyProperties) return false;
            if (!Sublines.IsEqualsTo(other.Sublines)) return false;
            
            return true;
        }
    }
}
