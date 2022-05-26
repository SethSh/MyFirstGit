using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus.Extensions
{
    public static class RateChangeSetModelExtensions
    {
        public static bool IsEqualTo(this RateChangeSetModel model, RateChangeSetModel otherModel)
        {
            //not checking name bc it's system generated and has changed in the past
            if (!model.Sublines.IsEqualsTo(otherModel.Sublines)) return false;
            return true;
        }
    }
}
