using MunichRe.Bex.ApiClient.CollectorApi;
using PionlearClient.CollectorClientPlus.Extensions;

namespace PionlearClient.CollectorClientPlus
{
    public class ExposureSetModelPlus : ExposureSetModel
    {
        public bool IsEqualTo(ExposureSetModel other)
        {
            //not checking name bc it's system generated and has changed in the past
            if (ExposureBaseId != other.ExposureBaseId) return false;
            if (!Sublines.IsEqualsTo(other.Sublines)) return false;
            return true;
        }
    }
}
