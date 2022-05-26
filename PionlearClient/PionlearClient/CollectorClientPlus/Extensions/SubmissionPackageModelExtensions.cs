using MunichRe.Bex.ApiClient.CollectorApi;

namespace PionlearClient.CollectorClientPlus.Extensions
{
    public static class SubmissionPackageModelExtensions 
    {
        public static bool IsEqualTo(this SubmissionPackageModel model, SubmissionPackageModel otherModel)
        {
            //ignore AsOfDate 

            if (model.Name != otherModel.Name) return false;
            if (model.CedentId != otherModel.CedentId) return false;
            if (model.Currency != otherModel.Currency) return false;
            if (model.AnalystId != otherModel.AnalystId) return false;
            if (model.UnderwritingYear != otherModel.UnderwritingYear) return false;
            return true;
        }
    }
}
