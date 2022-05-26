using MunichRe.Bex.ApiClient.ClientApi;
using Newtonsoft.Json;

namespace PionlearClient.BexReferenceData
{
    public class MinnesotaRetentionsFromBex : BaseReferenceDataFromBex<MinnesotaRetentionModel>
    {
        public MinnesotaRetentionsFromBex() : base(BexFileNames.MinnesotaRetentionsFileName)
        {
                
        }

        protected override string GetJson(IReferenceDataClient referenceData)
        {
            var task = referenceData.GetMinnesotaRetentionsAsync();
            task.Wait();
            return JsonConvert.SerializeObject(task.Result, Formatting.Indented);
        }
    }
}
