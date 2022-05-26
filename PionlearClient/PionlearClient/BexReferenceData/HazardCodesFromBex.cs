using MunichRe.Bex.ApiClient.ClientApi;
using Newtonsoft.Json;

namespace PionlearClient.BexReferenceData
{
    public class HazardCodesFromBex : BaseReferenceDataFromBex<HazardViewModel>
    {
        public HazardCodesFromBex() : base(BexFileNames.HazardCodesFileName)
        {

        }

        protected override string GetJson(IReferenceDataClient referenceData)
        {
            var task = referenceData.HazardsAsync();
            task.Wait();
            return JsonConvert.SerializeObject(task.Result, Formatting.Indented);
        }
    }
}
