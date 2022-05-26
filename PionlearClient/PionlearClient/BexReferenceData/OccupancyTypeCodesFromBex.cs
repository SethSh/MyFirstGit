using MunichRe.Bex.ApiClient.ClientApi;
using Newtonsoft.Json;

namespace PionlearClient.BexReferenceData
{
    public class OccupancyTypeCodesFromBex : BaseReferenceDataFromBex<OccupancyTypeViewModel>
    {
        public OccupancyTypeCodesFromBex() : base(BexFileNames.OccupancyTypesFileName)
        {
        
        }

        protected override string GetJson(IReferenceDataClient referenceData)
        {
            var task = referenceData.GetOccupancyTypesAsync();
            task.Wait();
            return JsonConvert.SerializeObject(task.Result, Formatting.Indented);
        }
    }
}