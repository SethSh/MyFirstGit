using MunichRe.Bex.ApiClient.ClientApi;
using Newtonsoft.Json;

namespace PionlearClient.BexReferenceData
{
    public class ConstructionTypeCodesFromBex : BaseReferenceDataFromBex<ConstructionTypeViewModel>
    {
        public ConstructionTypeCodesFromBex() : base(BexFileNames.ConstructionTypesFileName)
        {
    
        }

        protected override string GetJson(IReferenceDataClient referenceData)
        {
            var task = referenceData.GetConstructionTypesAsync();
            task.Wait();
            return JsonConvert.SerializeObject(task.Result, Formatting.Indented);
        }
    }
}
