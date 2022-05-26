using MunichRe.Bex.ApiClient.ClientApi;
using Newtonsoft.Json;

namespace PionlearClient.BexReferenceData
{
    public class ProtectionClassCodesFromBex : BaseReferenceDataFromBex<ProtectionClassViewModel>
    {
        public ProtectionClassCodesFromBex() : base(BexFileNames.ProtectionClassesFileName)
        {
        
        }

        protected override string GetJson(IReferenceDataClient referenceData)
        {
            var task = referenceData.GetProtectionClassesAsync();
            task.Wait();
            return JsonConvert.SerializeObject(task.Result, Formatting.Indented);
        }
    }
}
