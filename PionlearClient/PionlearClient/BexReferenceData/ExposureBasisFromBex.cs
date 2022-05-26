using System.Linq;
using MunichRe.Bex.ApiClient.ClientApi;
using Newtonsoft.Json;

namespace PionlearClient.BexReferenceData
{
    public class ExposureBasisFromBex : BaseReferenceDataFromBex<ExposureBaseViewModel>
    {
        public ExposureBasisFromBex() : base(BexFileNames.ExposureBasisFileName)
        {

        }

        public static int DefaultCode => ReferenceData.Single(basis => basis.ExposureBaseName == BexConstants.EarnedPremiumExposureBasisName).Id;

        protected override string GetJson(IReferenceDataClient referenceData)
        {
            var task = referenceData.ExposureBasesAsync();
            task.Wait();
            return JsonConvert.SerializeObject(task.Result, Formatting.Indented);
        }

        public static string GetExposureBasisName(int basisId)
        {
            return ReferenceData.Single(x => x.Id == basisId).ExposureBaseName;
        }

        
    }
}
