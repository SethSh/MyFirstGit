using System.Collections.Generic;
using System.Linq;
using MunichRe.Bex.ApiClient.ClientApi;
using Newtonsoft.Json;

namespace PionlearClient.BexReferenceData
{
    public class StateCodesFromBex : BaseReferenceDataFromBex<StateViewModel>
    {
        public StateCodesFromBex() : base(BexFileNames.StatesFileName)
        {
                
        }

        protected override string GetJson(IReferenceDataClient referenceData)
        {
            var task = referenceData.StatesWithCountrywideAsync();
            task.Wait();
            return JsonConvert.SerializeObject(task.Result, Formatting.Indented);
        }

        public static IEnumerable<StateViewModel> GetWorkersCompStates()
        {
            return ReferenceData.Where(item => item.Abbreviation != "CW").OrderBy(x => x.DisplayOrder);
        }

        public static IEnumerable<StateViewModel> GetLiabilityStates()
        {
            return ReferenceData.Where(x => x.Abbreviation != "USLH").OrderBy(x => x.DisplayOrder);
        }
    }
}
