using MunichRe.Bex.ApiClient.ClientApi;
using Newtonsoft.Json;

namespace PionlearClient.BexReferenceData
{
    public class SubjectPolicyAlaeTreatmentsFromBex : BaseReferenceDataFromBex<SubjectPolicyAlaeTreatmentViewModel>
    {
        public SubjectPolicyAlaeTreatmentsFromBex() : base(BexFileNames.SubjectPolicyAlaeTreatmentsFileName)
        {

        }

        protected override string GetJson(IReferenceDataClient referenceData)
        {
            var task = referenceData.SubjectPolicyAlaeTreatmentsAsync();
            task.Wait();
            return JsonConvert.SerializeObject(task.Result, Formatting.Indented);
        }
    }
}