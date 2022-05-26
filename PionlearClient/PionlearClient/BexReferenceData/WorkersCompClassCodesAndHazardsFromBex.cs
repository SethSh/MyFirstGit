using System.Collections.Generic;
using System.Linq;
using MunichRe.Bex.ApiClient.ClientApi;
using Newtonsoft.Json;

namespace PionlearClient.BexReferenceData
{
    public class WorkersCompClassCodesAndHazardsFromBex : BaseReferenceDataFromBex<WorkCompClassCodesAndGroupsModel>
    {
        public WorkersCompClassCodesAndHazardsFromBex() : base(BexFileNames.WorkersCompClassCodesFileName)
        {
            DurationDayCount = 7;
        }

        protected override string GetJson(IReferenceDataClient referenceData)
        {
            var task = referenceData.GetWorkCompClassCodesAndGroupsAsync();
            task.Wait();
            return JsonConvert.SerializeObject(task.Result, Formatting.Indented);
        }

        protected override void DeserializeJson(string json)
        {
            ReferenceData = new List<WorkCompClassCodesAndGroupsModel> {JsonConvert.DeserializeObject<WorkCompClassCodesAndGroupsModel>(json)};
        }

        public static IEnumerable<WorkCompHazardGroupModel> HazardGroups => ReferenceData.First().HazardGroups;

        public static IEnumerable<StateWorkCompClassCodeModel> StateClassCodes => ReferenceData.First().StateClassCodes;

        public static IDictionary<long, IDictionary<long, WorkCompClassCodeModel>> GetClassCodeByStateDictionary(IEnumerable<long> stateIds)
        {
            //keys are state id and class code id
            var classCodeByStateDictionary = new Dictionary<long, IDictionary<long, WorkCompClassCodeModel>>();
            foreach (var stateId in stateIds)
            {
                var classCodes = StateClassCodes.Single(ccs => ccs.State.Id == stateId).ClassCodes;
                var classCodeDictionary = classCodes.ToDictionary(classCode => classCode.Id);
                classCodeByStateDictionary.Add(stateId, classCodeDictionary);
            }
            return classCodeByStateDictionary;
        }

        public static IDictionary<string, IDictionary<long, WorkCompClassCodeModel>> GetClassCodeByStateDictionary(IEnumerable<string> stateAbbreviations)
        {
            //keys are state abbrev and class code
            var classCodeByStateDictionary = new Dictionary<string, IDictionary<long, WorkCompClassCodeModel>>();
            foreach (var stateAbbreviation in stateAbbreviations)
            {
                var classCodes = StateClassCodes.Single(ccs => ccs.State.Abbreviation == stateAbbreviation).ClassCodes;
                var classCodeDictionary = classCodes.ToDictionary(classCode => classCode.StateClassCode);
                classCodeByStateDictionary.Add(stateAbbreviation, classCodeDictionary);
            }
            return classCodeByStateDictionary;
        }
    }

}
