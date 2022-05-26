using System.Collections.Generic;
using System.Linq;
using MunichRe.Bex.ApiClient.ClientApi;
using Newtonsoft.Json;

namespace PionlearClient.BexReferenceData
{
    public class UmbrellaTypesFromBex : BaseReferenceDataFromBex<UmbrellaTypeViewModel>
    {
        public UmbrellaTypesFromBex() : base(BexFileNames.UmbrellaTypesFileName)
        {
        }

        public static long GetPersonalCode()
        {
            return ReferenceData.Single(umb => umb.IsPersonal).UmbrellaTypeCode;
        }

        public static IEnumerable<UmbrellaTypeViewModel> GetCommercialTypes()
        {
            return ReferenceData.Where(umbrellaType => !umbrellaType.IsPersonal);
        }

        protected override string GetJson(IReferenceDataClient referenceData)
        {
            var task = referenceData.UmbrellaTypesAsync();
            task.Wait();
            return JsonConvert.SerializeObject(task.Result, Formatting.Indented);
        }

        public static IList<long> GetCodes(IList<string> names)
        {
            var codes = new List<long>();
            foreach (var name in names)
            {
                var code = ReferenceData.Single(item => item.UmbrellaTypeName == name).UmbrellaTypeCode;
                codes.Add(code);
            }
            return codes;
        }

        public static IList<string> GetNames(IList<int> codes)
        {
            var names = new List<string>();
            foreach (var code in codes)
            {
                var name = ReferenceData.Single(item => item.UmbrellaTypeCode == code).UmbrellaTypeName;
                names.Add(name);
            }
            return names;
        }
        
        public static string GetName(long code)
        {
            return ReferenceData.Single(item => item.UmbrellaTypeCode == code).UmbrellaTypeName;
        }

        public static long GetCode(string name)
        {
            return ReferenceData.Single(item => item.UmbrellaTypeName == name).UmbrellaTypeCode;
        }

        public static string AbbreviateUmbrellaType(string umbrellaTypeName)
        {
            return umbrellaTypeName
                .Replace("Umbrella", "Umbr")
                .Replace("Supported", "Supp")
                .Replace("Unsupported", "Unsupp")
                .Replace("Excess", "Xs");
        }

        public static bool GetIsPersonal(int id)
        {
            return ReferenceData.Single(item => item.UmbrellaTypeCode == id).IsPersonal;
        }
    }
}
