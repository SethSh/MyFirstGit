using System.Collections.Generic;
using System.Linq;
using MunichRe.Bex.ApiClient.ClientApi;
using Newtonsoft.Json;
using PionlearClient.Extensions;

namespace PionlearClient.BexReferenceData
{
    public class SublineCodesFromBex : BaseReferenceDataFromBex<LineSublineViewModel>
    {
        public SublineCodesFromBex() : base(BexFileNames.LineSublineCodesFileName)
        {

        }

        protected override string GetJson(IReferenceDataClient referenceData)
        {
            var task = referenceData.LineSublinesAsync();
            task.Wait();
            return JsonConvert.SerializeObject(task.Result, Formatting.Indented);
        }

        public static IList<long> ConvertShortNameWithLobsToCodes(IList<string> names)
        {
            var codes = new List<long>();
            foreach (var name in names)
            {
                var code = ReferenceData.Single(item => $"{item.LineOfBusiness.ShortName.ConnectWithDash(item.SublineShortName)}" == name).SublineId;
                codes.Add(code);
            }
            return codes;
        }

        public static string ConvertCodeToShortNameWithLob(long code)
        {
            var subline = ReferenceData.Single(item => item.SublineId == code);
            return subline.LineOfBusiness.ShortName.ConnectWithDash(subline.SublineShortName);
        }

        public static IList<string> ConvertCodesToShortNameWithLobs(IList<long> codes)
        {
            var names = new List<string>();
            foreach (var code in codes)
            {
                var subline = ReferenceData.Single(item => item.SublineId == code);
                var name = subline.LineOfBusiness.ShortName.ConnectWithDash(subline.SublineShortName);
                names.Add(name);
            }
            return names;
        }

        public static string GetLobName(long sublineId)
        {
            return ReferenceData.Single(item => item.SublineId == sublineId).LineOfBusiness.Name;
        }

        public static IEnumerable<LineSublineViewModel> GetPropertySublines()
        {
            return GetAllSublinesSiblings(BexConstants.CommercialPropertySublineCode);
        }

        public static IEnumerable<LineSublineViewModel> GetWorkersCompensationSublineIds()
        {
            return GetAllSublinesSiblings(BexConstants.WorkersCompensationIndemnitySublineCode);
        }

        public static IEnumerable<LineSublineViewModel> GetAutoPhysicalDamageSublines()
        {
            return GetAllSublinesSiblings(BexConstants.AutoPhysicalDamageCommercialSublineCode);
        }

        private static IEnumerable<LineSublineViewModel> GetAllSublinesSiblings(int sublineCode)
        {
            var subline = ReferenceData.First(item => item.SublineId == sublineCode);
            var lineOfBusinessName = subline.LineOfBusiness.Name;
            return ReferenceData.Where(item => item.LineOfBusiness.Name == lineOfBusinessName);
        }
    }
    
}