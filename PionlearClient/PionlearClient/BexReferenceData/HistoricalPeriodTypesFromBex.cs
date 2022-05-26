using System;
using System.Linq;
using MunichRe.Bex.ApiClient.ClientApi;
using Newtonsoft.Json;

namespace PionlearClient.BexReferenceData
{
    public class HistoricalPeriodTypesFromBex : BaseReferenceDataFromBex<PeriodTypeViewModel>
    {
        public static int DefaultCode => ReferenceData.Single(periodType => periodType.Name == BexConstants.CalendarAccidentYearPeriodTypeName).Id; 
        
        public HistoricalPeriodTypesFromBex() : base(BexFileNames.HistoricalPeriodTypesFileName)
        {

        }

        public static string GetName(int code)
        {
            return ReferenceData.Single(x => x.Id == code).Name;
        }

        public static string GetDateFieldName(int code)
        {
            switch (code)
            {
                case 1: return BexConstants.AccidentDateName;
                case 2: return BexConstants.PolicyDateName;
                case 3: return BexConstants.ReportDateName;
                default: throw new ArgumentOutOfRangeException($"Can't find date type"); 
            }
        }
        
        protected override string GetJson(IReferenceDataClient referenceData)
        {
            var task = referenceData.PeriodTypesAsync();
            task.Wait();
            return JsonConvert.SerializeObject(task.Result, Formatting.Indented);
        }
    }
}