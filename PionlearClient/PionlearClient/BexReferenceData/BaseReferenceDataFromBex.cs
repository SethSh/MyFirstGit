using System;
using System.Collections.Generic;
using System.IO;
using MunichRe.Bex.ApiClient.ClientApi;
using Newtonsoft.Json;
using PionlearClient.Extensions;
using PionlearClient.TokenAuthentication;

namespace PionlearClient.BexReferenceData
{
    public abstract class BaseReferenceDataFromBex<T> 
    {
        private readonly string _fileName;
        public static IEnumerable<T> ReferenceData;
        

        protected BaseReferenceDataFromBex(string fileName)
        {
            _fileName = fileName;
            DurationDayCount = 1;
        }

        protected int DurationDayCount { get; set; }

        public virtual void GetReferenceData(string appDataFolder, string secretWord, string uwpfTokenUrl, string bexSubmissionsUrl, string bexBaseUrl)
        {
            if (ReferenceData != null) return;

            var filename = Path.Combine(appDataFolder, _fileName);
            string json;

            if (File.Exists(filename) && (DateTime.Now - File.GetLastWriteTime(filename).Date).TotalDays < DurationDayCount)
            {
                json = File.ReadAllText(filename);
                DeserializeJson(json);
            }
            else
            {
                
                var collectorClient = BexCollectorClientFactory.CreateBexCollectorClient(secretWord, uwpfTokenUrl, bexSubmissionsUrl, bexBaseUrl);
                json = GetJson(collectorClient.ReferenceDataClient);

                DeserializeJson(json);
                json.WriteJsonToFile(appDataFolder, filename);
            }
        }

        protected abstract string GetJson(IReferenceDataClient referenceData);

        protected virtual void DeserializeJson(string json)
        {
            ReferenceData = (IEnumerable<T>) JsonConvert.DeserializeObject(json, typeof(IEnumerable<T>));
        }
    }
}
