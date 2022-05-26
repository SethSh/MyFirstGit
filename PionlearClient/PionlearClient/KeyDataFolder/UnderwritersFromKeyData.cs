using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using Newtonsoft.Json;
using PionlearClient.Extensions;

namespace PionlearClient.KeyDataFolder
{
    public static class UnderwritersFromKeyData
    {
        public static List<Underwriter> UnderwriterReferenceData { get; set; }
        public static void GetUnderwriterReferenceData(string appDataFolder, string secretWord, string uwpfTokenUrl, string keyDataBaseUrl)
        {
            if (UnderwriterReferenceData != null) return;
            var filename = Path.Combine(appDataFolder, KeyDataConfiguration.UnderwritersFileName);
            string json;

            try
            {
                if (File.Exists(filename) && (DateTime.Now - File.GetLastWriteTime(filename).Date).TotalDays < 30)
                {
                    json = File.ReadAllText(filename);
                    MapJson(json);
                }
                else
                {
                    var underwriterFinder = new UnderwriterFinder();
                    UnderwriterReferenceData = underwriterFinder.Find(secretWord, uwpfTokenUrl, keyDataBaseUrl).ToList();
                    json = JsonConvert.SerializeObject(UnderwriterReferenceData, Formatting.Indented);
                    json.WriteJsonToFile(appDataFolder, KeyDataConfiguration.UnderwritersFileName);
                }
            }
            catch (WebException)
            {
                json = File.ReadAllText(filename);
                MapJson(json);
            }
        }

        public static string GetName(string code)
        {
            var uw = UnderwriterReferenceData.FirstOrDefault(u => u.Code == code);
            return uw != null ? uw.Name : string.Empty;
        }

        private static void MapJson(string json)
        {
            UnderwriterReferenceData = (List<Underwriter>)JsonConvert.DeserializeObject(json, typeof(List<Underwriter>));
        }
    }

}
