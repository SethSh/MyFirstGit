using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using Newtonsoft.Json;
using PionlearClient.Extensions;

namespace PionlearClient.KeyDataFolder
{
    public static class CurrenciesFromKeyData
    {
        public static List<Currency> CurrencyReferenceData { get; set; }

        public static void GetCurrencyReferenceData(string appDataFolder, string secretWord, string uwpfTokenUrl, string keyDataBaseUrl)
        {
            if (CurrencyReferenceData != null) return;
            var filename = Path.Combine(appDataFolder, KeyDataConfiguration.CurrenciesFileName);
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
                    var keyDataApiWrapperClientFacade = new KeyDataApiWrapperClientFacade(secretWord, uwpfTokenUrl, keyDataBaseUrl);
                    CurrencyReferenceData = keyDataApiWrapperClientFacade.GetAllActiveCurrencies();
                    json = JsonConvert.SerializeObject(CurrencyReferenceData, Formatting.Indented);
                    json.WriteJsonToFile(appDataFolder, KeyDataConfiguration.CurrenciesFileName);
                }
            }
            catch (WebException)
            {
                json = File.ReadAllText(filename);
                MapJson(json);
            }
        }

        private static void MapJson(string json)
        {
            CurrencyReferenceData = (List<Currency>)JsonConvert.DeserializeObject(json, typeof(List<Currency>));
        }
    }



    public class Currency
    {
        public string Name { get; set; }
        public string IsoCode { get; set; }
    }
}