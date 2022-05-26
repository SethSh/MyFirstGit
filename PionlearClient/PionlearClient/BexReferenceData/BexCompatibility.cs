using System;
using System.Net.Http;
using PionlearClient.TokenAuthentication;


namespace PionlearClient.BexReferenceData
{
    public class BexCompatibility
    {
        public static bool IsCompatible;
        public static bool IsConnected;

        public static void GetCompatibility(double version, string secretWord, string uwpfTokenUrl, string bexSubmissionsUrl, string bexBaseUrl)
        {
            try
            {
                var client = BexCollectorClientFactory.CreateBexCollectorClient(secretWord, uwpfTokenUrl, bexSubmissionsUrl, bexBaseUrl);
                var response = client.ReferenceDataClient.WorkbookIsCompatibleAsync((float)version, System.Threading.CancellationToken.None)
                    .GetAwaiter().GetResult();
                
                IsCompatible = response.Flag;
                IsConnected = true;
            }
            catch (HttpRequestException ex)
            {
                IsConnected = false;
                // ReSharper disable once PossibleNullReferenceException
                throw new Exception(ex.InnerException.Message);
            }
        }
    }
}
