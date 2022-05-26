using System;
using MunichRe.Bex.ApiClient;

namespace PionlearClient.TokenAuthentication
{
    public class BexCollectorClientFactory
    {
        private static CollectorClient _collectorClient;
        private static readonly object Padlock = new object();

        public static CollectorClient CreateBexCollectorClient(string secretWord, string uwpfTokenUrl, string bexSubmissionsUrl, string bexUrl)
        {
            lock (Padlock)
            {
                if (_collectorClient != null) return _collectorClient;

                _collectorClient = new CollectorClient(secretWord, new Uri(uwpfTokenUrl), bexSubmissionsUrl, bexUrl);
                return _collectorClient;
            }
        }
    }
}