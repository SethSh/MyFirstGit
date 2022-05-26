using System;
using KeyData = MunichRe.KeyData.ApiWrapper;

namespace PionlearClient.TokenAuthentication
{
    public class KeyDataApiFactory
    {
        private static KeyData.IApiWrapperClient _keyApiClient;
        private static readonly object Padlock = new object();

        public static KeyData.IApiWrapperClient CreateKeyApi(string secretWord, string uwpfTokenUrl, string keyDataUrl)
        {
            lock (Padlock)
            {
                if (_keyApiClient != null) return _keyApiClient;

                var credentials = new MunichReTokenServiceClientCredentials(new Uri(uwpfTokenUrl), $"\"{secretWord}\"");
                var uri = new Uri(keyDataUrl);
                _keyApiClient = new KeyData.ApiWrapperClient(uri, credentials);
                return _keyApiClient;
            }
        }
    }
}
