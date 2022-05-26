using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Rest;
using Newtonsoft.Json.Linq;

namespace PionlearClient.TokenAuthentication
{
    public class MunichReTokenServiceClientCredentials : ServiceClientCredentials
    {
        private readonly Uri _uri;
        private readonly string _secret;
        private readonly NetworkCredential _credentials;

        public MunichReTokenServiceClientCredentials(Uri uri, string secret) : this(uri, secret, null)
        {

        }

        public MunichReTokenServiceClientCredentials(Uri uri, string secret,
            NetworkCredential credentials)
        {
            _uri = uri;
            _secret = secret;
            _credentials = credentials;
        }

        private string _token = string.Empty;

        public override void InitializeServiceClient<T>(ServiceClient<T> client)
        {

            GetToken()
                .ContinueWith(task => client.HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", task.Result.Parameter))
                .Wait();
        }

        public override async Task ProcessHttpRequestAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }
            request.Headers.Authorization = await GetToken();
            await base.ProcessHttpRequestAsync(request, cancellationToken);
        }


        private Task<AuthenticationHeaderValue> GetToken()
        {
            var task = new Task<AuthenticationHeaderValue>(() =>
            {
                if (string.IsNullOrEmpty(_token))
                {
                    byte[] data;
                    using (var client = new WebClient())
                    {
                        if (_credentials != null)
                        {
                            client.Credentials = _credentials;
                        }
                        else
                        {
                            client.UseDefaultCredentials = true;
                        }
                        client.Headers.Add(HttpRequestHeader.ContentType, "application/json; charset=utf-8");
                        var uri = _uri.AbsoluteUri;
                        data = client.UploadData(uri, "POST", Encoding.UTF8.GetBytes(_secret));
                    }

                    var json = Encoding.UTF8.GetString(data);
                    _token = JObject.Parse(json)["Token"].ToString();

                }

                System.Diagnostics.Debug.WriteLine("Current token is " + _token);
                return new AuthenticationHeaderValue("Bearer", _token);
            });

            task.Start();
            return task;
        }
    }
}
