using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WopiHost.Discovery
{
    public class HttpDiscoveryFileProvider : IDiscoveryFileProvider
    {
        private XElement _discoveryXml;
        private readonly string _wopiClientUrl;
        private readonly IHttpClientFactory _clientFactory;

        public HttpDiscoveryFileProvider(string wopiClientUrl
            , IHttpClientFactory clientFactory)
        {
            _wopiClientUrl = wopiClientUrl;
            _clientFactory = clientFactory;
        }

        public async Task<XElement> GetDiscoveryXmlAsync()
        {
            if (_discoveryXml == null)
            {
                try
                {
                    Stream stream;
                    using (HttpClient client = _clientFactory.CreateClient())
                    {
                        stream = await client.GetStreamAsync($"{_wopiClientUrl}/hosting/discovery");
                    }
                    _discoveryXml = await XElement.LoadAsync(stream, LoadOptions.None, CancellationToken.None);
                    stream?.Close();
                }
                catch (HttpRequestException e)
                {
                    throw new DiscoveryException($"There was a problem retrieving the discovery file. Please check availability of the WOPI Client at '{_wopiClientUrl}'.", e);
                }
            }
            return _discoveryXml;
        }
    }
}
