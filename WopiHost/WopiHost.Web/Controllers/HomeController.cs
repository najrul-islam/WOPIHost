using System;
using System.IO;
using System.Text;
using WopiHost.Url;
using System.Net.Http;
using WopiHost.Discovery;
using System.Globalization;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using WopiHost.Discovery.Enumerations;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;

namespace WopiHost.Web.Controllers
{
    public class HomeController : Controller
    {
        private readonly IHttpClientFactory _clientFactory;
        public WopiUrlBuilder _urlGenerator;

        private IConfiguration Configuration { get; }

        public string WopiHostUrl;

        /// <summary>
        /// URL to OWA or OOS
        /// </summary>
        public string WopiClientUrl;

        public WopiDiscoverer Discoverer;

        //TODO: remove test culture value and load it from configuration SECTION
        public WopiUrlBuilder UrlGenerator;

        public HomeController(IConfiguration configuration
            , IHttpClientFactory clientFactory)
        {
            Configuration = configuration;
            _clientFactory = clientFactory;
            WopiHostUrl = Configuration.GetValue("WopiHostUrl", string.Empty);
            WopiClientUrl = Configuration.GetValue("WopiClientUrl", string.Empty);
            Discoverer = new WopiDiscoverer(new HttpDiscoveryFileProvider(WopiClientUrl, _clientFactory));
            UrlGenerator = _urlGenerator ??= new WopiUrlBuilder(Discoverer, new WopiUrlSettings { UI_LLCC = new CultureInfo("en-US") });
        }

        public async Task<ActionResult> Index([FromQuery] string containerUrl)
        {
            try
            {
                if (string.IsNullOrEmpty(containerUrl))
                {
                    //TODO: root folder id http://wopi.readthedocs.io/projects/wopirest/en/latest/ecosystem/GetRootContainer.html?highlight=EnumerateChildren (use ecosystem controller)
                    string containerId = Uri.EscapeDataString(Convert.ToBase64String(Encoding.UTF8.GetBytes(".\\")));
                    string rootContainerUrl = $"{WopiHostUrl}/wopi/containers/{containerId}";
                    containerUrl = rootContainerUrl;
                }

                var token = (await GetAccessToken(containerUrl)).AccessToken;

                dynamic data = await GetDataAsync(containerUrl + $"/children?access_token={token}");
                foreach (var file in data.ChildFiles)
                {
                    string fileUrl = file.Url.ToString();
                    string fileId = fileUrl.Substring($"{WopiHostUrl}/wopi/files/".Length);
                    fileId = fileId.Substring(0, fileId.IndexOf("?", StringComparison.Ordinal));
                    fileId = Uri.UnescapeDataString(fileId);
                    file.Id = fileId;

                    var fileDetails = await GetDataAsync(fileUrl);
                    file.EditUrl = await MakeEditUrl(fileId); //await UrlGenerator.GetFileUrlAsync(fileDetails.FileExtension.ToString().TrimStart('.'), fileUrl, WopiActionEnum.Edit) + "&access_token=xyz";
                    //file.EditUrl = editUrl;
                }
                //http://dotnet-stuff.com/tutorials/aspnet-mvc/how-to-render-different-layout-in-asp-net-mvc
                foreach (var container in data.ChildContainers)
                {
                    //TODO create hierarchy
                }

                return View(data);
            }
            catch (DiscoveryException ex)
            {
                return View("Error", ex);
            }
            catch (HttpRequestException ex)
            {
                return View("Error", ex);
            }
        }

        public async Task<string> MakeEditUrl(string id)
        {
            string output = string.Empty;
            string url = $"{WopiHostUrl}/wopi/files/{id}";
            //var extenssion = "docx";
            //var tokenInfo = await MakeURL(url, extenssion, WopiActionEnum.Edit);

            //ViewData["access_token"] = tokenInfo.AccessToken;
            //TODO: fix
            //ViewData["access_token_ttl"] = tokenInfo.AccessTokenExpiry;

            var bytes = Convert.FromBase64String(id);
            string extension = Path.GetExtension(Encoding.UTF8.GetString(bytes)).Replace(".", "");

            //dynamic fileDetails = await GetDataAsync(url + $"?access_token={tokenInfo.AccessToken}");
            //var extension = fileDetails.FileExtension.ToString().TrimStart('.');
            string urlsr = await UrlGenerator.GetFileUrlAsync(extension, url, WopiActionEnum.Edit);
            ViewData["favicon"] = await Discoverer.GetApplicationFavIconAsync(extension);

            var opt = await MakeURL(url, extension, WopiActionEnum.Edit); //urlsr + "&access_token=" + tokenInfo.AccessToken + "&access_token_ttl=0";
            return opt;
        }

        public async Task<ActionResult> Detail(string id)
        {
            string url = $"{WopiHostUrl}/wopi/files/{id}";
            var tokenInfo = await GetAccessToken(url);

            ViewData["access_token"] = tokenInfo.AccessToken;
            //TODO: fix
            //ViewData["access_token_ttl"] = tokenInfo.AccessTokenExpiry;

            dynamic fileDetails = await GetDataAsync(url + $"?access_token={tokenInfo.AccessToken}");
            var extension = fileDetails.FileExtension.ToString().TrimStart('.');
            ViewData["urlsrc"] = await UrlGenerator.GetFileUrlAsync(extension, url, WopiActionEnum.Edit);
            ViewData["favicon"] = await Discoverer.GetApplicationFavIconAsync(extension);
            return View();
        }
        private async Task<dynamic> GetAccessToken(string resourceUrl)
        {
            var getAccessTokenUrl = $"{WopiHostUrl}/wopibootstrapper";
            dynamic accessTokenData = await RequestDataAsync(getAccessTokenUrl, HttpMethod.Post, new Dictionary<string, string> { { "X-WOPI-EcosystemOperation", "GET_NEW_ACCESS_TOKEN" }, { "X-WOPI-WopiSrc", resourceUrl } });
            return accessTokenData.AccessTokenInfo;
        }
        private async Task<dynamic> MakeURL(string resourceUrl, string extenssion, WopiActionEnum action)
        {
            var getAccessTokenUrl = $"{WopiHostUrl}/wopibootstrapper/MakeURL";
            dynamic accessTokenData = await RequestDataAsync(getAccessTokenUrl, HttpMethod.Post, new Dictionary<string, string> { { "X-WOPI-EcosystemOperation", "GET_NEW_EDIT_URL" }, { "X-WOPI-WopiSrc", resourceUrl }, { "X-WOPI-Extenssion", "docx" }, { "X-WOPI-Action", action.ToString() } });
            return accessTokenData.EditURL;
        }

        private async Task<dynamic> GetDataAsync(string url)
        {
            try
            {
                using (HttpClient client = _clientFactory.CreateClient())
                {
                    using (Stream stream = await client.GetStreamAsync(url))
                    {
                        using (var sr = new StreamReader(stream))
                        {
                            using (var jsonTextReader = new JsonTextReader(sr))
                            {
                                var serializer = new JsonSerializer();
                                return serializer.Deserialize(jsonTextReader);
                            }
                        }
                    }
                }
            }
            catch (HttpRequestException e)
            {
                throw new HttpRequestException($"It was not possible to read data from '{url}'. Please check the availability of the server.", e);
            }
        }
        private async Task<dynamic> RequestDataAsync(string url, HttpMethod method = null, Dictionary<string, string> headers = null)
        {
            try
            {
                method = method ?? HttpMethod.Get;
                using (HttpClient client = _clientFactory.CreateClient())
                {
                    HttpRequestMessage requestMessage = new HttpRequestMessage(method, url);
                    if (headers != null)
                    {
                        foreach (var header in headers)
                        {
                            requestMessage.Headers.Add(header.Key, header.Value);
                        }
                    }

                    using (HttpResponseMessage responseMessage = await client.SendAsync(requestMessage))
                    {
                        string content = await responseMessage.Content.ReadAsStringAsync();
                        return JsonConvert.DeserializeObject(content);
                    }
                }
            }
            catch (HttpRequestException e)
            {
                throw new HttpRequestException($"It was not possible to read data from '{url}'. Please check the availability of the server.", e);
            }
        }
    }
}
