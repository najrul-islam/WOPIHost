using System;
using WopiHost.Url;
using System.Net.Http;
using WopiHost.Discovery;
using WopiHost.Web.Models;
using System.Globalization;
using WopiHost.Abstractions;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using WopiHost.FileSystemProvider;
using Microsoft.Extensions.Options;
using WopiHost.Discovery.Enumerations;
using Microsoft.AspNetCore.Authorization;
using WopiHost.Data.Models;
using System.Text;
using WopiHost.Webv2.Utility;
using Newtonsoft.Json;
using System.Reflection;
using System.Linq;
using System.IO;

namespace WopiHost.Web.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        private WopiUrlBuilder _urlGenerator;

        private IOptionsSnapshot<WopiHostOptions> _optionsSnapshot { get; }
        private IWopiStorageProvider StorageProvider { get; }
        private WopiDiscoverer Discoverer { get; }
        public string WopiClientUrl;
        private readonly IHttpClientFactory _clientFactory;
        private readonly IHttpClientHelperBase _httpClientHelperBase;


        //TODO: remove test culture value and load it from configuration SECTION
        public WopiUrlBuilder UrlGenerator => _urlGenerator ??= new WopiUrlBuilder(
            Discoverer, new WopiUrlSettings
            {
                UI_LLCC = new CultureInfo("en-US"),
                DC_LLCC = new CultureInfo("en-US"),
                DISABLE_CHAT = 1,
                BUSINESS_USER = 1,
                VALIDATOR_TEST_CATEGORY = "All"
            });

        public HomeController(IOptionsSnapshot<WopiHostOptions> wopiOptions,
            IWopiStorageProvider storageProvider,
            IHttpClientFactory clientFactory, IHttpClientHelperBase httpClientHelperBase)
        {
            _optionsSnapshot = wopiOptions;
            StorageProvider = storageProvider;
            _clientFactory = clientFactory;
            WopiClientUrl = _optionsSnapshot.Value.WopiClientUrl;//WopiO365ClientUrl
            Discoverer = new WopiDiscoverer(new HttpDiscoveryFileProvider(WopiClientUrl, _clientFactory));
            _httpClientHelperBase = httpClientHelperBase;
        }

        public async Task<ActionResult> Index()
        {
            try
            {
                string endPoint = $"{_optionsSnapshot.Value.WopiHostUrl}/wopi/Authorization/GetDocumentList";
                var documents = await _httpClientHelperBase.RequestDataAsync<IEnumerable<Document>, IEnumerable<Document>>(endPoint, HttpMethod.Get, null, HttpContext, null);
                //var claims = HttpContext.User.GetCommonClaimFromClaimIdentity();
                /*var files = StorageProvider.GetWopiFiles(StorageProvider.RootContainerPointer.Identifier);
                var fileViewModels = new List<FileViewModel>();
                foreach (var file in files)
                {
                    var iconUri = await Discoverer.GetApplicationFavIconAsync(file.Extension);
                    fileViewModels.Add(new FileViewModel
                    {
                        FileId = file.Identifier,
                        FileName = file.Name,
                        SupportsEdit = await Discoverer.SupportsActionAsync(file.Extension, WopiActionEnum.Edit),
                        SupportsView = await Discoverer.SupportsActionAsync(file.Extension, WopiActionEnum.View),
                        IconUri = !string.IsNullOrEmpty(iconUri) ? new Uri(iconUri, UriKind.Absolute) : new Uri("file.ico", UriKind.Relative)
                    });
                }*/
                if(documents?.Count() > 0)
                {
                    foreach (var document in documents)
                    {
                        document.Extension  = Path.GetExtension(document?.Blob);
                    }
                }
                ViewData["WopiHostUrl"] = _optionsSnapshot.Value.WopiHostUrl;
                return View(documents ?? new List<Document>());
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

        public async Task<ActionResult> Detail(string id, string wopiAction)
        {
            //download
            Document document = new()
            {
                Blob = id,
                DocumentName = id
            };
            id = EncodeIdentifier(id);
            var actionEnum = Enum.Parse<WopiActionEnum>(wopiAction);
            var securityHandler = new WopiSecurityHandler();
            var claims = HttpContext.User.GetCommonClaimFromClaimIdentity();

            var file = StorageProvider.GetWopiFile(id);
            var token = securityHandler.GenerateAccessToken(claims.UserId, claims.Email, claims.UserFullName, WopiDocsTypeEnum.WopiDocs, fileOriginalName: file.Name);

            ViewData["access_token"] = securityHandler.WriteToken(token);
            ViewData["access_token_ttl"] = securityHandler.GenerateAccessToken_TTL();

            var extension = file.Extension.TrimStart('.');
            ViewData["urlsrc"] = await UrlGenerator.GetFileUrlAsync(extension, $"{_optionsSnapshot.Value.WopiHostUrl}/wopi/files/{id}", actionEnum);
            ViewData["favicon"] = await Discoverer.GetApplicationFavIconAsync(extension);

            //download request to wopi
            string endPoint = $"{_optionsSnapshot.Value.WopiHostUrl}/wopi/Authorization/Download";
            _ = await _httpClientHelperBase.RequestDataAsync<Document, Document>(endPoint, HttpMethod.Get, null, HttpContext, document);
            return View();
        }

        #region Private

        private string DecodeIdentifier(string identifier)
        {
            var bytes = Convert.FromBase64String(identifier);
            return Encoding.UTF8.GetString(bytes);
        }

        private string EncodeIdentifier(string path)
        {
            var bytes = Encoding.UTF8.GetBytes(path);
            return Convert.ToBase64String(bytes);
        }
        /*#region Private
        private async Task<T> RequestDataAsync<T, Data>(string endPoint, HttpMethod method, Dictionary<string, string> headers, Data content = null) where T : class where Data : class
        {
            try
            {
                using var response = await _httpClientHelperBase.SendRequest(endPoint, method, headers, content);
                if (response.IsSuccessStatusCode)
                {
                    var reader = await response.Content.ReadAsStringAsync();
                    var res = JsonConvert.DeserializeObject<T>(!reader.Contains("empty") ? reader : "");
                    return res;
                }
                else
                {
                    return await HandleHttpClientErrorResponse<T>(response);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private async Task<T> HandleHttpClientErrorResponse<T>(HttpResponseMessage response) where T : class
        {
            System.Type tType = typeof(T);
            object newTInstance = Activator.CreateInstance(tType);//create an instance of T class
            PropertyInfo propertyTMessage = tType?.GetProperty("Message");
            PropertyInfo propertyTStatusCode = tType?.GetProperty("StatusCode");
            propertyTStatusCode?.SetValue(newTInstance, response.StatusCode);
            propertyTMessage?.SetValue(newTInstance, response?.ReasonPhrase);
            return await System.Threading.Tasks.Task.FromResult(newTInstance as T);
        }
        #endregion*/
        #endregion

    }
}
