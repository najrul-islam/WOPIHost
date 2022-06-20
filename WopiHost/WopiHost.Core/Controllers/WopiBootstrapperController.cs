using System;
using System.Linq;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Microsoft.Extensions.Primitives;
using WopiHost.Abstractions;
using WopiHost.Core.Models;
using WopiHost.Core.Results;
using WopiHost.Url;
using WopiHost.Discovery;
using System.Globalization;
using WopiHost.Discovery.Enumerations;
using System.IO;
using WopiHost.FileSystemProvider;
using Microsoft.Extensions.Configuration;
using WopiHost.Utility.Common;
using System.Net.Http;

namespace WopiHost.Core.Controllers
{
    [Route("wopibootstrapper")]
    public class WopiBootstrapperController : WopiControllerBase
    {

        public WopiUrlBuilder _urlGenerator;
        private IConfiguration _configuration { get; }
        private readonly IHttpClientFactory _clientFactory;
        public string WopiClientUrl; // "http://oos-stage.hoxro.com";

        //public string WopiClientUrl => Configuration.GetValue("WopiClientUrl", string.Empty);
        public WopiDiscoverer Discoverer;

        //TODO: remove test culture value and load it from configuration SECTION
        public WopiUrlBuilder UrlGenerator;


        public WopiBootstrapperController(IWopiStorageProvider fileProvider
            , IWopiSecurityHandler securityHandler
            , IOptionsSnapshot<WopiHostOptions> wopiHostOptions
            , IConfiguration configuration
            , IHttpClientFactory clientFactory)
            : base(fileProvider, securityHandler, wopiHostOptions)
        {
            _configuration = configuration;
            _clientFactory = clientFactory;
            WopiClientUrl = _configuration.GetValue("WopiClientUrl", string.Empty);
            Discoverer = new WopiDiscoverer(new HttpDiscoveryFileProvider(WopiClientUrl, _clientFactory));
            UrlGenerator = _urlGenerator ?? (_urlGenerator = new WopiUrlBuilder(Discoverer, new WopiUrlSettings { UI_LLCC = new CultureInfo("en-US") }));
        }

        [HttpPost]
        [Produces("application/json")]
        public IActionResult GetRootContainer()
        {
            var authorizationHeader = HttpContext.Request.Headers["Authorization"];
            var ecosystemOperation = HttpContext.Request.Headers[WopiHeaders.EcosystemOperation];
            string wopiSrc = HttpContext.Request.Headers[WopiHeaders.WopiSrc].FirstOrDefault();

            if (ValidateAuthorizationHeader(authorizationHeader))
            {
                //TODO: supply user
                var user = "Anonymous";
                var userFullName = "Anonymous";

                //TODO: implement bootstrap
                BootstrapRootContainerInfo bootstrapRoot = new BootstrapRootContainerInfo
                {
                    Bootstrap = new BootstrapInfo
                    {
                        EcosystemUrl = GetWopiUrl("ecosystem", accessToken: "TODO"),
                        SignInName = "",
                        UserFriendlyName = "",
                        UserId = ""
                    }
                };
                if (ecosystemOperation == "GET_ROOT_CONTAINER")
                {
                    var resourceId = StorageProvider.RootContainerPointer.Identifier;
                    var token = SecurityHandler.GenerateAccessToken(0, user, userFullName);

                    bootstrapRoot.RootContainerInfo = new RootContainerInfo
                    {
                        ContainerPointer = new ChildContainer
                        {
                            Name = StorageProvider.RootContainerPointer.Name,
                            Url = GetWopiUrl("containers", resourceId, SecurityHandler.WriteToken(token))
                        }
                    };
                }
                else if (ecosystemOperation == "GET_NEW_ACCESS_TOKEN")
                {
                    var token = SecurityHandler.GenerateAccessToken(0, user, userFullName);

                    bootstrapRoot.AccessTokenInfo = new AccessTokenInfo
                    {
                        AccessToken = SecurityHandler.WriteToken(token),
                        AccessTokenExpiry = token.ValidTo.ToUnixTimestamp()
                    };
                }
                else
                {
                    return new NotImplementedResult();
                }
                return new JsonResult(bootstrapRoot);
            }
            else
            {
                //TODO: implement WWW-authentication header https://wopirest.readthedocs.io/en/latest/bootstrapper/Bootstrap.html#www-authenticate-header
                string authorizationUri = "https://contoso.com/api/oauth2/authorize";
                string tokenIssuanceUri = "https://contoso.com/api/oauth2/token";
                string providerId = "tp_contoso";
                string urlSchemes = Uri.EscapeDataString("{\"iOS\" : [\"contoso\",\"contoso - EMM\"], \"Android\" : [\"contoso\",\"contoso - EMM\"], \"UWP\": [\"contoso\",\"contoso - EMM\"]}");
                Response.Headers.Add("WWW-Authenticate", $"Bearer authorization_uri=\"{authorizationUri}\",tokenIssuance_uri=\"{tokenIssuanceUri}\",providerId=\"{providerId}\", UrlSchemes=\"{urlSchemes}\"");
                return new UnauthorizedResult();
            }
        }


        [HttpPost("MakeURL")]
        [Produces("application/json")]
        public IActionResult MakeURL()
        {
            var authorizationHeader = HttpContext.Request.Headers["Authorization"];
            var ecosystemOperation = HttpContext.Request.Headers[WopiHeaders.EcosystemOperation];
            string wopiSrc = HttpContext.Request.Headers[WopiHeaders.WopiSrc].FirstOrDefault();
            var extenssion = HttpContext.Request.Headers[WopiHeaders.Extenssion];
            var action = HttpContext.Request.Headers[WopiHeaders.Action];
            var bucket = HttpContext.Request.Headers[WopiHeaders.SourceBlob];
            var s3Address = HttpContext.Request.Headers[WopiHeaders.SourceBlob];
            var userId = HttpContext.Request.Headers[WopiHeaders.UserName];
            var userFullName = HttpContext.Request.Headers[WopiHeaders.UserFullName];

            WopiActionEnum actionWop = (WopiActionEnum)Enum.Parse(typeof(WopiActionEnum), action);


            if (ValidateAuthorizationHeader(authorizationHeader))
            {
                //TODO: supply user
                var user = userId;//"Anonymous";

                //TODO: implement bootstrap
                BootstrapRootContainerInfo bootstrapRoot = new BootstrapRootContainerInfo
                {
                    Bootstrap = new BootstrapInfo
                    {
                        EcosystemUrl = GetWopiUrl("ecosystem", accessToken: "TODO"),
                        SignInName = "",
                        UserFriendlyName = "",
                        UserId = ""
                    }
                };

                if (ecosystemOperation == "GET_NEW_EDIT_URL")
                {

                    var token = SecurityHandler.GenerateAccessToken(0, user, userFullName);
                    string urlsr = UrlGenerator.GetFileUrlAsync(extenssion, wopiSrc, actionWop).Result;
                    bootstrapRoot.AccessTokenInfo = new AccessTokenInfo
                    {
                        AccessToken = SecurityHandler.WriteToken(token),
                        AccessTokenExpiry = token.ValidTo.ToUnixTimestamp()
                    };
                    var path = Path.Combine(WopiHostOptions.Value.WebRootPath, (WopiHostOptions.Value.WopiRootPath));
                    path = path.Replace("\\.\\", "\\");
                    path = $"{path}\\{Path.GetFileName(s3Address)}";
                    if (!System.IO.File.Exists(path))
                    {
                        //_s3Operator.Download(s3Address, path, bucket);
                    }

                    bootstrapRoot.EditURL = $"{urlsr}&access_token={bootstrapRoot.AccessTokenInfo.AccessToken}&access_token_ttl=0";
                }
                else
                {
                    return new NotImplementedResult();
                }
                return new JsonResult(bootstrapRoot);
            }
            else
            {
                //TODO: implement WWW-authentication header https://wopirest.readthedocs.io/en/latest/bootstrapper/Bootstrap.html#www-authenticate-header
                string authorizationUri = "https://contoso.com/api/oauth2/authorize";
                string tokenIssuanceUri = "https://contoso.com/api/oauth2/token";
                string providerId = "tp_contoso";
                string urlSchemes = Uri.EscapeDataString("{\"iOS\" : [\"contoso\",\"contoso - EMM\"], \"Android\" : [\"contoso\",\"contoso - EMM\"], \"UWP\": [\"contoso\",\"contoso - EMM\"]}");
                Response.Headers.Add("WWW-Authenticate", $"Bearer authorization_uri=\"{authorizationUri}\",tokenIssuance_uri=\"{tokenIssuanceUri}\",providerId=\"{providerId}\", UrlSchemes=\"{urlSchemes}\"");
                return new UnauthorizedResult();
            }
        }
        [HttpPost("UploadToServer")]
        [Produces("application/json")]
        public IActionResult UploadToServer()
        {
            bool output = false;
            var authorizationHeader = HttpContext.Request.Headers["Authorization"];
            var s3Address = HttpContext.Request.Headers[WopiHeaders.SourceBlob];
            if (ValidateAuthorizationHeader(authorizationHeader))
            {
                //TODO: supply user
                var path = Path.Combine(WopiHostOptions.Value.WebRootPath, (WopiHostOptions.Value.WopiRootPath));
                path = path.Replace("\\.\\", "\\");
                path = $"{path}\\{Path.GetFileName(s3Address)}";
                if (System.IO.File.Exists(path))
                {
                    output = true;
                }
                return new JsonResult(output);
            }
            else
            {
                return new UnauthorizedResult();
            }
        }

        private bool ValidateAuthorizationHeader(StringValues authorizationHeader)
        {
            //TODO: implement header validation http://wopi.readthedocs.io/projects/wopirest/en/latest/bootstrapper/GetRootContainer.html#sample-response
            // http://stackoverflow.com/questions/31948426/oauth-bearer-token-authentication-is-not-passing-signature-validation
            return true;
        }
    }
}
