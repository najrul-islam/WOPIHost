using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;
using Microsoft.Extensions.Primitives;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WopiHost.Abstractions;
using WopiHost.Core.Models;
using WopiHost.Discovery;
using WopiHost.Discovery.Enumerations;
using WopiHost.FileSystemProvider;
using WopiHost.Service.Document;
using WopiHost.Url;
using WopiHost.Utility.Common;
using WopiHost.Utility.ExtensionMetod;
using WopiHost.Utility.ViewModel;
using WopiHost.Utility.XMLProcess;

namespace WopiHost.Core.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class DocumentController : ControllerBase
    {
        #region Controller Properties
        //readonly string ROOT_PATH = @".\";
        readonly string ROOT_PATH = "";
        //public WopiUrlBuilder _urlGenerator;
        private IConfiguration Configuration { get; }

        readonly IAzureBlobStorage _azureBlobStorage;
        //private readonly ILiteDbFileStorageInfoManager _liteDbManager;
        private readonly IDocumentService _documentService;
        private readonly IHttpClientFactory _clientFactory;

        //public string WopiClientUrl => _configuration.GetValue(WopiHostOptions.Value.WopiClientUrl, string.Empty); 

        public string WopiClientUrl;
        public WopiDiscoverer Discoverer;

        //TODO: remove test culture value and load it from configuration SECTION
        //public WopiUrlBuilder UrlGenerator => _urlGenerator ?? (_urlGenerator = new WopiUrlBuilder(Discoverer, new WopiUrlSettings { UI_LLCC = new CultureInfo("en-US") }));
        //private string UI_LLCC => !string.IsNullOrEmpty(HttpContext.Request.Headers[WopiHeaders.UI_LLCC]) ? (string)HttpContext.Request.Headers[WopiHeaders.UI_LLCC] : "en-US";

        //get from api
        private string UI_LLCC => !string.IsNullOrEmpty(HttpContext?.Request.Headers[WopiHeaders.UI_LLCC])
            ? HttpContext?.Request.Headers[WopiHeaders.UI_LLCC].ToString() : "en-US";
        //For set OOS UI and Menu
        //private string UI_LLCC_UI => UI_LLCC == "en-GB" ? "en-US" : UI_LLCC;
        public WopiUrlBuilder UrlGenerator => new(Discoverer, new WopiUrlSettings
        {
            UI_LLCC = new CultureInfo(UI_LLCC == "en-GB" ? "en-US" : UI_LLCC),
            DC_LLCC = new CultureInfo(UI_LLCC == "en-GB" ? "en-US" : UI_LLCC),
            DISABLE_CHAT = 1,
            BUSINESS_USER = 1,
            VALIDATOR_TEST_CATEGORY = "All"
        });

        public IWopiStorageProvider FileProvider { get; }
        public IWopiSecurityHandler SecurityHandler { get; }
        public IOptionsSnapshot<WopiHostOptions> WopiHostOptions { get; }
        #endregion

        public DocumentController(IWopiStorageProvider fileProvider,
            IWopiSecurityHandler securityHandler,
            IOptionsSnapshot<WopiHostOptions> wopiHostOptions,
            IConfiguration configuration,
            IAzureBlobStorage azureBlobStorage,
            //ILiteDbFileStorageInfoManager liteDbManager,
            IDocumentService documentService,
            IHttpClientFactory clientFactory)
        {
            FileProvider = fileProvider;
            SecurityHandler = securityHandler;
            WopiHostOptions = wopiHostOptions;
            _azureBlobStorage = azureBlobStorage;
            //_liteDbManager = liteDbManager;
            _documentService = documentService;
            Configuration = configuration;
            _clientFactory = clientFactory;
            WopiClientUrl = WopiHostOptions.Value.WopiO365ClientUrl;//WopiO365ClientUrl
            Discoverer = new WopiDiscoverer(new HttpDiscoveryFileProvider(WopiClientUrl, clientFactory));
        }

        /*[HttpGet("GetLibraryDocumentCreateUrl/{type}/{fontName}")]
        public async Task<IActionResult> GetLibraryDocumentCreateUrl(string type, string fontName)
        {
            var UrlGeneratorNew = await NewOOSWopiUrlBuilder();
            NewContentTemplateResponse urlResponse = new NewContentTemplateResponse();
            try
            {
                var headers = new RequestProcessHeader(HttpContext.Request);
                string newFileName = $"{Guid.NewGuid()}.docx";
                var path = Path.Combine(WopiHostOptions.Value.WebRootPath, (WopiHostOptions.Value.WopiRootPath));
                path = path.AddSourceAddress(newFileName);
                string contentOrTemplate = headers?.OperationType == WopiOperationTypeEnum.BlankContent.ToString() ? "content" : headers?.OperationType == WopiOperationTypeEnum.BlankInvoice.ToString() ? "invoice" : "template";
                var blankFileAddress = $"{WopiHostOptions.Value.WebRootPath}\\BlankTemplate\\{contentOrTemplate}.docx";
                string newFileLocation = $"{WopiHostOptions.Value.WebRootPath}\\{WopiHostOptions.Value.WopiRootPath}\\{newFileName}";
                System.IO.File.Copy(blankFileAddress, newFileLocation, true);
                ProcessXmlDocument.AddFontAndLanguage(newFileLocation, UI_LLCC, fontName);
                //ProcessXmlDocument.CreateNewDocx(newFileLocation, headers?.OperationType, UI_LLCC, fontName);
                SetFileAttribute(newFileLocation, FileAttributes.Normal);
                byte[] bytes = System.Text.Encoding.UTF8.GetBytes($"{ROOT_PATH}{Path.GetFileName(newFileName)}");
                var fileIdentifier = Convert.ToBase64String(bytes);
                headers.WopiSource += fileIdentifier;
                var fileInfo = new FileStorageInfo(headers, fileIdentifier, newFileLocation);
                await _liteDbManager.InsertRequest(fileInfo);
                string urlsr = await UrlGeneratorNew.GetFileUrlAsync(headers.Extension, headers.WopiSource, WopiActionEnum.Edit);
                //return ReturnAction(urlsr, headers.UserName, headers.UserFullName, headers.WopiSource, WopiActionEnum.Edit);
                urlResponse.Url = await ReturnOOSUrlStringAsync(urlsr, headers.UserName, headers.UserFullName, headers.WopiSource, WopiActionEnum.Edit);
                urlResponse.NewFileName = Path.GetFileName(newFileName);
            }
            catch (Exception ex)
            {
                await ReturnExceptionAsync(ex, urlResponse);
            }
            return Ok(urlResponse);
        }*/
        [HttpGet("GetLibraryContentDocumentCreateUrl/{type}/{fontName}")]
        public async Task<IActionResult> GetLibraryContentDocumentCreateUrl(string type, string fontName)
        {
            //var UrlGeneratorNew = await NewOOSWopiUrlBuilder();
            NewContentTemplateResponse urlResponse = new();
            try
            {
                var headers = new RequestProcessHeader(HttpContext.Request);
                if (WopiHostOptions.Value.WopiClientUrl == "https://oos2019.hoxro.com")
                {
                    headers.WopiSrc = headers.WopiSrc.Replace("https://", "http://");
                }
                string newFileName = $"{Guid.NewGuid()}.docx";
                string contentOrTemplate = type;
                var blankFileAddress = $"{WopiHostOptions.Value.WebRootPath}\\BlankTemplate\\{contentOrTemplate}.docx";
                string newFileLocation = $"{WopiHostOptions.Value.WebRootPath}\\{WopiHostOptions.Value.WopiRootPath}\\{newFileName}";
                System.IO.File.Copy(blankFileAddress, newFileLocation, true);
                ProcessXmlDocument.AddFontAndLanguage(newFileLocation, UI_LLCC, fontName);
                //SetFileAttribute(newFileLocation, FileAttributes.Normal);
                headers.FileName = Path.GetFileName(newFileName);
                /*byte[] bytes = System.Text.Encoding.UTF8.GetBytes($"{ROOT_PATH}{Path.GetFileName(newFileName)}");
                var fileIdentifier = Convert.ToBase64String(bytes);
                headers.WopiSrc += fileIdentifier;
                //var fileInfo = new FileStorageInfo(headers, fileIdentifier, newFileLocation);
                //await _liteDbManager.InsertRequest(fileInfo);
                string urlsr = await UrlGeneratorNew.GetFileUrlAsync(headers.Extenssion, headers.WopiSrc, WopiActionEnum.Edit);
                //return ReturnAction(urlsr, headers.UserName, headers.UserFullName, headers.WopiSource, WopiActionEnum.Edit);
                urlResponse.Url = await ReturnOOSUrlStringAsync(urlsr, headers.UserName, headers.UserFullName, headers.WopiSrc, WopiActionEnum.Edit);*/
                urlResponse.Url = await GetFileUrlAsync(headers);
                urlResponse.NewFileName = headers.FileName;
            }
            catch (Exception ex)
            {
                await ReturnExceptionAsync(ex, urlResponse);
            }
            return Ok(urlResponse);
        }
        [HttpGet("GetLibraryTemplateDocumentCreateUrl/{fontName}")]
        public async Task<IActionResult> GetLibraryTemplateDocumentCreateUrl(string fontName)
        {
            NewContentTemplateResponse urlResponse = new();
            try
            {
                var headers = new RequestProcessHeader(HttpContext.Request);
                string newFileName = $"{Guid.NewGuid()}.docx";
                string pathTemplate = headers?.SourceBlob;
                string newFileLocation = $"{WopiHostOptions.Value.WebRootPath}\\{WopiHostOptions.Value.WopiRootPath}\\{newFileName}";
                var status = await DownloadFile(newFileLocation, headers.StorageName, headers.SourceBlob, headers.SourceBlob);
                ProcessXmlDocument.AddFontAndLanguage(newFileLocation, UI_LLCC, fontName);
                //SetFileAttribute(newFileLocation, FileAttributes.Normal);
                /*byte[] bytes = System.Text.Encoding.UTF8.GetBytes($"{ROOT_PATH}{Path.GetFileName(newFileName)}");
                var fileIdentifier = Convert.ToBase64String(bytes);
                headers.WopiSrc += fileIdentifier;
                var fileInfo = new FileStorageInfo(headers, fileIdentifier, newFileLocation);
                string urlsr = await UrlGenerator.GetFileUrlAsync(headers.Extenssion, headers.WopiSrc, WopiActionEnum.Edit);
                urlResponse.Url = await ReturnOOSUrlStringAsync(urlsr, headers.UserName, headers.UserFullName, headers.WopiSrc, WopiActionEnum.Edit);
                urlResponse.NewFileName = Path.GetFileName(newFileName);*/
                headers.FileName = Path.GetFileName(newFileName);
                headers.Action = nameof(WopiActionEnum.Edit);
                urlResponse.Url = await GetFileUrlAsync(headers);
                urlResponse.NewFileName = headers.FileName;
            }
            catch (Exception ex)
            {
                await ReturnExceptionAsync(ex, urlResponse);
            }
            return Ok(urlResponse);
        }

        [HttpGet("view")]
        public async Task<IActionResult> DocumentViewUrl()
        {
            WopiUrlResponse urlResponse = new();
            try
            {
                var headers = new RequestProcessHeader(HttpContext.Request);
                //var wopiAction = HttpContext.Request.Headers[WopiHeaders.Action];
                string path = $"{WopiHostOptions.Value.WebRootPath}\\{WopiHostOptions.Value.WopiRootPath}\\{headers.FileName}";
                await ViewDownloadFile(path, headers.StorageName, headers.SourceBlob, false);
                urlResponse.Url = await GetFileUrlAsync(headers);
            }
            catch (Exception ex)
            {
                if (System.IO.File.Exists($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt"))
                {
                    System.IO.File.AppendAllText($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt", $"GetViewUrl: {DateTime.UtcNow}->{ex.Message} {Environment.NewLine}");
                }
                urlResponse.Message = ex.Message;
                urlResponse.StatusCode = 404;
                await ReturnExceptionAsync(ex, urlResponse);
            }
            return Ok(urlResponse);
        }

        [HttpGet("edit")]
        public async Task<IActionResult> GetLibraryDocumentEditUrl()
        {
            //for new oos
            //var UrlGeneratorNew = await NewOOSWopiUrlBuilder();
            WopiUrlResponse urlResponse = new();
            try
            {
                var headers = new RequestProcessHeader(HttpContext.Request);
                if (WopiHostOptions.Value.WopiClientUrl == "https://oos2019.hoxro.com")
                {
                    headers.WopiSrc = headers.WopiSrc.Replace("https://", "http://");
                }
                //headers.FileName = headers.FileName;
                var path = Path.Combine(WopiHostOptions.Value.WebRootPath, (WopiHostOptions.Value.WopiRootPath));
                path = path.AddSourceAddress(headers.FileName);
                await EditDownloadFile(path, headers.StorageName, headers.SourceBlob, false);
                //SetFileAttribute(newFileLocation, FileAttributes.Normal);
                /*string newFileLocation = $"{WopiHostOptions.Value.WebRootPath}\\{WopiHostOptions.Value.WopiRootPath}\\{headers.FileName}";
                byte[] bytes = System.Text.Encoding.UTF8.GetBytes($"{ROOT_PATH}{Path.GetFileName(headers.FileName)}");
                var fileIdentifier = Convert.ToBase64String(bytes);
                headers.WopiSrc += fileIdentifier;
                /*var fileInfo = new FileStorageInfo(headers, fileIdentifier, newFileLocation)
                {
                    IsEditing = true
                };
                //await _liteDbManager.InsertRequest(fileInfo);
                string urlsr = await UrlGeneratorNew.GetFileUrlAsync(headers.Extenssion, headers.WopiSrc, WopiActionEnum.Edit);
                urlResponse.Url = await ReturnOOSUrlStringAsync(urlsr, headers.UserName, headers.UserFullName, headers.WopiSrc, WopiActionEnum.Edit, fileOriginalName: headers?.FileOriginalName ?? string.Empty);*/

                headers.Action = nameof(WopiActionEnum.Edit);
                urlResponse.Url = await GetFileUrlAsync(headers);
            }
            catch (Exception ex)
            {
                await ReturnExceptionAsync(ex, urlResponse);
            }
            return Ok(urlResponse);
        }

        [HttpGet("OOSViewEmailAndOthers")]
        public async Task<IActionResult> GetOOSViewUrlEmailAndOthers()
        {
            WopiUrlResponse urlResponse = new();
            try
            {
                var headers = new RequestProcessHeader(HttpContext.Request);
                var wopiAction = HttpContext.Request.Headers[WopiHeaders.Action];
                WopiActionEnum wopiActionEnum = (WopiActionEnum)Enum.Parse(typeof(WopiActionEnum), !string.IsNullOrEmpty(wopiAction) ? wopiAction : (StringValues)"EmbedView");

                string path = $"{WopiHostOptions.Value.WebRootPath}\\other-docs\\{headers.FileName}";
                await DownloadFileEmailAndOthers(path, headers.StorageName, headers.SourceBlob, false);

                /*var bytes = System.Text.Encoding.UTF8.GetBytes($"{ROOT_PATH}{Path.GetFileName(headers.FileName)}");
                var fileIdentifier = Convert.ToBase64String(bytes);
                SetFileAttribute($"{WopiHostOptions.Value.WebRootPath}\\other-docs\\{headers.FileName}", FileAttributes.Normal);
                headers.WopiSrc += fileIdentifier;
                string urlsr = await UrlGenerator.GetFileUrlAsync(headers.Extenssion, headers.WopiSrc, wopiActionEnum);
                urlResponse.Url = await ReturnOOSUrlStringAsync(urlsr, headers.UserName, headers.UserFullName, headers.WopiSrc, wopiActionEnum, WopiDocsTypeEnum.OtherDocs, fileOriginalName: headers?.FileOriginalName ?? string.Empty);
                //urlResponse.Url = urlResponse.Url.Replace("PdfMode=1&", "");*/

                headers.Action = nameof(WopiActionEnum.EmbedView);
                urlResponse.Url = await GetFileUrlAsync(headers, WopiDocsTypeEnum.OtherDocs);
            }
            catch (Exception ex)
            {
                if (System.IO.File.Exists($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt"))
                {
                    System.IO.File.AppendAllText($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt", $"OOSViewEmailAndOthers: {ex.Message.ToString()} {Environment.NewLine}");
                }
                return ReturnException(ex);
            }
            return Ok(urlResponse.Url);
        }

        [HttpPost("Cancel")]
        public async Task<IActionResult> Cancel()
        {
            var authorizationHeader = HttpContext.Request.Headers["Authorization"];
            //var s3Address = HttpContext.Request.Headers[WopiHeaders.SourceBlob];
            //var searchFileName = HttpContext.Request.Headers[WopiHeaders.FileName];

            //var newFilename = string.Empty;
            if (ValidateAuthorizationHeader(authorizationHeader))
            {
                //var extension = Path.GetExtension(s3Address);
                //var searchFileNamWithExtenssion = $"{Path.GetFileNameWithoutExtension(searchFileName)}{extension}";
                //var path = Path.Combine(WopiHostOptions.Value.WebRootPath, (WopiHostOptions.Value.WopiRootPath));
                //path = path.AddSourceAddress(searchFileNamWithExtenssion);
                //bool output = true;
                /*if (System.IO.File.Exists(path))
                {
                    //var bytes = System.Text.Encoding.UTF8.GetBytes($"{searchFileNamWithExtenssion}");
                    //var fileIdentifier = Convert.ToBase64String(bytes);
                    /*var item = await _liteDbManager.Get(fileIdentifier);
                    if (item is not null)
                    {
                        item.IsCancel = true;
                        await _liteDbManager.Update(item);
                        output = true;
                    }
                    else
                    {
                        output = false;
                    }
                }*/
                return await Task.FromResult(new JsonResult(true));
            }
            else
            {
                return new UnauthorizedResult();
            }
        }
        [HttpPost("UploadToServer")]
        public async Task<IActionResult> UploadToServer()
        {
            try
            {
                var authorizationHeader = HttpContext.Request.Headers["Authorization"];
                var blob = HttpContext.Request.Headers[WopiHeaders.SourceBlob];
                var userId = HttpContext.Request.Headers[WopiHeaders.UserName];
                var fileName = HttpContext.Request.Headers[WopiHeaders.FileName];
                var storageName = HttpContext.Request.Headers[WopiHeaders.StorageName];
                if (ValidateAuthorizationHeader(authorizationHeader))
                {
                    var extension = Path.GetExtension(blob);
                    if (extension == Utility.StaticData.Extensions.Docx)
                    {
                        var searchFileNamWithExtenssion = $"{Path.GetFileNameWithoutExtension(fileName)}{extension}";
                        var path = Path.Combine(WopiHostOptions.Value.WebRootPath, WopiHostOptions.Value.WopiRootPath);
                        path = path.AddSourceAddress(searchFileNamWithExtenssion);

                        var bytes = System.Text.Encoding.UTF8.GetBytes($"{searchFileNamWithExtenssion}");
                        var fileIdentifier = Convert.ToBase64String(bytes);
                        if (System.IO.File.Exists(path))
                        {
                            //upload to azure
                            await _azureBlobStorage.UploadAsync(storageName, blob, path);
                            /*var item = await _liteDbManager.Get(fileIdentifier);
                            if (item is not null)
                            {
                                item.IsUpDateDone = true;
                                item.Blob = blob;
                                item.UpSearchFileName = fileName;
                                item.UpUserId = userId;
                                await _liteDbManager.Update(item);
                            }*/
                            return await Task.FromResult(new JsonResult(true));
                        }
                    }
                }
                else
                {
                    return new UnauthorizedResult();
                }
                return await Task.FromResult(new JsonResult(false));
            }
            catch (Exception)
            {
                return await Task.FromResult(new JsonResult(false));
            }
        }
        /*[HttpPost("EditMatterDocument")]
        public async Task<IActionResult> EditMatterDocument([FromBody] List<GlobalData> gdatas = null)
        {
            //try
            {
                #region wopi-headers
                var headers = new RequestProcessHeader(HttpContext.Request);
                //var Authorization = HttpContext.Request.Headers[WopiHeaders.Authorization];
                //string WopiSrc = HttpContext.Request.Headers[WopiHeaders.WopiSrc].FirstOrDefault();//http://wopi.hoxro.com/wopi/files/
                //var Action = HttpContext.Request.Headers[WopiHeaders.Action];//action = "Edit";
                //var IsMatterDocument = HttpContext.Request.Headers[WopiHeaders.IsMatterDocument];//MatterDocument on not?
                //var MatterDocumentId = HttpContext.Request.Headers[WopiHeaders.MatterDocumentId];//MatterDocument MatterDocumentId
                //var ParentId = HttpContext.Request.Headers[WopiHeaders.ParentId];//MatterDocument ParentId
                //var VersionNo = HttpContext.Request.Headers[WopiHeaders.VersionNo];//MatterDocument Version
                //var StorageName = HttpContext.Request.Headers[WopiHeaders.StorageName];//Azure Container Name
                //var Checked = HttpContext.Request.Headers[WopiHeaders.Checked];
                //var SourceBlob = HttpContext.Request.Headers[WopiHeaders.SourceBlob];//matter/{MatterRef}/{SourceFileName}
                //var SourceTemplateURL = HttpContext.Request.Headers[WopiHeaders.SourceTemplateURL];
                //var NewBlob = HttpContext.Request.Headers[WopiHeaders.NewBlob];//matter/{MatterRef}/{NewFileName}
                //var OperationType = HttpContext.Request.Headers[WopiHeaders.OperationType];
                //var NewDocumentName = HttpContext.Request.Headers[WopiHeaders.NewDocumentName];//NewFileName
                //var FileName = HttpContext.Request.Headers[WopiHeaders.FileName];//NewFileName
                //var Extenssion = HttpContext.Request.Headers[WopiHeaders.Extenssion];
                //var Override = HttpContext.Request.Headers[WopiHeaders.Override];
                //var userName = HttpContext.Request.Headers[WopiHeaders.UserName];//UserName/Email
                //var userFullName = HttpContext.Request.Headers[WopiHeaders.UserFullName];//UserFullName
                ////var fileOriginalName = HttpContext.Request.Headers[WopiHeaders.FileOriginalName];//Original Name of the file
                #endregion wopi-headers

                var extenssion = headers.Extenssion.ToString().Replace(".", "");
                WopiActionEnum actionWopi = (WopiActionEnum)Enum.Parse(typeof(WopiActionEnum), headers.Action);
                headers.FileName = headers.FileName.ToString() == string.Empty ? headers.NewDocumentName : headers.FileName;
                bool isChecked = headers.Checked != string.Empty && Convert.ToBoolean(headers.Checked);
                byte[] bytes = System.Text.Encoding.UTF8.GetBytes($"{ROOT_PATH}{Path.GetFileName(headers.FileName)}");
                var fileIdentifier = Convert.ToBase64String(bytes);

                //chek file edit mode
                FileStorageInfo liteDbFileInfo = new();

                var address = string.Empty;
                if (headers.IsMatterDocument == "True" && headers.Action == "Edit")
                {
                    //address = NewSourceURL;
                    address = headers.NewDocumentName;
                    liteDbFileInfo = await _liteDbManager.CheckVersion(Convert.ToInt32(headers.ParentId), Convert.ToInt32(headers.VersionNo));
                }
                else
                {
                    //New file lication
                    address = headers.NewBlob == string.Empty ? headers.SourceBlob : headers.NewBlob;
                }
                // end with "/" come from API
                headers.WopiSrc += fileIdentifier;
                if (ValidateAuthorizationHeader(headers.Authorization))
                {
                    //TODO: supply user
                    var user = headers.UserName;
                    string urlsr = await UrlGenerator.GetFileUrlAsync(extenssion, headers.WopiSrc, actionWopi);
                    var path = Path.Combine(WopiHostOptions.Value.WebRootPath, (WopiHostOptions.Value.WopiRootPath));
                    path = path.AddSourceAddress(address);
                    var pathTemplate = string.Empty;
                    bool willDownload = true;
                    bool isFileExits = System.IO.File.Exists($"{WopiHostOptions.Value.WebRootPath}\\{WopiHostOptions.Value.WopiRootPath}\\{headers.FileName}");
                    //if (willDownload && (liteDbFileInfo?.IsEditing == true? (isFileExits ? false : true) : true))
                    if (willDownload && (liteDbFileInfo is null || !liteDbFileInfo?.IsEditing == true))
                    {
                        var status = await DownloadFile(path, headers.StorageName, headers.SourceBlob, headers.NewBlob);
                    }
                    else if (willDownload && liteDbFileInfo?.IsEditing == true && !isFileExits)
                    {
                        var status = await DownloadFile(path, headers.StorageName, headers.SourceBlob, headers.NewBlob);
                        //var status = DownloadFile($"{WopiHostOptions.Value.WebRootPath}\\{WopiHostOptions.Value.WopiRootPath}\\{SourceContentURL}", Container, SourceContentURL, NewSourceURL);
                    }

                    if (headers.Action == "Edit" && liteDbFileInfo?.IsEditing != true)
                    {
                        var itemFileStorageInfo = new FileStorageInfo()
                        {
                            UserId =headers.UserName,
                            Blob = headers.SourceBlob,
                            Container = headers.StorageName,
                            Id = fileIdentifier,
                            IsOverride = Convert.ToBoolean(headers.Override),
                            LocalPath = path,
                            BlobNew = headers.NewBlob,
                            BlobTemplate = headers.SourceTemplateURL,
                            IsUploaded = false,
                            StartTime = DateTime.Now,
                            FileName = headers.FileName,
                            NewFileName = headers.NewDocumentName
                        };
                        if (headers.IsMatterDocument == "True")
                        {
                            itemFileStorageInfo.IsEditing = true;
                            itemFileStorageInfo.MatterDocumentId = !string.IsNullOrEmpty(headers.MatterDocumentId) ? Convert.ToInt32(headers.MatterDocumentId) : 0;
                            itemFileStorageInfo.ParentId = !string.IsNullOrEmpty(headers.ParentId) ? Convert.ToInt32(headers.ParentId) : 0;
                            itemFileStorageInfo.VersionNo = !string.IsNullOrEmpty(headers.VersionNo) ? Convert.ToInt32(headers.VersionNo) + 1 : 0;
                        }
                        await _liteDbManager.InsertRequest(itemFileStorageInfo);
                    }
                    else if (headers.Action == "View" && Convert.ToBoolean(isChecked))
                    {
                        var fileStorageInfo = _liteDbManager.Get(fileIdentifier).Result;
                        if (fileStorageInfo is not null && !fileStorageInfo.IsUploaded)
                        {
                            string output = "empty";
                            return new JsonResult(output);
                        }
                    }
                    //if file is matter document 
                    if (headers.IsMatterDocument == "True" && headers.Action == "Edit")
                    {
                        var url = (JsonResult)ReturnAction(urlsr, user,headers.UserFullName, actionWopi, fileOriginalName: headers?.FileOriginalName ?? string.Empty);
                        NewLetterResponse urlResponse = new()
                        {
                            Url = url?.Value?.ToString(),
                            NewFileName = headers.FileName.ToString()
                        };
                        return Ok(urlResponse);
                    }
                    return await ReturnActionAsync(urlsr, user, headers.UserFullName, actionWopi, fileOriginalName: headers?.FileOriginalName ?? string.Empty);
                }
                else
                {
                    return new UnauthorizedResult();
                }
            }
            catch (Exception ex)
            {
                string txtPath = $"{WopiHostOptions.Value.WebRootPath}\\log\\log.txt";
                if (System.IO.File.Exists(txtPath))
                {
                    using (StreamWriter sw = new(txtPath))
                    {
                        sw.WriteLine($"{DateTime.UtcNow}(UTC)-->(MakeURL) {ex.Message}");
                        //sw.WriteLine(ex.StackTrace);
                    }
                }
                return ReturnException(ex);
            }
        }*/
        [HttpPost("EditMatterDocument")]
        public async Task<IActionResult> EditMatterDocument([FromBody] List<GlobalData> gdatas = null)
        {
            try
            {
                #region wopi-headers
                var headers = new RequestProcessHeader(HttpContext.Request);
                #endregion wopi-headers

                var extenssion = headers.Extenssion.ToString().Replace(".", "");
                WopiActionEnum actionWopi = (WopiActionEnum)Enum.Parse(typeof(WopiActionEnum), headers.Action);
                headers.FileName = headers.FileName.ToString() == string.Empty ? headers.NewDocumentName : headers.FileName;
                bool isChecked = headers.Checked != string.Empty && Convert.ToBoolean(headers.Checked);
                byte[] bytes = System.Text.Encoding.UTF8.GetBytes($"{ROOT_PATH}{Path.GetFileName(headers.FileName)}");
                var fileIdentifier = Convert.ToBase64String(bytes);

                //chek file edit mode
                FileStorageInfo liteDbFileInfo = new();

                var address = string.Empty;
                if (headers.IsMatterDocument == "True" && headers.Action == "Edit")
                {
                    //address = NewSourceURL;
                    address = headers.NewDocumentName;
                    //liteDbFileInfo = await _liteDbManager.CheckVersion(Convert.ToInt32(headers.ParentId), Convert.ToInt32(headers.VersionNo));
                }
                else
                {
                    //New file lication
                    address = headers.NewBlob == string.Empty ? headers.SourceBlob : headers.NewBlob;
                }
                // end with "/" come from API
                headers.WopiSrc += fileIdentifier;
                if (ValidateAuthorizationHeader(headers.Authorization))
                {
                    //TODO: supply user
                    var user = headers.UserName;
                    string urlsr = await UrlGenerator.GetFileUrlAsync(extenssion, headers.WopiSrc, actionWopi);
                    var path = Path.Combine(WopiHostOptions.Value.WebRootPath, (WopiHostOptions.Value.WopiRootPath));
                    path = path.AddSourceAddress(address);
                    var pathTemplate = string.Empty;
                    //bool willDownload = true;
                    bool isFileExits = System.IO.File.Exists($"{WopiHostOptions.Value.WebRootPath}\\{WopiHostOptions.Value.WopiRootPath}\\{headers.FileName}");
                    //if (willDownload && (liteDbFileInfo?.IsEditing == true? (isFileExits ? false : true) : true))
                    /*if (willDownload && (liteDbFileInfo is null || !liteDbFileInfo?.IsEditing == true))
                    {
                        var status = await DownloadFile(path, headers.StorageName, headers.SourceBlob, headers.NewBlob);
                    }
                    else if (willDownload && liteDbFileInfo?.IsEditing == true && !isFileExits)
                    {
                        //var status = DownloadFile($"{WopiHostOptions.Value.WebRootPath}\\{WopiHostOptions.Value.WopiRootPath}\\{SourceContentURL}", Container, SourceContentURL, NewSourceURL);
                    }*/

                    var status = await DownloadFile(path, headers.StorageName, headers.SourceBlob, headers.NewBlob);

                    /*if (headers.Action == "Edit" && liteDbFileInfo?.IsEditing != true)
                    {
                        var itemFileStorageInfo = new FileStorageInfo()
                        {
                            UserId = headers.UserName,
                            Blob = headers.SourceBlob,
                            Container = headers.StorageName,
                            Id = fileIdentifier,
                            IsOverride = Convert.ToBoolean(headers.Override),
                            LocalPath = path,
                            BlobNew = headers.NewBlob,
                            BlobTemplate = headers.SourceTemplateURL,
                            IsUploaded = false,
                            StartTime = DateTime.Now,
                            FileName = headers.FileName,
                            NewFileName = headers.NewDocumentName
                        };
                        if (headers.IsMatterDocument == "True")
                        {
                            itemFileStorageInfo.IsEditing = true;
                            itemFileStorageInfo.MatterDocumentId = !string.IsNullOrEmpty(headers.MatterDocumentId) ? Convert.ToInt32(headers.MatterDocumentId) : 0;
                            itemFileStorageInfo.ParentId = !string.IsNullOrEmpty(headers.ParentId) ? Convert.ToInt32(headers.ParentId) : 0;
                            itemFileStorageInfo.VersionNo = !string.IsNullOrEmpty(headers.VersionNo) ? Convert.ToInt32(headers.VersionNo) + 1 : 0;
                        }
                        await _liteDbManager.InsertRequest(itemFileStorageInfo);
                    }*/
                    /*else if (headers.Action == "View" && Convert.ToBoolean(isChecked))
                    {
                        var fileStorageInfo = _liteDbManager.Get(fileIdentifier).Result;
                        if (fileStorageInfo is not null && !fileStorageInfo.IsUploaded)
                        {
                            string output = "empty";
                            return new JsonResult(output);
                        }
                    }*/
                    WopiClaimTypesVm claimTypesVm = new()
                    {
                        UrlSr = urlsr,
                        UserName = user,
                        UserFullName = headers.UserFullName,
                        WopiActionEnum = actionWopi,
                        FileOriginalName = headers.FileOriginalName ?? string.Empty,
                    };
                    var url = (JsonResult)ReturnAction(claimTypesVm);
                    NewLetterResponse urlResponse = new()
                    {
                        Url = url?.Value?.ToString(),
                        NewFileName = headers.FileName.ToString()
                    };
                    return Ok(urlResponse);

                    /*//if file is matter document 
                    if (headers.IsMatterDocument == "True" && headers.Action == "Edit")
                    {
                       
                    }
                    return await ReturnActionAsync(urlsr, user, headers.UserFullName, actionWopi, fileOriginalName: headers?.FileOriginalName ?? string.Empty);
                    */
                }
                else
                {
                    return new UnauthorizedResult();
                }
            }
            catch (Exception ex)
            {
                string txtPath = $"{WopiHostOptions.Value.WebRootPath}\\log\\log.txt";
                if (System.IO.File.Exists(txtPath))
                {
                    using (StreamWriter sw = new(txtPath))
                    {
                        sw.WriteLine($"{DateTime.UtcNow}(UTC)-->(MakeURL) {ex.Message}");
                        //sw.WriteLine(ex.StackTrace);
                    }
                }
                return ReturnException(ex);
            }
        }
        [HttpPost("MergeTemplate")]
        public async Task<NewLetterResponse> MergeTemplate([FromBody] List<GlobalDataVm> gdatas = null)
        {
            NewLetterResponse res = new();
            try
            {
                //DateTime dateTime1 = DateTime.Now;
                var headers = new RequestProcessHeader(HttpContext.Request);
                WopiActionEnum actionWopi = (WopiActionEnum)Enum.Parse(typeof(WopiActionEnum), headers.Action);
                string newFileName = $"{Guid.NewGuid()}.{headers.Extenssion ?? "docx"}";
                string wopiRootPath = Path.Combine(WopiHostOptions.Value.WebRootPath, (WopiHostOptions.Value.WopiRootPath));
                var contentPath = wopiRootPath.AddSourceAddress(headers.SourceBlob);//content
                var templatePath = wopiRootPath.AddSourceAddress(headers.SourceTemplateURL);//template
                var signaturePath = wopiRootPath.AddSourceAddress(headers.SignatureUrl);//signature template

                var newFilePath = wopiRootPath.AddSourceAddress(newFileName);
                if (headers.SourceTemplateURL != "undefined")
                {
                    List<Task> downloadTasks = new();
                    /*{
                        Task.Run(async () =>
                        {
                            // content
                            await DownloadFile(contentPath, headers.StorageName, headers.SourceBlob, headers.NewBlob);
                            // template
                            await DownloadFile(templatePath, headers.StorageName, headers.SourceTemplateURL, headers.SourceTemplateURL);
                            // signature
                            if (!string.IsNullOrEmpty(headers.SignatureUrl))
                            {
                                await DownloadFile(signaturePath, headers.StorageName, headers.SignatureUrl, headers.SignatureUrl);
                            }
                        })
                    };*/

                    // content
                    downloadTasks.Add(DownloadFile(contentPath, headers.StorageName, headers.SourceBlob, headers.NewBlob));
                    // template
                    downloadTasks.Add(DownloadFile(templatePath, headers.StorageName, headers.SourceTemplateURL, headers.SourceTemplateURL));
                    // signature
                    if (!string.IsNullOrEmpty(headers.SignatureUrl))
                    {
                        downloadTasks.Add(DownloadFile(signaturePath, headers.StorageName, headers.SignatureUrl, headers.SignatureUrl));
                    }
                    //download parallel
                    await Task.WhenAll(downloadTasks);
                    //var res = ProcessXmlDocument.MergeDocuments(tempPath, contentPath, pathTemplate, gdatas, out unMergeGlobalData);
                    res = await ProcessBulkXMLMailMerge.MergeDocuments(templatePath, contentPath, signaturePath, newFilePath, gdatas, headers.SkipFieldIds);
                    if (res.StatusCode != 200)
                    {
                        return await Task.FromResult(res);
                    }
                }
                /*var bytes = System.Text.Encoding.UTF8.GetBytes($"{ROOT_PATH}{Path.GetFileName(newFileName)}");
                var fileIdentifier = Convert.ToBase64String(bytes);
                headers.WopiSrc += fileIdentifier;
                string urlsr = await UrlGenerator.GetFileUrlAsync(headers.Extenssion, headers.WopiSrc, actionWopi);

                var fileInfo = new FileStorageInfo(headers, fileIdentifier, newFilePath);
                //await _liteDbManager.InsertRequest(fileInfo);
                //return await ReturnActionAsync(urlsr, headers.UserName, headers.UserFullName, headers.WopiSource, actionWopi, WopiDocsTypeEnum.WopiDocs, WopiVersionableTypeEnum.False);
                res.Url = await ReturnOOSUrlStringAsync(urlsr, headers.UserName, headers.UserFullName, headers.WopiSrc, actionWopi);
                //res.NewFileName = Path.GetFileNameWithoutExtension(newFilePath);*/
                headers.FileName = newFilePath.GetFileName();
                res.Url = await GetFileUrlAsync(headers);
                res.NewFileName = Path.GetFileName(newFilePath);
                return res;
            }
            catch (Exception ex)
            {
                string txtPath = $"{WopiHostOptions.Value.WebRootPath}\\log\\log.txt";
                if (System.IO.File.Exists(txtPath))
                {
                    System.IO.File.AppendAllText(txtPath, $"{DateTime.UtcNow}(UTC)-->(MergeTemplate) {ex.Message}{Environment.NewLine}");
                }
                res.Message = ex.Message.ToString();
                res.StatusCode = 500;
                return res;
            }
        }

        [HttpPost("MergeOnExistingDocument")]
        public async Task<NewLetterResponse> MergeOnExistingDocument([FromBody] List<GlobalDataVm> gDatas = null)
        {
            NewLetterResponse urlResponse = new();
            try
            {
                var headers = new RequestProcessHeader(HttpContext.Request);
                string newFileName = headers.NewDocumentName;
                if (string.IsNullOrEmpty(newFileName))
                {
                    newFileName = $"{Guid.NewGuid()}.{headers.Extenssion ?? "docx"}";
                }
                string wopiRootPath = Path.Combine(WopiHostOptions.Value.WebRootPath, WopiHostOptions.Value.WopiRootPath);
                string newFilePath = wopiRootPath.AddSourceAddress(newFileName);
                string existingDocumentPath = $"{WopiHostOptions.Value.WebRootPath}\\{WopiHostOptions.Value.WopiRootPath}\\{Path.GetFileName(headers.SourceBlob)}";
                if (!System.IO.File.Exists(newFilePath))
                {
                    await DownloadFile(newFilePath, headers.StorageName, headers.SourceBlob, existingDocumentPath);
                }
                Regex regex = new(MailMergeRegex.MailMergePattern, RegexOptions.Compiled);
                ProcessBulkXMLMailMerge.MailMergeWithGlobalDataAndRemoveHighlight(newFilePath, regex, MailMergeRegex.MailMergePattern, MailMergeRegex.StartText, MailMergeRegex.EndText, gDatas, urlResponse, headers.SkipFieldIds);
                /*WopiActionEnum wopiAction = (WopiActionEnum)Enum.Parse(typeof(WopiActionEnum), !string.IsNullOrEmpty(headers.Action) ? headers.Action : "Edit");
                var bytes = System.Text.Encoding.UTF8.GetBytes($"{ROOT_PATH}{Path.GetFileName(newFilePath)}");
                var fileIdentifier = Convert.ToBase64String(bytes);
                headers.WopiSrc += fileIdentifier;
                string urlsr = await UrlGenerator.GetFileUrlAsync(headers.Extenssion, headers.WopiSrc, wopiAction);
                urlResponse.Url = await ReturnOOSUrlStringAsync(urlsr, headers.UserName, headers.UserFullName, headers.WopiSrc, wopiAction);
                urlResponse.NewFileName = Path.GetFileName(newFilePath);*/
                headers.FileName = newFilePath.GetFileName();
                urlResponse.NewFileName = Path.GetFileName(newFilePath);
                urlResponse.Url = await GetFileUrlAsync(headers);
            }
            catch (Exception ex)
            {
                if (System.IO.File.Exists($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt"))
                {
                    System.IO.File.AppendAllText($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt", $"ReplaceSignatureOnExistingFile: {DateTime.UtcNow}->{ex.Message} {Environment.NewLine}");
                }
                await ReturnExceptionAsync(ex, urlResponse);
            }
            return await Task.FromResult(urlResponse);
        }
        #region Private Methods

        /*private void SetFileAttribute(string location, FileAttributes attribute)
        {
            System.IO.File.SetAttributes(location, attribute);
        }*/
        private IActionResult ReturnAction(WopiClaimTypesVm claimTypesVm)
        {
            //urlsr = ReplaceUI_LLCC(urlsr);
            var token = SecurityHandler.GenerateAccessToken(0, claimTypesVm.UserName, claimTypesVm.UserFullName, fileOriginalName: claimTypesVm.FileOriginalName);
            var AccessTokenInfo = new AccessTokenInfo
            {
                AccessToken = SecurityHandler.WriteToken(token),
                AccessTokenExpiry = token.ValidTo.ToUnixTimestamp()
            };
            //If action == Edit then access_token_ttl has value otherwise = 0
            var output = $"{claimTypesVm.UrlSr}&access_token={AccessTokenInfo.AccessToken}&access_token_ttl={(claimTypesVm.WopiActionEnum == WopiActionEnum.Edit ? SecurityHandler.GenerateAccessToken_TTL() : 0)}";
            return new JsonResult(output);
        }
        /*private async Task<IActionResult> ReturnActionAsync(string urlsr, string username, string userFullName, WopiActionEnum action, WopiDocsTypeEnum wopiDocsType = WopiDocsTypeEnum.WopiDocs, WopiVersionableTypeEnum isVersionable = WopiVersionableTypeEnum.True, string fileOriginalName = "")
        {
            //urlsr = ReplaceUI_LLCC(urlsr);
            var token = SecurityHandler.GenerateAccessToken(0, username, userFullName, wopiDocsType, isVersionable, fileOriginalName: fileOriginalName);
            var AccessTokenInfo = new AccessTokenInfo
            {
                AccessToken = SecurityHandler.WriteToken(token),
                AccessTokenExpiry = token.ValidTo.ToUnixTimestamp()
            };
            //If action == Edit then access_token_ttl has value otherwise = 0
            var output = $"{urlsr}&access_token={AccessTokenInfo.AccessToken}&access_token_ttl={(action == WopiActionEnum.Edit ? SecurityHandler.GenerateAccessToken_TTL() : 0)}";
            return await Task.FromResult(new JsonResult(output));
        }*/
        private async Task<string> ReturnOOSUrlStringAsync(string urlsr, string username, string userFullName, string wopiSrc, WopiActionEnum action, WopiDocsTypeEnum wopiDocsType = WopiDocsTypeEnum.WopiDocs, WopiVersionableTypeEnum isVersionable = WopiVersionableTypeEnum.True, string fileOriginalName = "")
        {
            var token = SecurityHandler.GenerateAccessToken(0, username, userFullName, wopiDocsType: wopiDocsType, isVersionable: isVersionable, fileOriginalName: fileOriginalName, wopiSrc: wopiSrc);
            var AccessTokenInfo = new AccessTokenInfo
            {
                AccessToken = SecurityHandler.WriteToken(token),
                AccessTokenExpiry = token.ValidTo.ToUnixTimestamp()
            };
            //If action == Edit then access_token_ttl has value otherwise = 0
            var output = $"{urlsr}&access_token={AccessTokenInfo.AccessToken}&access_token_ttl={(action == WopiActionEnum.Edit ? SecurityHandler.GenerateAccessToken_TTL() : 0)}";
            return await Task.FromResult(output);
        }

        private async Task<string> GetFileUrlAsync(RequestProcessHeader headers, WopiDocsTypeEnum wopiDocsType = WopiDocsTypeEnum.WopiDocs, WopiVersionableTypeEnum isVersionable = WopiVersionableTypeEnum.True)
        {
            WopiActionEnum wopiAction = (WopiActionEnum)Enum.Parse(typeof(WopiActionEnum), !string.IsNullOrEmpty(headers.Action) ? headers.Action : nameof(WopiActionEnum.View));
            var bytes = System.Text.Encoding.UTF8.GetBytes($"{ROOT_PATH}{Path.GetFileName(headers.FileName)}");
            var fileIdentifier = Convert.ToBase64String(bytes);
            headers.WopiSrc += fileIdentifier;
            string urlsr = await UrlGenerator.GetFileUrlAsync(headers.Extenssion, headers.WopiSrc, wopiAction);
            return await ReturnOOSUrlStringAsync(urlsr, headers.UserName, headers.UserFullName, headers.WopiSrc, wopiAction, wopiDocsType: wopiDocsType, fileOriginalName: headers?.FileOriginalName ?? string.Empty);
        }

        private async Task<CheckFileInfoVm> GetFileMultipleUrlAsync(RequestProcessHeader headers, WopiDocsTypeEnum wopiDocsType = WopiDocsTypeEnum.WopiDocs, WopiVersionableTypeEnum isVersionable = WopiVersionableTypeEnum.True)
        {
            CheckFileInfoVm fileInfoVm = new();
            var bytes = System.Text.Encoding.UTF8.GetBytes($"{ROOT_PATH}{Path.GetFileName(headers.FileName)}");
            var fileIdentifier = Convert.ToBase64String(bytes);
            headers.WopiSrc += fileIdentifier;

            WopiActionEnum wopiAction = WopiActionEnum.View;
            string urlsr = await UrlGenerator.GetFileUrlAsync(headers.Extenssion, headers.WopiSrc, wopiAction);
            fileInfoVm.HostViewUrl = await ReturnOOSUrlStringAsync(urlsr, headers.UserName, headers.UserFullName, headers.WopiSrc, wopiAction, wopiDocsType: wopiDocsType, fileOriginalName: headers?.FileOriginalName ?? string.Empty);

            wopiAction = WopiActionEnum.Edit;
            urlsr = await UrlGenerator.GetFileUrlAsync(headers.Extenssion, headers.WopiSrc, wopiAction);
            fileInfoVm.HostEditUrl = await ReturnOOSUrlStringAsync(urlsr, headers.UserName, headers.UserFullName, headers.WopiSrc, wopiAction, wopiDocsType: wopiDocsType, fileOriginalName: headers?.FileOriginalName ?? string.Empty);
            return fileInfoVm;
        }

        private IActionResult ReturnException(Exception ex)
        {
            if (ex.Message.Contains("The specified container does not exist"))
            {
                return StatusCode(404, "The specified blob does not exist");
            }
            else if (ex.Message.Contains("Could not find file"))
            {
                return StatusCode(404, "Could not find file");
            }
            return StatusCode(500, "internal server error");
        }
        /// <summary>
        /// return Generic exception. 
        /// T = response model class, 
        /// T class must have "Message" and "StatusCode" property
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="ex"></param>
        /// <param name="response"></param>
        /// <returns></returns>
        private async Task ReturnExceptionAsync<T>(Exception ex, T response) where T : class
        {
            Type typeT = response?.GetType();
            PropertyInfo propertyTMessage = typeT?.GetProperty("Message");
            PropertyInfo propertyTStatusCode = typeT?.GetProperty("StatusCode");
            if (ex.Message.Contains("The specified container does not exist"))
            {
                propertyTMessage?.SetValue(response, "The specified blob does not exist");
                propertyTStatusCode?.SetValue(response, 404);
            }
            else if (ex.Message.Contains("Could not find file"))
            {
                propertyTMessage?.SetValue(response, "Could not find file");
                propertyTStatusCode?.SetValue(response, 404);
            }
            propertyTMessage?.SetValue(response, "internal server error");
            propertyTStatusCode?.SetValue(response, 500);
            await Task.FromResult(response);
        }
        //view download
        private async Task ViewDownloadFile(string path, string container, string sourceUrl, bool isCheck)
        {
            if (!System.IO.File.Exists(path))
            {
                await _azureBlobStorage.DownloadAsync(container, sourceUrl, path);
            }
            else
            {
                await _azureBlobStorage.UploadAsync(container, sourceUrl, path);
            }
        }
        //edit download
        private async Task EditDownloadFile(string path, string container, string sourceUrl, bool isCheck)
        {
            if (!System.IO.File.Exists(path))
            {
                await _azureBlobStorage.DownloadAsync(container, sourceUrl, path);
            }
            else
            {
                await _azureBlobStorage.UploadAsync(container, sourceUrl, path);
            }
        }
        private async Task DownloadFile(string path, string container, string sourceUrl)
        {
            await _azureBlobStorage.DownloadAsync(container, sourceUrl, path);
        }
        /// <summary>
        /// Download Email and others files
        /// </summary>
        /// <param name="path"></param>
        /// <param name="container"></param>
        /// <param name="sourceUrl"></param>
        /// <param name="isCheck"></param>
        private async Task DownloadFileEmailAndOthers(string path, string container, string sourceUrl, bool isCheck)
        {
            if (!System.IO.File.Exists(path))
            {
                await _azureBlobStorage.DownloadAsync(container, sourceUrl, path);
            }
        }

        private async Task<bool> DownloadFile(string path, string Container, string SourceURL, string NewSourceURL)
        {
            var output = false;
            if (!System.IO.File.Exists(path))
            {
                //return new JsonResult("false");
                var isExists = await _azureBlobStorage.ExistsAsync(Container, SourceURL);
                if (isExists)
                {
                    await _azureBlobStorage.DownloadAsync(Container, SourceURL, path);
                }
                else
                {
                    try
                    {
                        path = Path.Combine(WopiHostOptions.Value.WebRootPath, WopiHostOptions.Value.WopiRootPath);
                        string uploadPath = path.AddSourceAddress(NewSourceURL);
                        path = path.AddSourceAddress(SourceURL);
                        if (System.IO.File.Exists(uploadPath))
                        {
                            await _azureBlobStorage.UploadAsync(Container, SourceURL, uploadPath);//.Wait();
                            System.IO.File.Move(uploadPath, path);
                        }
                        else if (System.IO.File.Exists(path))//if uploadPath file not exist
                        {
                            System.IO.File.Copy(path, uploadPath);
                        }
                    }
                    catch (Exception)
                    {

                    }
                }
            }
            else
            {
                var isExists = await _azureBlobStorage.ExistsAsync(Container, SourceURL);
                if (!isExists)
                {
                    try
                    {
                        await _azureBlobStorage.UploadAsync(Container, SourceURL, path);//.Wait();
                        output = true;
                    }
                    catch (Exception)
                    {
                        //var fileStorageInfo = _liteDbManager.Get(fileIdentifier).Result;
                        //FilesController.LockStorage.Remove(fileStorageInfo.Id);
                        //_azureBlobStorage.UploadAsync(Container, SourceURL, path).Wait();
                    }

                }
            }
            return await Task.FromResult(output);
        }

        private bool ValidateAuthorizationHeader(StringValues authorizationHeader)
        {
            //TODO: implement header validation http://wopi.readthedocs.io/projects/wopirest/en/latest/bootstrapper/GetRootContainer.html#sample-response
            // http://stackoverflow.com/questions/31948426/oauth-bearer-token-authentication-is-not-passing-signature-validation
            return true;
        }

        private async Task<WopiUrlBuilder> NewOOSWopiUrlBuilder()
        {
            //WopiClientUrl = Configuration.GetValue("WopiClientUrl", string.Empty);
            string WopiClientUrlNew = WopiHostOptions.Value.WopiClientUrl; //Configuration.GetValue("WopiClientUrl", string.Empty);
            WopiDiscoverer DiscovererNew = new WopiDiscoverer(new HttpDiscoveryFileProvider(WopiClientUrlNew, _clientFactory));
            return await Task.FromResult(new WopiUrlBuilder(DiscovererNew, new WopiUrlSettings { UI_LLCC = new CultureInfo(UI_LLCC == "en-GB" ? "en-US" : UI_LLCC) }));
        }

        #endregion

        #region Matter Work Flow 

        [HttpPost("MatterUnMailMergeList")]
        public async Task<WopiResponseUnMailMergeVM> MatterUnMailMergeList(WopiRequestUnMailMergeVM wopiRequestUnMailMergeVM)
        {
            WopiResponseUnMailMergeVM wopiResponse = new WopiResponseUnMailMergeVM();
            try
            {
                var headers = new RequestProcessHeader(HttpContext.Request);
                wopiResponse.ContentDocumnets = await _documentService.UnMailMergeList(headers, wopiRequestUnMailMergeVM, true);
            }
            catch (Exception ex)
            {
                if (System.IO.File.Exists($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt"))
                {
                    System.IO.File.AppendAllText($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt", $"MatterUnMailMergeList: {DateTime.UtcNow.ToString()}->{ex.Message.ToString()} {Environment.NewLine}");
                }
            }
            return wopiResponse;
        }

        [HttpPost("MakeMatterStageActivityMailMerge")]
        public async Task<MatterStageActivityMailMergeResponse> MakeMatterStageActivityMailMerge(MatterStageActivityMailMergeRequest matterStageActivity)
        {
            MatterStageActivityMailMergeResponse wopiResponse = new MatterStageActivityMailMergeResponse();
            try
            {
                var headers = new RequestProcessHeader(HttpContext.Request);
                wopiResponse.DocumentContents = await _documentService.MatterStageDocumentMailMerge(headers, matterStageActivity);
            }
            catch (Exception ex)
            {
                if (System.IO.File.Exists($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt"))
                {
                    System.IO.File.AppendAllText($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt", $"MakeMatterStageActivityMailMerge: {DateTime.UtcNow.ToString()}->{ex.Message.ToString()} {Environment.NewLine}");
                }
            }
            return wopiResponse;
        }
        #endregion

        #region Get Document UnMailMergeList
        [HttpPost("UnMailMergeList")]
        public async Task<WopiResponseUnMailMergeVM> UnMailMergeList(WopiRequestUnMailMergeVM wopiRequestUnMailMergeVM)
        {
            WopiResponseUnMailMergeVM wopiResponse = new WopiResponseUnMailMergeVM();
            try
            {
                var headers = new RequestProcessHeader(HttpContext.Request);
                wopiResponse.ContentDocumnets = await _documentService.UnMailMergeList(headers, wopiRequestUnMailMergeVM, false);
            }
            catch (Exception ex)
            {
                if (System.IO.File.Exists($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt"))
                {
                    System.IO.File.AppendAllText($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt", $"UnMailMergeList: {DateTime.UtcNow.ToString()}->{ex.Message.ToString()} {Environment.NewLine}");
                }
            }
            return wopiResponse;
        }

        #endregion

        #region Invoice Template Merge
        [HttpPost("InvoiceTemplateMerge")]
        public async Task<IActionResult> InvoiceTemplateMerge([FromBody] InvoiceTemplateMergeVM mergeVM)
        {
            WopiUrlResponse urlResponse = new();
            try
            {
                string fileName = $"{Guid.NewGuid()}.docx";
                string path = $"{WopiHostOptions.Value.WebRootPath}\\{WopiHostOptions.Value.WopiRootPath}\\{fileName}";
                var headers = new RequestProcessHeader(HttpContext.Request);
                await DownloadFile(path, headers.StorageName, headers.SourceBlob);
                urlResponse = ProcessInvoiceTemplateMerge.InvoiceTemplateMerge(path, mergeVM.InvoiceVM, mergeVM.GDatas);
                headers.FileName = fileName;
                /*WopiActionEnum wopiAction = (WopiActionEnum)Enum.Parse(typeof(WopiActionEnum), !string.IsNullOrEmpty(headers.Action) ? headers.Action : "View");
                var bytes = System.Text.Encoding.UTF8.GetBytes($"{ROOT_PATH}{Path.GetFileName(fileName)}");
                var fileIdentifier = Convert.ToBase64String(bytes);
                headers.WopiSrc += fileIdentifier;
                string urlsr = await UrlGenerator.GetFileUrlAsync(headers.Extenssion, headers.WopiSrc, wopiAction);
                urlResponse.Url = await ReturnOOSUrlStringAsync(urlsr, headers.UserName, headers.UserFullName, headers.WopiSrc, wopiAction);*/
                CheckFileInfoVm multipleUrl = await GetFileMultipleUrlAsync(headers);
                urlResponse.Url = multipleUrl.HostViewUrl;
                urlResponse.EditUrl = multipleUrl.HostEditUrl;
                urlResponse.FileName = fileName;
            }
            catch (Exception ex)
            {
                if (System.IO.File.Exists($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt"))
                {
                    System.IO.File.AppendAllText($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt", $"InvoiceTemplateMerge: {DateTime.UtcNow}->{ex.Message} {Environment.NewLine}");
                }
                await ReturnExceptionAsync(ex, urlResponse);
            }
            return Ok(urlResponse);
        }
        [HttpGet("InvoiceDownload/{fileName}")]
        public async Task<IActionResult> InvoiceDownload(string fileName)
        {
            try
            {
                var headers = new RequestProcessHeader(HttpContext.Request);
                string path = $"{WopiHostOptions.Value.WebRootPath}\\{WopiHostOptions.Value.WopiRootPath}\\{fileName}";
                await _azureBlobStorage.UploadAsync(headers.StorageName, headers.SourceBlob, path);
                return Ok(true);
            }
            catch (Exception)
            {
                return Ok(false);
            }
        }
        #endregion

        #region Signature
        [HttpPost("ReplaceSignatureOnExistingFile")]
        public async Task<IActionResult> ReplaceSignatureOnExistingFile()
        {
            NewLetterResponse urlResponse = new();
            try
            {
                var headers = new RequestProcessHeader(HttpContext.Request);
                string signaturePath = $"{WopiHostOptions.Value.WebRootPath}\\{WopiHostOptions.Value.WopiRootPath}\\{Path.GetFileName(headers.SignatureUrl)}";
                string filePath = $"{WopiHostOptions.Value.WebRootPath}\\{WopiHostOptions.Value.WopiRootPath}\\{headers.NewDocumentName}";
                await DownloadFile(signaturePath, headers.StorageName, headers.SignatureUrl);
                //replace signature
                urlResponse = ProcessBulkXMLMailMerge.ReplaceSignatureOnExistingFile(filePath, signaturePath);
                if (urlResponse.StatusCode == 200)
                {
                    WopiActionEnum wopiAction = (WopiActionEnum)Enum.Parse(typeof(WopiActionEnum), !string.IsNullOrEmpty(headers.Action) ? headers.Action : "View");
                    /*var bytes = System.Text.Encoding.UTF8.GetBytes($"{ROOT_PATH}{Path.GetFileName(filePath)}");
                    var fileIdentifier = Convert.ToBase64String(bytes);
                    headers.WopiSrc += fileIdentifier;
                    string urlsr = await UrlGenerator.GetFileUrlAsync(headers.Extenssion, headers.WopiSrc, wopiAction);
                    urlResponse.Url = await ReturnOOSUrlStringAsync(urlsr, headers.UserName, headers.UserFullName, headers.WopiSrc, wopiAction);*/
                    urlResponse.Url = await GetFileUrlAsync(headers);
                    urlResponse.NewFileName = Path.GetFileName(filePath);
                }
            }
            catch (Exception ex)
            {
                if (System.IO.File.Exists($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt"))
                {
                    System.IO.File.AppendAllText($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt", $"ReplaceSignatureOnExistingFile: {DateTime.UtcNow}->{ex.Message} {Environment.NewLine}");
                }
                await ReturnExceptionAsync(ex, urlResponse);
            }
            return Ok(urlResponse);
        }
        #endregion
    }
}