
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using WopiHost.Abstractions;
using WopiHost.Core.Models;
using WopiHost.Core.Results;
using WopiHost.Core.Security;
using WopiHost.Core.Security.Authentication;
using WopiHost.Discovery;
using WopiHost.Discovery.Enumerations;
using WopiHost.FileSystemProvider;
using WopiHost.Service;
using WopiHost.Url;
using WopiHost.Utility.Common;
using WopiHost.Utility.ExtensionMetod;
using WopiHost.Utility.Model;
using WopiHost.Utility.ViewModel;

namespace WopiHost.Core.Controllers
{
    /// <summary>
    /// Implementation of WOPI server protocol https://msdn.microsoft.com/en-us/library/hh659001.aspx
    /// </summary>
    [Route("wopi/[controller]")]
    public class FilesController : WopiControllerBase
    {
        //readonly string ROOT_PATH = "";
        private readonly IAuthorizationService _authorizationService;
        public ICobaltProcessor CobaltProcessor { get; set; }
        //private readonly ILiteDbFileStorageInfoManager _liteDbManager;
        private readonly IAzureBlobStorage _azureBlobStorage;
        private readonly IMemoryCache _memoryCache;
        private readonly IHttpClientFactory _clientFactory;
        private readonly IRedisCacheWopiLockService _redisCacheWopiLockService;
        private WopiUrlBuilder _urlGenerator;
        private HostCapabilities HostCapabilities => new()
        {
            SupportsCobalt = CobaltProcessor != null,
            SupportsGetLock = true,
            SupportsLocks = true,
            SupportsExtendedLockLength = true,
            SupportsFolders = true,//?
            SupportsCoauth = true,//?
            SupportsUpdate = nameof(PutFile) != null, //&& PutRelativeFile
            SupportsContainers = false,
        };
        /// <summary>
        /// Collection holding information about locks. Should be persistent.
        /// </summary>
        //public static IDictionary<string, LockInfo> LockStorage;

        private string WopiOverrideHeader => HttpContext.Request.Headers[WopiHeaders.WopiOverride];

        private readonly IOptionsSnapshot<WopiHostOptions> _wopiHostOptions;
        public string WopiClientUrl;
        public WopiDiscoverer _wopiDiscoverer;
        //public WopiClaimTypesVm _wopiClaimTypesVm { get; set; }

        public WopiUrlBuilder UrlGenerator => _urlGenerator ??= new WopiUrlBuilder(
             _wopiDiscoverer, new WopiUrlSettings
             {
                 UI_LLCC = new CultureInfo("en-US"),
                 DC_LLCC = new CultureInfo("en-US"),
                 DISABLE_CHAT = 1,
                 BUSINESS_USER = 1,
                 VALIDATOR_TEST_CATEGORY = "All"
             });

        public FilesController(
            //ILiteDbFileStorageInfoManager liteDbManager
            IAzureBlobStorage azureBlobStorage
            , IWopiStorageProvider storageProvider
            , IWopiSecurityHandler securityHandler
            , IOptionsSnapshot<WopiHostOptions> wopiHostOptions
            , IAuthorizationService authorizationService
            //, IDictionary<string, LockInfo> lockStorage
            , IMemoryCache memoryCache
            , IHttpClientFactory clientFactory
            , IRedisCacheWopiLockService redisCacheWopiLockService
            , ICobaltProcessor cobaltProcessor = null)
            : base(storageProvider, securityHandler, wopiHostOptions)
        {
            _authorizationService = authorizationService;
            //LockStorage = lockStorage;
            CobaltProcessor = cobaltProcessor;
            //_liteDbManager = liteDbManager;
            _azureBlobStorage = azureBlobStorage;
            _wopiHostOptions = wopiHostOptions;
            _memoryCache = memoryCache;
            WopiClientUrl = _wopiHostOptions.Value.WopiO365ClientUrl;//WopiO365ClientUrl
            _clientFactory = clientFactory;
            _wopiDiscoverer = new WopiDiscoverer(new HttpDiscoveryFileProvider(WopiClientUrl, clientFactory), memoryCache: _memoryCache);
            _redisCacheWopiLockService = redisCacheWopiLockService;
        }

        /// <summary>
        /// Returns the metadata about a file specified by an identifier.
        /// Specification: https://msdn.microsoft.com/en-us/library/hh643136.aspx
        /// Example URL: HTTP://server/<...>/wopi*/files/<id>
        /// </summary>
        /// <param name="id">File identifier.</param>
        /// <returns></returns>
        [HttpGet("{id}")]
        public async Task<IActionResult> GetCheckFileInfo(string id)
        {
            try
            {
                //validate WopiProof
                bool isValidProof = //true;
                await ValidateWopiProof(HttpContext);
                if (!isValidProof)
                {
                    return StatusCode(StatusCodes.Status500InternalServerError, $"Request not came from office online server.");
                }
                if (!(await _authorizationService.AuthorizeAsync(User, new FileResource(id), WopiOperations.Read)).Succeeded)
                {
                    return Unauthorized();
                }
                WopiDocsTypeEnum wopiDocsType = (WopiDocsTypeEnum)(Enum.Parse(typeof(WopiDocsTypeEnum), User?.FindFirst(WopiClaimTypes.WopiDocsType)?.Value) ?? WopiDocsTypeEnum.WopiDocs);
                // Get file
                var file = StorageProvider.GetWopiFile(id, wopiDocsType);
                //get View/EmbeddedView/Edit URL
                CheckFileInfoVm fileInfoVm = await GetFileUrlAsync(file, User, wopiDocsType: wopiDocsType);
                return new JsonResult(await file?.GetCheckFileInfo(HttpContext, User, HostCapabilities, _wopiHostOptions, fileInfoVm));
            }
            catch (Exception ex)
            {
                if (System.IO.File.Exists($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt"))
                {
                    System.IO.File.AppendAllText($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt", $"GetCheckFileInfo: {DateTime.UtcNow}->{ex.Message} {Environment.NewLine}");
                }
                return Unauthorized();
            }
        }

        /// <summary>
        /// Returns contents of a file specified by an identifier.
        /// Specification: https://msdn.microsoft.com/en-us/library/hh657944.aspx
        /// Example URL: HTTP://server/<...>/wopi*/files/<id>/contents
        /// </summary>
        /// <param name="id">File identifier.</param>
        /// <returns></returns>
        [HttpGet("{id}/contents")]
        public async Task<IActionResult> GetFile(string id)
        {
            try
            {
                // Check permissions
                if (!(await _authorizationService.AuthorizeAsync(User, new FileResource(id), WopiOperations.Read)).Succeeded)
                {
                    return Unauthorized();
                }
                if (string.IsNullOrEmpty(AccessToken))
                {
                    return StatusCode(StatusCodes.Status401Unauthorized, "User Unauthorized");
                }

                WopiDocsTypeEnum wopiDocsType = (WopiDocsTypeEnum)(Enum.Parse(typeof(WopiDocsTypeEnum), User.FindFirst(WopiClaimTypes.WopiDocsType)?.Value) ?? WopiDocsTypeEnum.WopiDocs);
                // Get file
                var file = StorageProvider.GetWopiFile(id, wopiDocsType);
                if (file is null)
                {
                    return StatusCode(StatusCodes.Status404NotFound, "File Unknown");
                }

                // Check expected size
                int? maximumExpectedSize = HttpContext.Request.Headers[WopiHeaders.MaxExpectedSize].ToString().ToNullableInt();
                if (maximumExpectedSize != null && (await file.GetCheckFileInfo(HttpContext, User, HostCapabilities, _wopiHostOptions)).Size > maximumExpectedSize.Value)
                {
                    return new PreconditionFailedResult();
                }
                Response.Headers[WopiHeaders.WopiItemVersion] = file?.LastWriteTimeUtc.ToString("O") ?? DateTime.UtcNow.ToString("O");
                // Try to read content from a stream
                return new FileStreamResult(file.GetReadStream(), "application/octet-stream");
            }
            catch (Exception ex)
            {
                if (System.IO.File.Exists($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt"))
                {
                    System.IO.File.AppendAllText($"{WopiHostOptions.Value.WebRootPath}\\log\\ErrorLogs.txt", $"GetFile: {DateTime.UtcNow}->{ex.Message} {Environment.NewLine}");
                }
            }
            return new NotFoundResult();
        }

        /// <summary>
        /// Updates a file specified by an identifier. (Only for non-cobalt files.)
        /// Specification: https://msdn.microsoft.com/en-us/library/hh657364.aspx
        /// Example URL: HTTP://server/<...>/wopi*/files/<id>/contents
        /// </summary>
        /// <param name="id">File identifier.</param>
        /// <returns></returns>
        [HttpPut("{id}/contents")]
        [HttpPost("{id}/contents")]
        public async Task<IActionResult> PutFile(string id)
        {
            // Check permissions
            /*var authorizationResult = await _authorizationService.AuthorizeAsync(User, new FileResource(id), WopiOperations.Update);
            if (!authorizationResult.Succeeded)
            {
                return Unauthorized();
            }*/

            if (string.IsNullOrEmpty(AccessToken))
            {
                return StatusCode(StatusCodes.Status401Unauthorized, "User Unauthorized");
            }
            var file = StorageProvider.GetWopiFile(id);
            if (file is null)
            {
                return StatusCode(StatusCodes.Status404NotFound, "File Unknown");
            }

            // Acquire lock
            var lockResult = await ProcessLockAsync(id);

            if (lockResult is OkResult)
            {
                byte[] newContent = await HttpContext.Request.Body.ReadBytesAsync();
                using (var stream = file.GetWriteStream())
                {
                    await stream.WriteAsync(newContent.AsMemory(0, newContent.Length));
                }
                //var fileStorageInfo = UpdateInfo.contents.SingleOrDefault(x => x.Id == id);
                FileStorageInfo fileStorageInfo = null;//await _liteDbManager.Get(id);

                //var azureAddress = fileStorageInfo.IsOverride == true ? fileStorageInfo.Blob : fileStorageInfo.BlobNew;

                if (fileStorageInfo?.IsUpDateDone == true && fileStorageInfo != null)
                {
                    try
                    {
                        await _azureBlobStorage.UploadAsync(fileStorageInfo.Container, fileStorageInfo.BlobNew ?? fileStorageInfo.Blob, fileStorageInfo.LocalPath);
                        fileStorageInfo.IsUploaded = true;
                        //await _liteDbManager.Update(fileStorageInfo);
                    }
                    catch (Exception)
                    {
                        WriteToLogTxt("PutFile error while upload to Blob Container.");
                        //throw ex;
                    }
                }
                else if (fileStorageInfo?.IsCancel == true)
                {
                    //await _liteDbManager.Update(fileStorageInfo);
                }
                Response.Headers[WopiHeaders.WopiItemVersion] = file?.LastWriteTimeUtc.ToString("O") ?? DateTime.UtcNow.ToString("O");
                return await Task.FromResult(new OkResult());
            }
            Response.Headers[WopiHeaders.WopiItemVersion] = file?.LastWriteTimeUtc.ToString("O") ?? DateTime.UtcNow.ToString("O");
            return await Task.FromResult(lockResult);
        }


        /// <summary>
        /// Changes the contents of the file in accordance with [MS-FSSHTTP] and performs other operations like locking.
        /// MS-FSSHTTP Specification: https://msdn.microsoft.com/en-us/library/dd943623.aspx
        /// Specification: https://msdn.microsoft.com/en-us/library/hh659581.aspx
        /// Example URL: HTTP://server/<...>/wopi*/files/<id>
        /// </summary>
        /// <param name="id"></param>
        /*[HttpPost("{id}")]
        public async Task<IActionResult> PerformAction(string id)
        {
            // Check permissions
            if (!(await _authorizationService.AuthorizeAsync(User, new FileResource(id), WopiOperations.Update)).Succeeded)
            {
                return Unauthorized();
            }
            var file = StorageProvider.GetWopiFile(id);
            switch (WopiOverrideHeader)
            {
                case "COBALT":
                    var responseAction = CobaltProcessor.ProcessCobalt(file, User, await HttpContext.Request.Body.ReadBytesAsync());
                    HttpContext.Response.Headers.Add(WopiHeaders.CorrelationId, HttpContext.Request.Headers[WopiHeaders.CorrelationId]);
                    HttpContext.Response.Headers.Add("request-id", HttpContext.Request.Headers[WopiHeaders.CorrelationId]);
                    return new Results.FileResult(responseAction, "application/octet-stream");

                case "LOCK":
                case "UNLOCK":
                case "REFRESH_LOCK":
                case "GET_LOCK":
                    return ProcessLock2(id);

                case "PUT_RELATIVE":
                    //convert doc to docx
                    return new NotImplementedResult();

                default:
                    // Unsupported action
                    return new NotImplementedResult();
            }
        }*/

        [HttpPost("{id}"), WopiOverrideHeader(new[] { "PUT_RELATIVE" })]
        public async Task<IActionResult> PutRelativeFile(string id)
        {
            return Ok(await Task.FromResult($"{nameof(PutRelativeFile)} is not implemented yet."));
        }

        /// <summary>
        /// Changes the contents of the file in accordance with [MS-FSSHTTP].
        /// MS-FSSHTTP Specification: https://msdn.microsoft.com/en-us/library/dd943623.aspx
        /// Specification: https://msdn.microsoft.com/en-us/library/hh659581.aspx
        /// Example URL path: /wopi/files/(file_id)
        /// </summary>
        /// <param name="id">File identifier.</param>
        [HttpPost("{id}"), WopiOverrideHeader(new[] { "COBALT" })]
        public async Task<IActionResult> ProcessCobalt(string id)
        {
            // Check permissions
            if (!(await _authorizationService.AuthorizeAsync(User, new FileResource(id), WopiOperations.Update)).Succeeded)
            {
                return Unauthorized();
            }

            var file = StorageProvider.GetWopiFile(id);

            // TODO: remove workaround https://github.com/aspnet/Announcements/issues/342 (use FileBufferingWriteStream)
            var syncIoFeature = HttpContext.Features.Get<IHttpBodyControlFeature>();
            if (syncIoFeature is not null)
            {
                syncIoFeature.AllowSynchronousIO = true;
            }

            var responseAction = CobaltProcessor.ProcessCobalt(file, User, await HttpContext.Request.Body.ReadBytesAsync());
            HttpContext.Response.Headers.Add(WopiHeaders.CorrelationId, HttpContext.Request.Headers[WopiHeaders.CorrelationId]);
            HttpContext.Response.Headers.Add("request-id", HttpContext.Request.Headers[WopiHeaders.CorrelationId]);
            return new Results.FileResult(responseAction, "application/octet-stream");
        }

        /// <summary>
        /// Processes lock-related operations.
        /// Specification: https://wopi.readthedocs.io/projects/wopirest/en/latest/files/Lock.html
        /// Example URL path: /wopi/files/(file_id)
        /// </summary>
        /// <param name="id">File identifier.</param>
        [HttpPost("{id}"), WopiOverrideHeader(new[] { "LOCK", "UNLOCK", "REFRESH_LOCK", "GET_LOCK" })]
        public async Task<IActionResult> ProcessLockAsync(string id)
        {
            //var authorizationResult = await _authorizationService.AuthorizeAsync(User, new FileResource(id), WopiOperations.Read);
            // Check permissions
            if (string.IsNullOrEmpty(AccessToken))
            {
                return await Task.FromResult(StatusCode(StatusCodes.Status401Unauthorized, "User Unauthorized"));
            }
            var file = StorageProvider.GetWopiFile(id);
            if (file is null)
            {
                return await Task.FromResult(StatusCode(StatusCodes.Status404NotFound, "File Unknown"));
            }
            Response.Headers[WopiHeaders.WopiItemVersion] = file?.LastWriteTimeUtc.ToString("O") ?? DateTime.UtcNow.ToString("O");

            string oldLock = Request.Headers[WopiHeaders.OldLock];
            string newLock = Request.Headers[WopiHeaders.Lock];
            //byte[] newContent = await HttpContext.Request.Body.ReadBytesAsync();

            //lock (LockStorage)
            //{
            //var lockAcquired = await TryGetLock(id, out var existingLock);
            var existingLock = await GetLockFromRedisCache(id);
            var lockAcquired = existingLock != null;
            switch (WopiOverrideHeader)
            {
                case "GET_LOCK":
                    if (lockAcquired && !string.IsNullOrEmpty(existingLock?.Lock))
                    {
                        Response.Headers[WopiHeaders.Lock] = lockAcquired ? existingLock.Lock : " ";
                        //Response.Headers.Add(WopiHeaders.Lock, lockAcquired ? existingLock.Lock : string.Empty);
                    }
                    else
                    {
                        Response.Headers[WopiHeaders.Lock] = " ";
                    }
                    //HttpContext.Response.Headers[WopiHeaders.Lock] = existingLock?.Lock ?? string.Empty;
                    //WriteToLogTxt($"newLock:{newLock ?? ""} oldLock:{oldLock ?? ""} lockAcquired: {lockAcquired} existingLock: {existingLock} WopiHeaders.Lock: {Response.Headers[WopiHeaders.Lock]} LockStorageCount: {LockStorage?.Count}");
                    return new OkResult();

                case "LOCK":
                case "PUT":
                    if (oldLock is null)
                    {
                        // Lock / put
                        if (lockAcquired && !string.IsNullOrEmpty(existingLock?.Lock))// has lock
                        {
                            if (existingLock.Lock == newLock)
                            {
                                // File is currently locked and the lock ids match, refresh lock
                                existingLock.DateCreated = DateTime.UtcNow;
                                //update Redis lock
                                await this.UpdateLockFromRedisCache(id, existingLock);
                                return new OkResult();
                            }
                            else
                            {
                                //System.IO.File.AppendAllText($"{_wopiHostOptions.Value.WebRootPath}\\log\\LockMismatch.txt", $"{DateTime.UtcNow}(UTC) --> {User.Identity.Name} Override:{WopiOverrideHeader} Response:{Response.Headers[WopiHeaders.Lock]} [Lock / put] existingLock:{existingLock.Lock} newLock: {newLock} lockAcquired:{lockAcquired} LockStorageCount: LockStorage.Count {Environment.NewLine}");
                                // There is a valid existing lock on the file
                                return ReturnLockMismatch(Response, existingLock.Lock);
                            }
                        }
                        else
                        {
                            if (WopiOverrideHeader == "PUT")
                            {
                                // If the file is 0 bytes, this is document creation
                                if (file?.Length == 0)
                                {
                                    Response.Headers[WopiHeaders.WopiItemVersion] = file?.LastWriteTimeUtc.ToString("O") ?? DateTime.UtcNow.ToString("O");
                                    //System.IO.File.AppendAllText($"{_wopiHostOptions.Value.WebRootPath}\\log\\LockMismatch.txt", $"{DateTime.UtcNow}(UTC) File size 0--> Override:{WopiOverrideHeader} Response:{Response.Headers[WopiHeaders.Lock]} newContent {file?.Length} existingLock:{existingLock?.Lock} newLock: {newLock} lockAcquired:{lockAcquired} LockStorageCount: {LockStorage.Count} {Environment.NewLine}");
                                    return new OkResult();
                                }
                                else
                                {
                                    Response.Headers[WopiHeaders.WopiItemVersion] = file?.LastWriteTimeUtc.ToString("O") ?? DateTime.UtcNow.ToString("O");
                                    //System.IO.File.AppendAllText($"{_wopiHostOptions.Value.WebRootPath}\\log\\LockMismatch.txt", $"{DateTime.UtcNow}(UTC) File size can't be 0--> Override:{WopiOverrideHeader} Response:{Response.Headers[WopiHeaders.Lock]} newContent {file?.Length} existingLock:{existingLock?.Lock} newLock: {newLock} lockAcquired:{lockAcquired} LockStorageCount: {LockStorage.Count} {Environment.NewLine}");
                                    return ReturnLockMismatch(Response, existingLock?.Lock, "File size can't be 0");
                                }
                            }
                            // The file is not currently locked, create and store new lock information
                            //LockStorage[id] = new LockInfo { DateCreated = DateTime.UtcNow, Lock = newLock };

                            //****Redis Cache Set
                            LockInfo newLockInfo = new LockInfo { DateCreated = DateTime.UtcNow, Lock = newLock };
                            await this.SetLockOnRedisCache(id, newLockInfo);
                            return new OkResult();
                        }
                    }
                    else
                    {
                        // Unlock and relock (http://wopi.readthedocs.io/projects/wopirest/en/latest/files/UnlockAndRelock.html)
                        if (lockAcquired && !string.IsNullOrEmpty(existingLock?.Lock))
                        {
                            if (existingLock.Lock == oldLock)
                            {
                                // Replace the existing lock with the new one
                                //LockStorage[id] = new LockInfo { DateCreated = DateTime.UtcNow, Lock = newLock };
                                //****Redis Cache Set
                                LockInfo newLockInfo = new LockInfo { DateCreated = DateTime.UtcNow, Lock = newLock };
                                await this.SetLockOnRedisCache(id, newLockInfo);
                                return new OkResult();
                            }
                            else
                            {
                                //System.IO.File.AppendAllText($"{_wopiHostOptions.Value.WebRootPath}\\log\\LockMismatch.txt", $"{DateTime.UtcNow}(UTC) --> {User.Identity.Name} Override:{WopiOverrideHeader} Response:{Response.Headers[WopiHeaders.Lock]} [Unlock and relock] lockAcquired:{lockAcquired} LockStorageCount: LockStorage.Count {Environment.NewLine}");
                                // The existing lock doesn't match the requested one. Return a lock mismatch error along with the current lock
                                return ReturnLockMismatch(Response, existingLock.Lock, " Unlock and relock");
                            }
                        }
                        else
                        {
                            // The requested lock does not exist which should result in a lock mismatch error.
                            return ReturnLockMismatch(Response, reason: "File not locked");
                        }
                    }

                case "UNLOCK":
                    if (lockAcquired && !string.IsNullOrEmpty(existingLock?.Lock))
                    {
                        if (existingLock.Lock == newLock)
                        {
                            // Remove valid lock
                            //LockStorage.Remove(id);

                            //remove from redis cache
                            await this.DeleteLockFromRedisCache(id);
                            return new OkResult();
                        }
                        else
                        {
                            // The existing lock doesn't match the requested one. Return a lock mismatch error along with the current lock
                            return ReturnLockMismatch(Response, existingLock.Lock);
                        }
                    }
                    else
                    {
                        //WriteToLogTxt($"newLock:{newLock ?? ""} oldLock:{oldLock ?? ""} lockAcquired: {lockAcquired} existingLock: {existingLock} WopiHeaders.Lock: {Response.Headers[WopiHeaders.Lock]} LockStorageCount: {LockStorage.Count}");
                        // The requested lock does not exist.
                        return ReturnLockMismatch(Response, reason: $"File not locked [{WopiOverrideHeader}]");
                    }

                case "REFRESH_LOCK":
                    if (lockAcquired && !string.IsNullOrEmpty(existingLock?.Lock))
                    {
                        if (existingLock.Lock == newLock)
                        {
                            // Extend the lock timeout
                            existingLock.DateCreated = DateTime.UtcNow;
                            //update Redis lock
                            await this.UpdateLockFromRedisCache(id, existingLock);
                            return new OkResult();
                        }
                        else
                        {
                            // The existing lock doesn't match the requested one. Return a lock mismatch error along with the current lock
                            return ReturnLockMismatch(Response, existingLock.Lock);
                        }
                    }
                    else
                    {
                        // The requested lock does not exist.  That's also a lock mismatch error.
                        return ReturnLockMismatch(Response, reason: "File not locked");
                    }
            }
            //}

            return new OkResult();
        }
        #region "Locking"

        /*private IActionResult ProcessLock1(string id)
        {
            string oldLock = Request.Headers[WopiHeaders.OldLock];//For UnlockAndRelock
            string newLock = Request.Headers[WopiHeaders.Lock];


            //adding lock info
            var fileStorageInfo = _liteDbManager.Get(id).Result;
            if (fileStorageInfo != null && !string.IsNullOrEmpty(oldLock))
            {
                fileStorageInfo.OldLockId = oldLock;
            }

            // end lock info

            lock (LockStorage)
            {
                bool lockAcquired = TryGetLock(id, out var existingLock);
                switch (WopiOverrideHeader)
                {
                    case "GET_LOCK":
                        if (lockAcquired)
                        {
                            Response.Headers[WopiHeaders.Lock] = existingLock.Lock;
                            WriteToLogTxt($"newLock:{newLock ?? ""} oldLock:{oldLock ?? ""} Response:{Response.Headers[WopiHeaders.Lock]} Reason: GET_LOCK called successfull. FileId: {id}. LockStorageCount: {LockStorage.Count}");
                        }
                        else//Expect Lock but not there
                        {
                            string newGetLock = existingLock.Lock = !string.IsNullOrEmpty(fileStorageInfo?.LockId) && fileStorageInfo?.IsEditing == true ? fileStorageInfo?.LockId : Guid.NewGuid().ToString();
                            WriteToLogTxt($"newLock:{newLock ?? ""} oldLock:{oldLock ?? ""} Response:{newGetLock} Reason: GET_LOCK Create NEWGUID. FileId: {id}. LockStorageCount: {LockStorage.Count}");
                            LockStorage[id] = new LockInfo { DateCreated = DateTime.UtcNow, Lock = newGetLock };
                            Response.Headers[WopiHeaders.Lock] = newGetLock;
                        }
                        return new OkResult();

                    case "LOCK":
                    case "PUT":
                        if (oldLock == null)
                        {
                            // Lock / put
                            if (lockAcquired)
                            {
                                if (existingLock.Lock == newLock)
                                {
                                    // File is currently locked and the lock ids match, refresh lock
                                    existingLock.DateCreated = DateTime.UtcNow;
                                    WriteToLogTxt($"{id} File is currently locked and the lock ids match, refresh lock.[lockAcquired]");
                                    return new OkResult();
                                }
                                else
                                {
                                    // There is a valid existing lock on the file
                                    return ReturnLockMismatch(Response, existingLock.Lock, $"File has locked already.[lockAcquired]", newLock);
                                }
                            }
                            else//Expect Lock but not there
                            {
                                // The file is not currently locked, create and store new lock information
                                //if LiteDBInfo exist fileStorageInfo.LockId = newLock else newLock = newLock
                                newLock = !string.IsNullOrEmpty(fileStorageInfo?.LockId) && fileStorageInfo?.IsEditing == true ? fileStorageInfo?.LockId : newLock;
                                LockStorage[id] = new LockInfo { DateCreated = DateTime.UtcNow, Lock = newLock };
                                UpdateLiteDBInfo(fileStorageInfo, newLock);
                                System.IO.File.AppendAllText($"{_wopiHostOptions.Value.WebRootPath}\\log\\LockMismatch.txt", $"{DateTime.UtcNow}(UTC) --> {User.Identity.Name} Override:{WopiOverrideHeader} Response:{newLock} Reason: File locked succussfull. FileId: {id}. LockStorageCount: {LockStorage.Count} {Environment.NewLine}");
                                return new OkResult();
                            }
                        }
                        else
                        {
                            // Unlock and relock (http://wopi.readthedocs.io/projects/wopirest/en/latest/files/UnlockAndRelock.html)
                            if (lockAcquired)
                            {
                                if (existingLock.Lock == oldLock)
                                {
                                    // Replace the existing lock with the new one
                                    UpdateLiteDBInfo(fileStorageInfo, newLock);
                                    LockStorage[id] = new LockInfo { DateCreated = DateTime.UtcNow, Lock = newLock };
                                    WriteToLogTxt($"{id} Replace the existing lock with the new one. New lock {newLock}");
                                    return new OkResult();
                                }
                                else
                                {
                                    // The existing lock doesn't match the requested one. Return a lock mismatch error along with the current lock
                                    return ReturnLockMismatch(Response, existingLock.Lock, $"UnlockAndRelock- existing lock != current one[lockAcquired]", newLock);
                                }
                            }
                            else//Expect Lock but not there
                            {
                                // The requested lock does not exist which should result in a lock mismatch error.
                                //return ReturnLockMismatch(Response, reason: "UnlockAndRelock- File not locked", RequestLock: newLock);
                                //***********New Logic Start
                                existingLock.Lock = !string.IsNullOrEmpty(fileStorageInfo?.LockId) && fileStorageInfo?.IsEditing == true ? fileStorageInfo?.LockId : oldLock;
                                if (existingLock.Lock == oldLock)
                                {
                                    // Replace the existing lock with the new one
                                    UpdateLiteDBInfo(fileStorageInfo, newLock);
                                    LockStorage[id] = new LockInfo { DateCreated = DateTime.UtcNow, Lock = newLock };
                                    WriteToLogTxt($"{id} Replace the existing lock with the new one. New lock {newLock}");
                                    return new OkResult();
                                }
                                else
                                {
                                    LockStorage[id] = new LockInfo { DateCreated = DateTime.UtcNow, Lock = existingLock.Lock };
                                    // The existing lock doesn't match the requested one. Return a lock mismatch error along with the current lock
                                    return ReturnLockMismatch(Response, existingLock.Lock, $"UnlockAndRelock- existing lock != current one", newLock);
                                }
                                //***********New Logic End
                            }
                        }
                    case "CANCEL":
                    case "UNLOCK":
                        if (lockAcquired)
                        {
                            if (existingLock.Lock == newLock)
                            {
                                //update LiteDB
                                fileStorageInfo.IsEditing = false;
                                UpdateLiteDBInfo(fileStorageInfo);
                                // Remove valid lock
                                LockStorage.Remove(id);
                                WriteToLogTxt($"Remove {id} from LockStorage on UNLOCK call[lockAcquired]. existingLock= {existingLock.Lock} newLock = {newLock}");
                                return new OkResult();
                            }
                            else
                            {
                                //LockStorage.Remove(id);
                                // The existing lock doesn't match the requested one. Return a lock mismatch error along with the current lock
                                return ReturnLockMismatch(Response, existingLock.Lock, "existing lock doesn't match the requested one[lockAcquired]", RequestLock: newLock);
                            }
                        }
                        else//Expect Lock but not there
                        {
                            // The requested lock does not exist.
                            //return ReturnLockMismatch(Response, reason: "File not locked");
                            if (existingLock != null)
                            {
                                existingLock.Lock = !string.IsNullOrEmpty(fileStorageInfo?.LockId) && fileStorageInfo?.IsEditing == true ? fileStorageInfo?.LockId : Guid.NewGuid().ToString();
                                LockStorage[id] = new LockInfo { DateCreated = DateTime.UtcNow, Lock = existingLock.Lock };
                                if (existingLock.Lock == newLock)
                                {
                                    //update LiteDB
                                    fileStorageInfo.IsEditing = false;
                                    UpdateLiteDBInfo(fileStorageInfo);
                                    // Remove valid lock
                                    LockStorage.Remove(id);
                                    WriteToLogTxt($"Remove {id} from LockStorage on UNLOCK call.");
                                    return new OkResult();
                                }
                                else
                                {
                                    // The existing lock doesn't match the requested one. Return a lock mismatch error along with the current lock
                                    return ReturnLockMismatch(Response, existingLock.Lock, "existing lock doesn't match the requested one", RequestLock: newLock);
                                }
                            }
                            else
                            {
                                return new OkResult();
                            }

                        }

                    case "REFRESH_LOCK":
                        if (lockAcquired)
                        {
                            if (existingLock.Lock == newLock)
                            {
                                // Extend the lock timeout
                                existingLock.DateCreated = DateTime.UtcNow;
                                WriteToLogTxt($"Refresh lock successfully called[lockAcquired].");
                                return new OkResult();
                            }
                            else
                            {
                                // The existing lock doesn't match the requested one. Return a lock mismatch error along with the current lock
                                return ReturnLockMismatch(Response, existingLock.Lock, "existing lock doesn't match.[lockAcquired]", newLock);
                            }
                        }
                        else//Expect Lock but not there
                        {
                            // The requested lock does not exist.  That's also a lock mismatch error.
                            //return ReturnLockMismatch(Response, reason: "File not locked");
                            //********New Logic Start
                            existingLock.Lock = !string.IsNullOrEmpty(fileStorageInfo?.LockId) && fileStorageInfo?.IsEditing == true ? fileStorageInfo?.LockId : Guid.NewGuid().ToString();
                            LockStorage[id] = new LockInfo { DateCreated = DateTime.UtcNow, Lock = existingLock.Lock };
                            if (existingLock.Lock == newLock)
                            {
                                // Extend the lock timeout
                                existingLock.DateCreated = DateTime.UtcNow;
                                WriteToLogTxt($"Refresh lock successfully called[lockAcquired].");
                                return new OkResult();
                            }
                            else
                            {
                                // The existing lock doesn't match the requested one. Return a lock mismatch error along with the current lock
                                return ReturnLockMismatch(Response, existingLock.Lock, "existing lock doesn't match.", newLock);
                            }
                            //********New Logic End
                        }
                }
            }
            return new OkResult();
        }


        public IActionResult ProcessLock2(string id)
        {
            string oldLock = Request.Headers[WopiHeaders.OldLock];
            string newLock = Request.Headers[WopiHeaders.Lock];

            lock (LockStorage)
            {
                var lockAcquired = TryGetLock(id, out var existingLock);
                switch (WopiOverrideHeader)
                {
                    case "GET_LOCK":
                        if (lockAcquired)
                        {
                            Response.Headers[WopiHeaders.Lock] = existingLock.Lock ?? string.Empty;
                        }
                        return new OkResult();

                    case "LOCK":
                    case "PUT":
                        if (oldLock is null)
                        {
                            // Lock / put
                            if (lockAcquired)
                            {
                                if (existingLock.Lock == newLock)
                                {
                                    // File is currently locked and the lock ids match, refresh lock
                                    existingLock.DateCreated = DateTime.UtcNow;
                                    return new OkResult();
                                }
                                else
                                {
                                    // There is a valid existing lock on the file
                                    return ReturnLockMismatch(Response, existingLock.Lock);
                                }
                            }
                            else
                            {
                                // The file is not currently locked, create and store new lock information
                                LockStorage[id] = new LockInfo { DateCreated = DateTime.UtcNow, Lock = newLock };
                                return new OkResult();
                            }
                        }
                        else
                        {
                            // Unlock and relock (http://wopi.readthedocs.io/projects/wopirest/en/latest/files/UnlockAndRelock.html)
                            if (lockAcquired)
                            {
                                if (existingLock.Lock == oldLock)
                                {
                                    // Replace the existing lock with the new one
                                    LockStorage[id] = new LockInfo { DateCreated = DateTime.UtcNow, Lock = newLock };
                                    return new OkResult();
                                }
                                else
                                {
                                    // The existing lock doesn't match the requested one. Return a lock mismatch error along with the current lock
                                    return ReturnLockMismatch(Response, existingLock.Lock);
                                }
                            }
                            else
                            {
                                // The requested lock does not exist which should result in a lock mismatch error.
                                return ReturnLockMismatch(Response, reason: "File not locked");
                            }
                        }

                    case "UNLOCK":
                        if (lockAcquired)
                        {
                            if (existingLock.Lock == newLock)
                            {
                                // Remove valid lock
                                LockStorage.Remove(id);
                                return new OkResult();
                            }
                            else
                            {
                                // The existing lock doesn't match the requested one. Return a lock mismatch error along with the current lock
                                return ReturnLockMismatch(Response, existingLock.Lock);
                            }
                        }
                        else
                        {
                            // The requested lock does not exist.
                            return ReturnLockMismatch(Response, reason: $"File not locked [{WopiOverrideHeader}]");
                        }

                    case "REFRESH_LOCK":
                        if (lockAcquired)
                        {
                            if (existingLock.Lock == newLock)
                            {
                                // Extend the lock timeout
                                existingLock.DateCreated = DateTime.UtcNow;
                                return new OkResult();
                            }
                            else
                            {
                                // The existing lock doesn't match the requested one. Return a lock mismatch error along with the current lock
                                return ReturnLockMismatch(Response, existingLock.Lock);
                            }
                        }
                        else
                        {
                            // The requested lock does not exist.  That's also a lock mismatch error.
                            return ReturnLockMismatch(Response, reason: "File not locked");
                        }
                }
            }

            return new OkResult();
        }*/
        /*/// <summary>
        /// Get Lock
        /// </summary>
        /// <param name="fileId"></param>
        /// <param name="lockInfo"></param>
        /// <returns></returns>
        private bool TryGetLock(string fileId, out LockInfo lockInfo)
        {
            if (LockStorage.TryGetValue(fileId, out lockInfo))
            {
                if (lockInfo.Expired)
                {
                    WriteToLogTxt($"Remove {fileId} from LockStorage on TryGetLock.");
                    LockStorage.Remove(fileId);
                    return false;
                }
                return true;
            }
            return false;
        }*/

        /// <summary>
        /// Get Lock
        /// </summary>
        /// <param name="fileId"></param>
        /// <param name="lockInfo"></param>
        /// <returns></returns>
        private async Task<LockInfo> GetLockFromRedisCache(string fileId)
        {
            return await _redisCacheWopiLockService.GetWopiLockFromRedisCache(0, 0, 0, fileId);
        }
        /// <summary>
        /// Get Lock
        /// </summary>
        /// <param name="fileId"></param>
        /// <param name="lockInfo"></param>
        /// <returns></returns>

        private async Task<LockInfo> SetLockOnRedisCache(string fileId, LockInfo lockInfo)
        {
            return await _redisCacheWopiLockService.SetWopiLockFromRedisCache(0, 0, 0, fileId, lockInfo);
        }
        //Update
        private async Task<LockInfo> UpdateLockFromRedisCache(string fileId, LockInfo lockInfo)
        {
            return await _redisCacheWopiLockService.UpdateWopiLockFromRedisCache(0, 0, 0, fileId, lockInfo);
        }
        //Delete
        private async Task DeleteLockFromRedisCache(string fileId)
        {
            await _redisCacheWopiLockService.DeleteWopiLockFromRedisCache(0, 0, 0, fileId);
        }

        /*/// <summary>
        /// Update LiteDB Information. If param "fileStorageInfo" is null info will not update.
        /// </summary>
        /// <param name="fileStorageInfo"></param>
        /// <param name="newLock"></param>
        private void UpdateLiteDBInfo(FileStorageInfo fileStorageInfo, string newLock = null)
        {
            if (fileStorageInfo != null)
            {
                fileStorageInfo.LockId = newLock ?? fileStorageInfo.LockId;
                fileStorageInfo.DateCreated = DateTime.UtcNow;
                _liteDbManager.Update(fileStorageInfo);
            }
        }
        /// <summary>
        /// Remove LiteDB Information. If param "id" is null info will not remove.
        /// </summary>
        /// <param name="fileStorageInfo"></param>
        /// <param name="newLock"></param>
        private void RemoveLiteDBInfo(string id)
        {
            if (!string.IsNullOrEmpty(id))
            {
                _liteDbManager.Remove(id);
            }
        }*/
        private StatusCodeResult ReturnLockMismatch(HttpResponse response, string existingLock = null, string reason = null, string RequestLock = "")
        {
            response.Headers[WopiHeaders.Lock] = existingLock ?? " ";
            /*if(WopiOverrideHeader == "UNLOCK")
            {
                response.Headers[WopiHeaders.Lock] = string.Empty;
            }*/
            if (!string.IsNullOrEmpty(reason))
            {
                response.Headers[WopiHeaders.LockFailureReason] = $"{reason} [{response.Headers[WopiHeaders.Lock]}]";
            }
            //System.IO.File.AppendAllText($"{_wopiHostOptions.Value.WebRootPath}\\log\\LockMismatch.txt", $"{DateTime.UtcNow}(UTC) --> {User.Identity.Name} Override:{WopiOverrideHeader} Response:{response.Headers[WopiHeaders.Lock]} Reason:{reason ?? string.Empty} Request:{RequestLock ?? "null"} LockStorageCount: {LockStorage.Count} {Environment.NewLine}");
            //write the logs int text file
            /*if (System.IO.File.Exists($"{_wopiHostOptions.Value.WebRootPath}\\log\\LockMismatch.txt"))
            {
            }
            WopiVersionableTypeEnum wopiVersionableType = (WopiVersionableTypeEnum)(Enum.Parse(typeof(WopiVersionableTypeEnum), User.FindFirst(WopiClaimTypes.IsVersionable)?.Value) ?? WopiVersionableTypeEnum.True);
            if (wopiVersionableType == WopiVersionableTypeEnum.False)
            {
                //return new OkResult();
            }
            else
            {
                //return new OkResult();
            }*/
            return new Results.ConflictResult();
            //return new OkResult();
        }

        private void WriteToLogTxt(string message = "")
        {
            if (System.IO.File.Exists($"{_wopiHostOptions.Value.WebRootPath}\\log\\log.txt"))
            {
                System.IO.File.AppendAllText($"{_wopiHostOptions.Value.WebRootPath}\\log\\log.txt", $"{DateTime.UtcNow}(UTC)--> {User.Identity.Name} Override:{WopiOverrideHeader} {message} LockStorageCount: LockStorage?.Count  {Environment.NewLine}");
            }
        }
        private void WriteLockMismatch(HttpResponse response, string existingLock = null, string reason = null, string RequestLock = "")
        {
            if (System.IO.File.Exists($"{_wopiHostOptions.Value.WebRootPath}\\log\\LockMismatch.txt"))
            {
                System.IO.File.AppendAllText($"{_wopiHostOptions.Value.WebRootPath}\\log\\LockMismatch.txt", $"{DateTime.UtcNow}(UTC) --> {User.Identity.Name} Override:{WopiOverrideHeader} ExistingLock: {existingLock ?? "null"} Response:{response.Headers[WopiHeaders.Lock]} Reason:{reason ?? string.Empty} Request:{RequestLock ?? "null"} {Environment.NewLine}");
            }
        }

        #region Validate Proof Key

        /// <summary>
        /// Validates the WOPI Proof on an incoming WOPI request
        /// </summary>
        public async Task<bool> ValidateWopiProof(HttpContext context)
        {
            // Make sure the request has the correct headers
            if (string.IsNullOrEmpty(context.Request.Headers[WopiHeaders.Proof].ToString()) ||
                string.IsNullOrEmpty(context.Request.Headers[WopiHeaders.Time_Stamp].ToString()))
            {
                return await Task.FromResult(false);
            }

            // Set the requested proof values
            var requestProof = context.Request.Headers[WopiHeaders.Proof];
            var requestProofOld = string.Empty;
            if (context.Request.Headers[WopiHeaders.Proof_Old].ToString() != null)
            {
                requestProofOld = context.Request.Headers[WopiHeaders.Proof_Old];
            }

            // Get the WOPI proof info from discovery
            var discoProof = await GetWopiProof(context);
            string token = HttpContext.Request.Query[AccessTokenDefaults.AccessTokenQueryName].ToString();
            var accessTokenBytes = Encoding.UTF8.GetBytes(token);
            var hostUrl = context.Request.GetDisplayUrl();
            var timeStamp = Convert.ToInt64(context.Request.Headers[WopiHeaders.Time_Stamp]);
            var hostUrlBytes = Encoding.UTF8.GetBytes(hostUrl.ToUpperInvariant());
            var timeStampBytes = BitConverter.GetBytes(timeStamp).Reverse().ToArray();
            // Build expected proof
            List<byte> expected = new(
                4 + accessTokenBytes.Length +
                4 + hostUrlBytes.Length +
                4 + timeStampBytes.Length);

            // Add the values to the expected variable
            expected.AddRange(BitConverter.GetBytes(accessTokenBytes.Length).Reverse().ToArray());
            expected.AddRange(accessTokenBytes);
            expected.AddRange(BitConverter.GetBytes(hostUrlBytes.Length).Reverse().ToArray());
            expected.AddRange(hostUrlBytes);
            expected.AddRange(BitConverter.GetBytes(timeStampBytes.Length).Reverse().ToArray());
            expected.AddRange(timeStampBytes);
            byte[] expectedBytes = expected.ToArray();
            //UnixTimeToDateTime(timeStamp);
            var proofResult = VerifyProof(expectedBytes, requestProof, discoProof.value) ||
                VerifyProof(expectedBytes, requestProof, discoProof.oldvalue) ||
                VerifyProof(expectedBytes, requestProofOld, discoProof.value);

            #region Validate X-WOPI-TimeStamp
            var timeStampDatetime = new DateTime(timeStamp);
            bool expire = timeStampDatetime.AddMinutes(20) < DateTime.UtcNow;
            if (expire)
            {
                return false;
            }
            #endregion
            return await Task.FromResult(proofResult);
        }
        /// <summary>
        /// Verifies the proof against a specified key
        /// </summary>
        private static bool VerifyProof(byte[] expectedProof, string proofFromRequest, string proofFromDiscovery)
        {
            using (RSACryptoServiceProvider rsaProvider = new RSACryptoServiceProvider())
            {
                try
                {
                    rsaProvider.ImportCspBlob(Convert.FromBase64String(proofFromDiscovery));
                    bool isValidProof = rsaProvider.VerifyData(expectedProof, "SHA256", Convert.FromBase64String(proofFromRequest));
                    return isValidProof;
                }
                catch (FormatException)
                {
                    return false;
                }
                catch (CryptographicException)
                {
                    return false;
                }
            }
        }
        /// <summary>
        /// Gets the WOPI proof details from the WOPI discovery endpoint and caches it appropriately
        /// </summary>
        private async Task<WopiProof> GetWopiProof(HttpContext context)
        {
            var wopiProof = await _wopiDiscoverer.GetWopiProof();
            return wopiProof;
        }
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

        private async Task<CheckFileInfoVm> GetFileUrlAsync(IWopiFile file, ClaimsPrincipal user, WopiActionEnum wopiAction = WopiActionEnum.View, WopiDocsTypeEnum wopiDocsType = WopiDocsTypeEnum.WopiDocs, WopiVersionableTypeEnum isVersionable = WopiVersionableTypeEnum.True)
        {
            CheckFileInfoVm fileInfoVm = new();
            WopiClaimTypesVm claimTypesVm = await user.GetClaimInfoFromClaimIdentity();
            //get View
            string urlsr = await UrlGenerator.GetFileUrlAsync(file.Extension, claimTypesVm.WopiSrc, wopiAction);
            fileInfoVm.HostViewUrl = await ReturnOOSUrlStringAsync(urlsr, claimTypesVm.Email, claimTypesVm.UserFullName, claimTypesVm.WopiSrc, wopiAction, wopiDocsType: wopiDocsType, fileOriginalName: claimTypesVm?.FileOriginalName ?? string.Empty);
            //EmbeddedView
            wopiAction = WopiActionEnum.EmbedView;
            urlsr = await UrlGenerator.GetFileUrlAsync(file.Extension, claimTypesVm.WopiSrc, wopiAction);
            fileInfoVm.HostEmbeddedViewUrl = await ReturnOOSUrlStringAsync(urlsr, claimTypesVm.Email, claimTypesVm.UserFullName, claimTypesVm.WopiSrc, wopiAction, wopiDocsType: wopiDocsType, fileOriginalName: claimTypesVm?.FileOriginalName ?? string.Empty);
            //Edit
            wopiAction = WopiActionEnum.Edit;
            urlsr = await UrlGenerator.GetFileUrlAsync(file.Extension, claimTypesVm.WopiSrc, wopiAction);
            fileInfoVm.HostEditUrl = await ReturnOOSUrlStringAsync(urlsr, claimTypesVm.Email, claimTypesVm.UserFullName, claimTypesVm.WopiSrc, wopiAction, wopiDocsType: wopiDocsType, fileOriginalName: claimTypesVm?.FileOriginalName ?? string.Empty);
            return fileInfoVm;
        }
        #endregion

        #endregion
    }
}
