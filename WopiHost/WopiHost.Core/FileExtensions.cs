using System;
using System.IO;
using WopiHost.Core.Models;
using System.Security.Claims;
using WopiHost.Abstractions;
using Microsoft.AspNetCore.Http;
using System.Security.Cryptography;
using Microsoft.Extensions.Options;
using WopiHost.Core.Security.Authentication;
using Microsoft.AspNetCore.Authentication;
using System.Threading.Tasks;

namespace WopiHost.Core
{
    public static class FileExtensions
    {
        //private static readonly SHA256 SHA = SHA256.Create();

        public static async Task<CheckFileInfo> GetCheckFileInfo(this IWopiFile file, HttpContext context, ClaimsPrincipal principal, HostCapabilities capabilities, IOptionsSnapshot<WopiHostOptions> _wopiHostOptions = null, CheckFileInfoVm fileInfoVm = null)
        {
            CheckFileInfo CheckFileInfo = new();
            try
            {
                HttpRequest request = context.Request;
                var authenticateInfo = await context.AuthenticateAsync(AccessTokenDefaults.AuthenticationScheme);
                string token = authenticateInfo?.Properties?.GetTokenValue(AccessTokenDefaults.AccessTokenQueryName);

                if (principal != null)
                {
                    CheckFileInfo.UserId = principal.FindFirst(WopiClaimTypes.UserId)?.Value.ToSafeIdentity();
                    CheckFileInfo.UserFriendlyName = principal.FindFirst(WopiClaimTypes.UserFullName)?.Value;

                    CheckFileInfo.OwnerId = principal.FindFirst(WopiClaimTypes.UserId)?.Value.ToSafeIdentity();

                    WopiUserPermissions permissions = (WopiUserPermissions)Enum.Parse(typeof(WopiUserPermissions), principal.FindFirst(WopiClaimTypes.UserPermissions).Value);

                    CheckFileInfo.ReadOnly = permissions.HasFlag(WopiUserPermissions.ReadOnly);
                    CheckFileInfo.RestrictedWebViewOnly = permissions.HasFlag(WopiUserPermissions.RestrictedWebViewOnly);
                    CheckFileInfo.UserCanAttend = permissions.HasFlag(WopiUserPermissions.UserCanAttend);
                    CheckFileInfo.UserCanNotWriteRelative = permissions.HasFlag(WopiUserPermissions.UserCanNotWriteRelative);
                    CheckFileInfo.UserCanPresent = permissions.HasFlag(WopiUserPermissions.UserCanPresent);
                    CheckFileInfo.UserCanRename = permissions.HasFlag(WopiUserPermissions.UserCanRename);
                    CheckFileInfo.UserCanWrite = permissions.HasFlag(WopiUserPermissions.UserCanWrite);

                    //CheckFileInfo.WebEditingDisabled = permissions.HasFlag(WopiUserPermissions.WebEditingDisabled);
                }
                else
                {
                    CheckFileInfo.UserFriendlyName = "Anonymous";
                    CheckFileInfo.IsAnonymousUser = true;
                    CheckFileInfo.OwnerId = "Hoxro Limited";
                }

                //Required properties
                CheckFileInfo.BaseFileName = !string.IsNullOrEmpty(principal?.FindFirst(WopiClaimTypes.FileOriginalName)?.Value) ? principal?.FindFirst(WopiClaimTypes.FileOriginalName)?.Value : file.Name;
                //CheckFileInfo.BaseFileName = !string.IsNullOrEmpty(principal?.FindFirst(WopiClaimTypes.FileOriginalName)?.Value) ? Path.GetExtension(principal?.FindFirst(WopiClaimTypes.FileOriginalName)?.Value) : file.Name;
                CheckFileInfo.Size = file.Exists ? file.Length : 0;
                CheckFileInfo.Version = file.LastWriteTimeUtc.ToString("O");

                //Co-auth properties
                //CheckFileInfo.SupportsCoauth = capabilities.SupportsCoauth;
                CheckFileInfo.SupportsFolders = capabilities.SupportsFolders;
                CheckFileInfo.SupportsLocks = capabilities.SupportsLocks;

                // Set host capabilities
                //CheckFileInfo.SupportsScenarioLinks = capabilities.SupportsScenarioLinks;
                //CheckFileInfo.SupportsSecureStore = capabilities.SupportsSecureStore;
                //CheckFileInfo.SupportsGetFileWopiSrc = capabilities.SupportsGetFileWopiSrc;
                //CheckFileInfo.SupportsFileCreation = capabilities.SupportsFileCreation;

                //New OOS Start
                

                CheckFileInfo.LicenseCheckForEditIsEnabled = true;
                CheckFileInfo.SupportsGetLock = capabilities.SupportsGetLock;
                CheckFileInfo.SupportsExtendedLockLength = capabilities.SupportsExtendedLockLength;
                CheckFileInfo.SupportsEcosystem = capabilities.SupportsEcosystem;
                CheckFileInfo.SupportedShareUrlTypes = capabilities.SupportedShareUrlTypes;
                CheckFileInfo.SupportsUpdate = capabilities.SupportsUpdate;
                CheckFileInfo.SupportsCobalt = capabilities.SupportsCobalt;
                CheckFileInfo.SupportsRename = capabilities.SupportsRename;
                CheckFileInfo.SupportsDeleteFile = capabilities.SupportsDeleteFile;
                CheckFileInfo.SupportsUserInfo = capabilities.SupportsUserInfo;

                CheckFileInfo.BreadcrumbBrandName = "Hoxro Ltd";
                CheckFileInfo.BreadcrumbBrandUrl = "https://hoxro.com";
                CheckFileInfo.BreadcrumbDocName = CheckFileInfo.BaseFileName;
                CheckFileInfo.BreadcrumbFolderName = "Hoxro Folder";
                CheckFileInfo.BreadcrumbFolderUrl = "https://hoxro.com";
                CheckFileInfo.CloseUrl = "https://hoxro.com";
                CheckFileInfo.DownloadUrl = "https://hoxro.com";
                CheckFileInfo.FileExtension = Path.GetExtension(CheckFileInfo.BaseFileName);

                CheckFileInfo.FileUrl = string.Format($"{request.Scheme}://{request.Host}{request.Path}/contents?access_token={token}");

                CheckFileInfo.HostEditUrl = fileInfoVm?.HostEditUrl ?? "";
                CheckFileInfo.HostEmbeddedViewUrl = fileInfoVm?.HostEmbeddedViewUrl ?? "";
                CheckFileInfo.HostViewUrl = fileInfoVm?.HostViewUrl ?? "";
                CheckFileInfo.LastModifiedTime = "Last Modified Date";
                CheckFileInfo.SignInUrl = "https://ccsp.hoxro.com/Account/Login";
                CheckFileInfo.SignoutUrl = "https://ccsp.hoxro.com/Account/Logout";
                CheckFileInfo.UniqueContentId = file?.Name ?? "Hoxro ContentId";
                CheckFileInfo.UserInfo = CheckFileInfo?.UserFriendlyName ?? "Hoxro User Info";

                CheckFileInfo.FileVersionUrl = "https://hoxro.com";
                CheckFileInfo.FileSharingUrl = "https://ffc-onenote.officeapps.live.com";

                
                //New OOS End

                /*using (var stream = file.GetReadStream())
                {
                    byte[] checksum = SHA.ComputeHash(stream);
                    CheckFileInfo.Sha256 = Convert.ToBase64String(checksum);
                }*/

                CheckFileInfo.FileExtension = $".{file.Extension.TrimStart('.')}";
                CheckFileInfo.LastModifiedTime = file.LastWriteTimeUtc.ToString("o");
            }
            catch (Exception ex)
            {
                if (File.Exists($"{_wopiHostOptions.Value.WebRootPath}\\log\\log.txt"))
                {
                    File.AppendAllText($"{_wopiHostOptions.Value.WebRootPath}\\log\\log.txt", $"{DateTime.UtcNow}(UTC)--> {ex}  {Environment.NewLine}");
                }
                //throw ex;
                //Do nothing
            }
            return await Task.FromResult(CheckFileInfo);
        }
    }
}