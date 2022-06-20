using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using WopiHost.Abstractions;
using WopiHost.Data.Models;
using WopiHost.Data.ViewModel;
using WopiHost.FileSystemProvider;
using WopiHost.Service.UserService;

namespace WopiHost
{
    [ApiController]
    [Authorize]
    [Route("wopi/[controller]")]
    public class AuthorizationController : ControllerBase
    {
        private const string StorageName = "hxr-wopi";
        private readonly IWopiSecurityHandler _wopiSecurityHandler;
        private readonly IUserService _userService;
        private readonly IOptionsSnapshot<WopiHostOptions> _optionsSnapshot;
        private readonly IAzureBlobStorage _azureBlobStorage;

        public AuthorizationController(
            IWopiSecurityHandler wopiSecurityHandler,
            IUserService userService,
            IOptionsSnapshot<WopiHostOptions> optionsSnapshot,
            IAzureBlobStorage azureBlobStorage)
        {
            _wopiSecurityHandler = wopiSecurityHandler;
            _userService = userService;
            _optionsSnapshot = optionsSnapshot;
            _azureBlobStorage = azureBlobStorage;
        }
        // GET: wopi/Authorization/Token
        [AllowAnonymous]
        [HttpGet("Token")]
        public async Task<IActionResult> Token([FromBody] TokenInfo tokenInfo)
        {
            TokenResult tokenResult = new();
            try
            {
                var user = await _userService.GetLoginInfo(tokenInfo.UserName, tokenInfo.Password);
                if (user is not null)
                {
                    var userResult = await user.FirstOrDefaultAsync();
                    if (userResult?.UserId > 0)
                    {
                        var token = _wopiSecurityHandler.GenerateAccessToken(userResult.UserId, userResult.UserName, userResult.UserFullName);
                        tokenResult = new TokenResult
                        {
                            Access_token = _wopiSecurityHandler.WriteToken(token),
                            StatusCode = 200,
                            Message = "Success"
                        };
                        return Ok(tokenResult);
                    }
                }
            }
            catch (Exception)
            {

            }
            return Ok(tokenResult);
        }

        // POST: wopi/Authorization/AddUser
        [AllowAnonymous]
        [HttpPost("AddUser")]
        public async Task<IActionResult> AddUser([FromBody] User user)
        {
            try
            {
                await _userService.AddUser(user);
            }
            catch (Exception)
            {

            }
            return Ok(user);
        }

        // GET: wopi/Authorization/GetDocumentList
        [AllowAnonymous]
        [HttpGet("GetDocumentList")]
        public async Task<IActionResult> GetDocumentList()
        {
            IEnumerable<Document> documents = new List<Document>();
            try
            {
                documents = await _userService.GetDocumentList();
            }
            catch (Exception)
            {

            }
            return Ok(documents);
        }
        // GET: wopi/Authorization/GetDocumentList
        [AllowAnonymous]
        [HttpPost("AddDocument")]
        public async Task<IActionResult> AddDocument(Document document)
        {
            try
            {
                document = await _userService.AddDocument(document);
            }
            catch (Exception)
            {

            }
            return Ok(document);
        }

        // GET: wopi/Authorization/Download
        [AllowAnonymous]
        [HttpGet("Download")]
        public async Task<IActionResult> Download(Document document)
        {
            try
            {
                string path = $"{_optionsSnapshot.Value.WebRootPath}\\{_optionsSnapshot.Value.WopiRootPath}\\{document.Blob}";
                var bytes = System.Text.Encoding.UTF8.GetBytes($"{Path.GetFileName(document.Blob)}");
                var fileIdentifier = Convert.ToBase64String(bytes);
                if (!System.IO.File.Exists(path))
                {
                    await _azureBlobStorage.DownloadAsync(StorageName, document.Blob, path);
                }
            }
            catch (Exception)
            {

            }
            return Ok(document);
        }
    }
}
