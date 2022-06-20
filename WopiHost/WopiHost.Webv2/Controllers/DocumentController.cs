using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using Microsoft.AspNetCore.Authorization;
using Microsoft.Extensions.Options;
using WopiHost.Webv2.Utility;
using WopiHost.Abstractions;
using WopiHost.Data.Models;
using System.Net.Http;
using Microsoft.AspNetCore.Http;
using WopiHost.Data.ViewModel;
using WopiHost.FileSystemProvider;
using System.IO;

namespace WopiHost.Webv2.Controllers
{
    [Authorize]
    public class DocumentController : Controller
    {
        private readonly IHttpClientHelperBase _httpClientHelperBase;
        private readonly IOptionsSnapshot<WopiHostOptions> _optionsSnapshot;
        private readonly IAzureBlobStorage _azureBlobStorage;
        private const string StorageName = "hxr-wopi";

        public DocumentController(IHttpClientHelperBase httpClientHelperBase,
            IOptionsSnapshot<WopiHostOptions> optionsSnapshot,
            IAzureBlobStorage azureBlobStorage)
        {
            _httpClientHelperBase = httpClientHelperBase;
            _optionsSnapshot = optionsSnapshot;
            _azureBlobStorage = azureBlobStorage;
        }

        [HttpGet]
        public IActionResult AddDocument()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> AddDocument(FileUpload fileUpload)
        {
            try
            {
                IFormCollection formAsync = await Request.ReadFormAsync();
                foreach (var file in formAsync?.Files)
                {
                    if (file?.Length > 0)
                    {
                        string fileName = Path.GetFileNameWithoutExtension(file.FileName);
                        fileName = fileName.Replace("'", "");
                        string extension = Path.GetExtension(file.FileName)?.ToLowerInvariant();
                        Guid guid = Guid.NewGuid();
                        Document document = new()
                        {
                            DocumentGuid = guid,
                            DocumentName = fileName,
                            //FileName = guid.ToString() + extension,
                            Blob = $"{guid}{extension}",
                            CreatedDate = DateTime.UtcNow
                        };
                        MemoryStream stream = new();
                        await file.CopyToAsync(stream);
                        //upload to blob
                        await _azureBlobStorage.UploadAsync(StorageName, document.Blob, stream);
                        string endPoint = $"{_optionsSnapshot.Value.WopiHostUrl}/wopi/Authorization/AddDocument";
                        var saveDocument = await _httpClientHelperBase.RequestDataAsync<Document, Document>(endPoint, HttpMethod.Post, null, document);
                    }
                }
                return RedirectToAction("Index", "Home");
            }
            catch (Exception)
            {

            }
            return View();
        }
    }
}
