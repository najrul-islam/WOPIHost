using System;
using System.IO;
using System.Text;
using WopiHost.Abstractions;
using System.Threading.Tasks;
using WopiHost.Utility.Common;
using System.Collections.Generic;
using WopiHost.FileSystemProvider;
using Microsoft.Extensions.Options;
using WopiHost.Utility.ViewModel;
using WopiHost.Utility.XMLProcess;

namespace WopiHost.Service.Document
{
    public interface IDocumentService
    {
        /// <summary>
        /// Get Matter wise UnMailMerge Count
        /// </summary>
        /// <param name="headers"></param>
        /// <param name="wopiRequestMatterStage"></param>
        /// <returns></returns>
        Task<List<ContentDocument>> UnMailMergeList(RequestProcessHeader headers, WopiRequestUnMailMergeVM wopiRequestUnMailMergeVM, bool isMatterSpecific);
        /// <summary>
        /// Make Matter Stage Activity Document MailMerge
        /// </summary>
        /// <param name="headers"></param>
        /// <param name="matterStageActivity"></param>
        /// <returns></returns>
        Task<List<DocumentContent>> MatterStageDocumentMailMerge(RequestProcessHeader headers, MatterStageActivityMailMergeRequest matterStageActivity);
    }
    public class DocumentService : IDocumentService
    {
        public IOptionsSnapshot<WopiHostOptions> WopiHostOptions { get; }

        private readonly IAzureBlobStorage _azureBlobStorage;

        public DocumentService(IOptionsSnapshot<WopiHostOptions> wopiHostOptions,
            IAzureBlobStorage azureBlobStorage)
        {
            WopiHostOptions = wopiHostOptions;
            _azureBlobStorage = azureBlobStorage;
        }
        public async Task<List<ContentDocument>> UnMailMergeList(RequestProcessHeader headers, WopiRequestUnMailMergeVM wopiRequestUnMailMergeVM, bool isMatterSpecific)
        {
            List<ContentDocument> libraryContentResponse = new();
            if (wopiRequestUnMailMergeVM?.ListGlobalDataVm?.Count > 0)
            {
                foreach (var item in wopiRequestUnMailMergeVM?.ContentDocumnets)
                {
                    List<GlobalDataVm> unMergeList = new();
                    string documentPath = $"{WopiHostOptions.Value.WebRootPath}\\wopi-docs\\{Path.GetFileName(item.Bolb)}";
                    if (!string.IsNullOrEmpty(Path.GetFileName(item.Bolb)))
                    {
                        await DownloadFile(documentPath, headers.StorageName, item.Bolb);
                    }
                    if (File.Exists(documentPath))
                    {
                        ProcessMailMergeCount.DocumentMailMergeCount(documentPath, wopiRequestUnMailMergeVM?.ListGlobalDataVm, unMergeList, isMatterSpecific);
                    }
                    else
                    {
                        item.Message = "Blob Not Found.";
                        item.StatusCode = 404;
                    }
                    item.ListUnMergeGlobalDataVm = unMergeList;
                    libraryContentResponse.Add(item);
                }
            }
            return libraryContentResponse;
        }

        public async Task<List<DocumentContent>> MatterStageDocumentMailMerge(RequestProcessHeader headers, MatterStageActivityMailMergeRequest matterStageActivity)
        {
            if (matterStageActivity?.DocumentContents?.Count > 0)
            {
                string storageName = headers.StorageName;
                foreach (var item in matterStageActivity?.DocumentContents)
                {
                    string contentPath = $"{WopiHostOptions.Value.WebRootPath}\\wopi-docs\\{Path.GetFileName(item.ContentBlob)}";
                    string templatePath = $"{WopiHostOptions.Value.WebRootPath}\\wopi-docs\\{Path.GetFileName(item.TemplateBlob)}";
                    string newFilePath = $"{WopiHostOptions.Value.WebRootPath}\\wopi-docs\\{Path.GetFileName(item.Blob)}";
                    WopiUrlResponse res = await MergeContentWithTemplate(storageName, templatePath, contentPath, newFilePath, item.ContentBlob, item.TemplateBlob, matterStageActivity?.GDatas);
                    if (res.StatusCode == 200)
                    {
                        bool upload = await _azureBlobStorage.UploadExistingAsync(storageName, item.Blob, newFilePath);
                        if (!upload)
                        {
                            res.StatusCode = 404;
                            res.Message = "Error occured while upload to Blob";
                        }
                    }
                    item.StatusCode = res.StatusCode;
                    item.Message = res.Message;
                }
            }
            return await Task.FromResult(matterStageActivity?.DocumentContents);
        }


        public async Task<WopiUrlResponse> MergeContentWithTemplate(string storageName, string templatePath, string contentPath, string newFilePath, string contentBlob, string templateBlob, List<GlobalDataVm> gDatas)
        {
            var content = await DownloadFile(contentPath, storageName, contentBlob);
            WopiUrlResponse res = new() { StatusCode = 404, Message = "Download failed." };
            if (content)//if content downloded
            {
                var template = await DownloadFile(templatePath, storageName, templateBlob);
                if (template)//if template downloded
                {
                    res = await ProcessBulkXMLMailMerge.MergeDocuments(templatePath, contentPath, null, newFilePath, gDatas);
                }
            }
            return await Task.FromResult(res);
        }

        #region Download Document From Azure Blob Storage

        public async Task<bool> DownloadFile(string path, string container, string sourceUrl)
        {
            bool result;
            try
            {
                if (!File.Exists(path))
                {
                    await _azureBlobStorage.DownloadAsync(container, sourceUrl, path);
                    result = true;
                }
                else
                {
                    await _azureBlobStorage.UploadAsync(container, sourceUrl, path);
                    result = true;
                }
            }
            catch (Exception)
            {
                result = false;
            }
            return await Task.FromResult(result);
        }

        #endregion
    }
}