using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using WopiHost.Abstractions;
using WopiHost.Core.Models;

namespace WopiHost.Core.LiteDb
{
    /*public interface ILiteDbFileStorageInfoManager
    {
        Task<bool> InsertRequest(FileStorageInfo fileStorageInfo);
        Task<FileStorageInfo> Get(string id);
        Task<bool> Update(FileStorageInfo item);
        Task<bool> Remove(string id);
        Task<FileStorageInfo> CheckVersion(int matterDocumentId, int versionNo);
        Task<bool> CheckExists(string id);
    }
    public class LiteDbFileStorageInfoManager : ILiteDbFileStorageInfoManager
    {
        readonly LiteDB.LiteDatabase _context;
        LiteDB.LiteCollection<FileStorageInfo> Index { get; set; }
        private readonly IOptionsSnapshot<WopiHostOptions> _wopiHostOptions;

        public LiteDbFileStorageInfoManager(IOptionsSnapshot<WopiHostOptions> wopiHostOptions)
        {
            _wopiHostOptions = wopiHostOptions;
            _context = new LiteDB.LiteDatabase(_wopiHostOptions.Value.WebRootPath + "\\LiteDb\\request.db");
            Index = _context.GetCollection<FileStorageInfo>("FileStorageInfo");
        }
        public async Task<bool> InsertRequest(FileStorageInfo fileStorageInfo)
        {
            if(Index.Exists(x => x.Id == fileStorageInfo.Id))
            {
                Index.Delete(x => x.Id == fileStorageInfo.Id);
            }
            Index.Insert(fileStorageInfo);
            return await Task.FromResult(true);
        }
        public async Task<FileStorageInfo> Get(string id)
        {
            //var item = _index.FindOne(x => x.Id == id && x.IsUploaded == false);
            var item = Index.FindOne(x => x.Id == id);
            return await Task.FromResult(item);
        }
        public async Task<bool> Update(FileStorageInfo item)
        {
            Index.Update(item);
            bool output = true;
            return await Task.FromResult(output);
        }
        public async Task<bool> Remove(string id)
        {
            Index.Delete(x => x.Id == id);
            bool output = true;
            return await Task.FromResult(output);
        }
        public async Task<FileStorageInfo> CheckVersion(int parentId, int versionNo)
        {
            var item = Index.FindOne(x => x.ParentId == parentId && x.IsEditing == true);
            return await Task.FromResult(item);
        }
        public async Task<bool> CheckExists(string id)
        {
            return await Task.FromResult(Index.Exists(x => x.Id == id));
        }
    }*/
}
