using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using WopiHost.Abstractions;
using WopiHost.Core.Models;

namespace WopiHost.Core.LiteDb
{
    public interface ILockStorageDbManager
    {
        Task<bool> InsertRequest(FileStorageInfo fileStorageInfo);
        Task<FileStorageInfo> Get(string id);
        Task<bool> Update(FileStorageInfo item);
    }
    public class LockStorageDbManager : ILockStorageDbManager
    {

        LiteDB.LiteDatabase _context;
        LiteDB.LiteCollection<FileStorageInfo> _index { get; set; }
        private readonly IOptionsSnapshot<WopiHostOptions> _wopiHostOptions;

        public LockStorageDbManager(IOptionsSnapshot<WopiHostOptions> wopiHostOptions)
        {

            this._wopiHostOptions = wopiHostOptions;
            _context = new LiteDB.LiteDatabase(this._wopiHostOptions.Value.WebRootPath + "\\LiteDb\\request.db");
            _index = _context.GetCollection<FileStorageInfo>("LockStorage");
        }
        public async Task<bool> InsertRequest(FileStorageInfo fileStorageInfo)
        {
            _index.Delete(x => x.IsUploaded == true);
            _index.Insert(fileStorageInfo);
            return await Task.FromResult(true);
        }
        public async Task<FileStorageInfo> Get(string id)
        {
            var item = _index.FindOne(x => x.Id == id && x.IsUploaded == false);
            return await Task.FromResult(item);
        }
        public async Task<bool> Update(FileStorageInfo item)
        {
            var output = false;
            _index.Update(item);
            output = true;
            return await Task.FromResult(output);
        }

    }

}
