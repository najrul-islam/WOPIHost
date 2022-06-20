using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WopiHost.Service.CasheService;
using WopiHost.Utility.Enum;
using WopiHost.Utility.Model;

namespace WopiHost.Service
{
    public interface IRedisCacheWopiLockService
    {
        //get
        Task<LockInfo> GetWopiLockFromRedisCache(int companyId, int branchId, int userId, string fileId);
        //set
        Task<LockInfo> SetWopiLockFromRedisCache(int companyId, int branchId, int userId, string fileId, LockInfo lockInfo);
        //update
        Task<LockInfo> UpdateWopiLockFromRedisCache(int companyId, int branchId, int userId, string fileId, LockInfo updateLockInfo);
        //delete
        Task DeleteWopiLockFromRedisCache(int companyId, int branchId, int userId, string fileId);
    }
    public class RedisCacheWopiLockService : IRedisCacheWopiLockService
    {
        private readonly IRedisCacheServiceBase _redisCacheServiceBase;
        private readonly EnumRedisCacheModulesKeyName modulesKeyName = EnumRedisCacheModulesKeyName.WOPILock;
        public RedisCacheWopiLockService(IRedisCacheServiceBase redisCacheServiceBase)
        {
            _redisCacheServiceBase = redisCacheServiceBase;
        }

        public async Task<LockInfo> GetWopiLockFromRedisCache(int companyId, int branchId, int userId, string fileId)
        {
            LockInfo response = null;
            //make key
            string key = await _redisCacheServiceBase.GetKeyAsync(modulesKeyName, fileId);
            if (!string.IsNullOrEmpty(key))
            {
                string cacheResponse = await _redisCacheServiceBase.GetStringAsync(key);
                if (cacheResponse != null)
                {
                    response = JsonConvert.DeserializeObject<LockInfo>(cacheResponse);
                }
            }
            return await Task.FromResult(response);
        }

        public async Task<LockInfo> SetWopiLockFromRedisCache(int companyId, int branchId, int userId, string fileId, LockInfo lockInfo)
        {
            //make key
            string key = await _redisCacheServiceBase.GetKeyAsync(modulesKeyName, fileId);
            if (!string.IsNullOrEmpty(key))
            {
                await _redisCacheServiceBase.SetStringAsync(key, lockInfo, 2.0);
            }
            return await Task.FromResult(lockInfo);
        }

        public async Task<LockInfo> UpdateWopiLockFromRedisCache(int companyId, int branchId, int userId, string fileId, LockInfo updateLockInfo)
        {
            //make key
            string key = await _redisCacheServiceBase.GetKeyAsync(modulesKeyName, fileId);
            if (!string.IsNullOrEmpty(key))
            {
                await _redisCacheServiceBase.UpdateStringAsync(key, updateLockInfo, 2.0);
            }
            return await Task.FromResult(updateLockInfo);
        }

        //delete from cache
        public async Task DeleteWopiLockFromRedisCache(int companyId, int branchId, int userId, string fileId)
        {
            //make key
            string key = await _redisCacheServiceBase.GetKeyAsync(modulesKeyName, fileId);
            if (!string.IsNullOrEmpty(key))
            {
                await _redisCacheServiceBase.DeleteStringAsync(key);
            }
        }

    }
}
