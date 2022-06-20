using System;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using WopiHost.Utility.Enum;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using Newtonsoft.Json.Serialization;
using Microsoft.Extensions.Caching.Distributed;

namespace WopiHost.Service.CasheService
{
    public interface IRedisCacheServiceBase
    {
        Task<string> GetStringAsync(string key);
        Task SetStringAsync(string key, object response, int minute);
        Task SetStringAsync(string key, object response, double hours);

        Task UpdateStringAsync(string key, object updateObject, int minute);
        Task UpdateStringAsync(string key, object updateObject, double hours);

        Task DeleteStringAsync(string key);

        Task<string> GetKeyAsync(EnumRedisCacheModulesKeyName redisCacheModules);
        Task<string> GetKeyAsync(EnumRedisCacheModulesKeyName redisCacheModules, string token);
    }
    public class RedisCacheServiceBase : IRedisCacheServiceBase
    {
        //private readonly IHttpContextAccessor _contextAccessor;
        private readonly IDistributedCache _redisCache;
        public RedisCacheServiceBase(
            //IHttpContextAccessor contextAccessor,
            IDistributedCache redisCache)
        {
            //_contextAccessor = contextAccessor;
            _redisCache = redisCache;
        }
        //get value by key
        public async Task<string> GetStringAsync(string key)
        {
            return await _redisCache.GetStringAsync(key);
        }
        //create in minute
        public async Task SetStringAsync(string key, object response, int minute)
        {
            var options = new DistributedCacheEntryOptions();
            options.SetSlidingExpiration(TimeSpan.FromMinutes(minute));
            await _redisCache.SetStringAsync(key, JsonConvert.SerializeObject(response, new JsonSerializerSettings { ContractResolver = new CamelCasePropertyNamesContractResolver() }), options);
        }
        //in hours
        public async Task SetStringAsync(string key, object response, double hours)
        {
            var options = new DistributedCacheEntryOptions();
            options.SetSlidingExpiration(TimeSpan.FromHours(hours));
            await _redisCache.SetStringAsync(key, JsonConvert.SerializeObject(response, new JsonSerializerSettings { ContractResolver = new CamelCasePropertyNamesContractResolver() }), options);
        }

        //update in minute
        public async Task UpdateStringAsync(string key, object updateObject, int minute)
        {
            var options = new DistributedCacheEntryOptions();
            options.SetSlidingExpiration(TimeSpan.FromMinutes(minute));
            await _redisCache.SetStringAsync(key, JsonConvert.SerializeObject(updateObject, new JsonSerializerSettings { ContractResolver = new CamelCasePropertyNamesContractResolver() }), options);
        }
        //in hours
        public async Task UpdateStringAsync(string key, object updateObject, double hours)
        {
            var options = new DistributedCacheEntryOptions();
            options.SetSlidingExpiration(TimeSpan.FromHours(hours));
            await _redisCache.SetStringAsync(key, JsonConvert.SerializeObject(updateObject, new JsonSerializerSettings { ContractResolver = new CamelCasePropertyNamesContractResolver() }), options);
        }

        //delete
        public async Task DeleteStringAsync(string key)
        {
            await _redisCache.RemoveAsync(key);
        }


        public async Task<string> GetKeyAsync(EnumRedisCacheModulesKeyName redisCacheModules)
        {
            string companyGuid = "";//await _contextAccessor.HttpContext.User.GetCompanyGuidFromClaimIdentity();
            if (!string.IsNullOrEmpty(companyGuid))
            {
                return await Task.FromResult($"{redisCacheModules}");
            }
            else
            {
                throw new Exception("CompanyGuid can not be null.");
            }
        }
        public async Task<string> GetKeyAsync(EnumRedisCacheModulesKeyName redisCacheModules, string token)
        {
            //string companyGuid = "";//await _contextAccessor.HttpContext.User.GetCompanyGuidFromClaimIdentity();
            if (!string.IsNullOrEmpty(token))
            {
                return await Task.FromResult($"{redisCacheModules}_{token}");
            }
            else
            {
                throw new Exception("CompanyGuid can not be null.");
            }
        }
    }
}
