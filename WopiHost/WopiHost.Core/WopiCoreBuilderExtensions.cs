using System.Reflection;
using Microsoft.AspNetCore.Authorization;
using Microsoft.Extensions.DependencyInjection;
using Newtonsoft.Json.Serialization;
using WopiHost.Abstractions;
using WopiHost.Core.Security.Authentication;
using WopiHost.Core.Security.Authorization;
using WopiHost.FileSystemProvider;
using WopiHost.Core.LiteDb;
using WopiHost.Service.Document;
using WopiHost.Service.CasheService;
using WopiHost.Service;
using WopiHost.Service.UserService;

namespace WopiHost.Core
{
    public static class WopiCoreBuilderExtensions
    {
        public static bool IsUseLocalStorage { get; set; }
        public static void AddWopi(this IServiceCollection services, IWopiSecurityHandler securityHandler)
        {
            services.AddAuthorizationCore();

            // Add authorization handler
            services.AddSingleton<IAuthorizationHandler, WopiAuthorizationHandler>();
            services.AddScoped<IAzureBlobStorage, AzureBlobStorage>();
            services.AddScoped<IDocumentService, DocumentService>();
            //services.AddSingleton<ILiteDbFileStorageInfoManager, LiteDbFileStorageInfoManager>();
            //RedisCache Service
            services.AddScoped<IRedisCacheServiceBase, RedisCacheServiceBase>();
            services.AddScoped<IRedisCacheWopiLockService, RedisCacheWopiLockService>();

            //User
            services.AddScoped<IUserService, UserService>();


            services.AddControllers()
              .AddApplicationPart(typeof(WopiCoreBuilderExtensions).GetTypeInfo().Assembly) // Add controllers from this assembly
              .AddJsonOptions(o => o.JsonSerializerOptions.PropertyNamingPolicy = null); // Ensure PascalCase property name-style

            services.AddAuthentication(o => { o.DefaultScheme = AccessTokenDefaults.AuthenticationScheme; })
                .AddTokenAuthentication(AccessTokenDefaults.AuthenticationScheme, AccessTokenDefaults.AuthenticationScheme, options => { options.SecurityHandler = securityHandler; });
        }
    }
}
