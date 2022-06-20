using System;
using Autofac;
using WopiHost.Core;
using WopiHost.Core.Models;
using WopiHost.Abstractions;
using Microsoft.Extensions.Options;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Autofac.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.PlatformAbstractions;
using WopiHost.FileSystemProvider;
using System.Reflection;
using System.Runtime.Versioning;
using WopiHost.Utility.Model;
using WopiHost.Data;
using Microsoft.EntityFrameworkCore;

namespace WopiHost
{
    public class Startup
    {
        public IConfiguration Configuration { get; set; }

        public Startup(IWebHostEnvironment env, IConfiguration configuration)
        {
            var appEnv = PlatformServices.Default.Application;

            var builder = new ConfigurationBuilder()
                .SetBasePath(env.ContentRootPath)
                //.AddJsonFile("appsettings.json", true, true)
                .AddJsonFile($"appsettings.{env.EnvironmentName}.json", true, true)
                .AddInMemoryCollection(new Dictionary<string, string>
                {
                    { nameof(env.WebRootPath), env.WebRootPath },
                    { nameof(env.ContentRootPath), env.ContentRootPath },
                    { nameof(appEnv.ApplicationBasePath), appEnv.ApplicationBasePath } 
                })
                //.AddJsonFile($"config.{env.EnvironmentName}.json", true)
                .AddEnvironmentVariables();

            Configuration = builder.Build();
        }

        /// <summary>
        /// Sets up the DI container. Loads types dynamically (http://docs.autofac.org/en/latest/register/scanning.html)
        /// </summary>
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddControllers();//.AddControllersAsServices(); https://autofaccn.readthedocs.io/en/latest/integration/aspnetcore.html#controllers-as-services
            //cache Discoverer Xml
            services.AddMemoryCache();
            
            //Add HttpClientFactory Middleware
            services.AddHttpClient();

            // Ideally, pass a persistent dictionary implementation
            //services.AddSingleton<IDictionary<string, LockInfo>>(d => new Dictionary<string, LockInfo>());

            services.AddLogging(loggingBuilder =>
            {
                loggingBuilder.AddConsole();//Configuration.GetSection("Logging")
                loggingBuilder.AddDebug();
            });

            // Configuration
            services.AddOptions();

            services.Configure<WopiHostOptions>(Configuration);

            /*services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme).AddJwtBearer(options =>
            {
                options.TokenValidationParameters = new TokenValidationParameters
                {
                    ValidateIssuer = false,
                    ValidateAudience = false,
                    ValidateLifetime = true,
                    ValidateIssuerSigningKey = true,
                    ValidIssuer = Configuration["Jwt:Issuer"],
                    ValidAudience = Configuration["Jwt:Issuer"],
                    IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(Configuration["Jwt:Key"]))
                };
            });*/

            services.AddScoped<IWopiStorageProvider, WopiFileSystemProvider>();
            services.AddScoped<IWopiSecurityHandler, WopiSecurityHandler>();
            // Add WOPI (depends on file provider)
            services.AddWopi(GetSecurityHandler(services, Configuration.Get<WopiHostOptions>().WopiFileProviderAssemblyName));

            //Sqlite
            /*string wopisqlite =  $"DataSource={Configuration.Get<WopiHostOptions>().WebRootPath}\\LiteDb\\wopisqlite.db";
            services.AddEntityFrameworkSqlite().AddDbContext<WopiHostContext>(options => options.UseSqlite(wopisqlite));*/

            //UseSqlServer
            string sqlServer = $"{Configuration.GetSection("WopiHostContext").Value}";
            services.AddDbContext<WopiHostContext>(options => options.UseSqlServer(sqlServer));

            services.AddScoped<DbContext, WopiHostContext>();
            //RedisCache Azure
            services.AddStackExchangeRedisCache(options =>
            {
                options.Configuration = Configuration.Get<WopiHostOptions>().RedisConnectionString;
                options.InstanceName = "WOPI_";
                /*options.ConfigurationOptions = new StackExchange.Redis.ConfigurationOptions()
                {
                    ConnectRetry = 3
                };*/
            });
        }

        private IWopiSecurityHandler GetSecurityHandler(IServiceCollection services, string wopiFileProviderAssemblyName)
        {
            var providerBuilder = new ContainerBuilder();
            // Add file provider implementation
            providerBuilder.AddFileProvider(wopiFileProviderAssemblyName);
            providerBuilder.Populate(services);
            var providerContainer = providerBuilder.Build();
            return providerContainer.Resolve<IWopiSecurityHandler>();
        }


        /// <summary>
        /// Configure is called after ConfigureServices is called.
        /// </summary>
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            //loggerFactory.AddConsole(Configuration.GetSection("Logging"));
            //loggerFactory.AddDebug();

            if (env.EnvironmentName == "Development" || env.EnvironmentName == "Staging" || env.EnvironmentName == "StagingLocal")
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                //app.UseHttpsRedirection();
            }

            app.UseRouting();

            // Automatically authenticate
            app.UseAuthentication();

            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });

            app.Run(async context =>
            {
                await context.Response.WriteAsync(DefaultShow(env));
            });
        }

        protected string DefaultShow(IWebHostEnvironment env)
        {
            string ver = Assembly.GetEntryAssembly()?.GetCustomAttribute<TargetFrameworkAttribute>()?.FrameworkName;
            string html = "<!DOCTYPE html><html><head>";
            html = $"{html}<title>WOPI Host - Hoxro</title>";
            html = $"{html}<style>.loader{{border: 16px solid #f3f3f3;border-radius: 50%;border-top: 16px solid blue;";
            html = $"{html}border-bottom: 16px solid blue;width: 120px;height: 120px;-webkit-animation: spin 2s linear infinite;";
            html = $"{html}animation: spin 2s linear infinite;}}";
            html = $"{html}@-webkit-keyframes spin{{0% {{ -webkit-transform: rotate(0deg); }}100% {{ -webkit-transform: rotate(360deg); }}}}";
            html = $"{html}@keyframes spin {{0% {{ transform: rotate(0deg); }}100% {{ transform: rotate(360deg); }}}}";
            html = $"{html}</style></head>";
            html = $"{html}<body><div style='margin-top:10%'>";
            html = $"{html}<div class='loader' style ='margin-left:46%;'></div>";
            html = $"{html}<br/>";
            html = $"{html}<div style ='text-align: center;'><h1>Wopihost is running</h1>";
            html = $"{html}<br/>";
            html = $"{html}<div style ='text-align: center;'><h3>Environment: {env.EnvironmentName} </br> Version: {ver}</h3>";
            html = $"{html}</div></div></body></html>";
            return html;
        }
    }
}
