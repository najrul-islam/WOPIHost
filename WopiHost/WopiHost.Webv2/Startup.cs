using System;
using System.Text;
using WopiHost.Discovery;
using WopiHost.Web.Models;
using WopiHost.Abstractions;
using WopiHost.Webv2.Utility;
using Microsoft.Identity.Web;
using Microsoft.AspNetCore.Http;
using Microsoft.Identity.Web.UI;
using WopiHost.FileSystemProvider;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Hosting;
using System.Collections.Generic;
using Microsoft.IdentityModel.Tokens;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Authorization;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Mvc.Authorization;
using Microsoft.Extensions.DependencyInjection;
//using Microsoft.Extensions.PlatformAbstractions;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using System.Net;
using System.Threading.Tasks;

namespace WopiHost.Web
{
    public class Startup
    {
        public IConfiguration Configuration { get; }

        public Startup(IWebHostEnvironment env)
        {
            //var appEnv = PlatformServices.Default.Application;
            var builder = new ConfigurationBuilder()
                .SetBasePath(env.ContentRootPath)
                .AddJsonFile("appsettings.json", true, true)
                .AddJsonFile($"appsettings.{env.EnvironmentName}.json", true, true)
                .AddInMemoryCollection(new Dictionary<string, string>
                {
                    { nameof(env.WebRootPath), env.WebRootPath },
                    { nameof(env.ContentRootPath), env.ContentRootPath },
                    //{ nameof(appEnv.ApplicationBasePath), appEnv.ApplicationBasePath } 
                })
                //.AddJsonFile($"config.{env.EnvironmentName}.json", true)
                .AddEnvironmentVariables();
            Configuration = builder.Build();
        }

        /// <summary>
        /// Sets up the DI container.
        /// </summary>
        public void ConfigureServices(IServiceCollection services)
        {
            //services.AddAuthentication(OpenIdConnectDefaults.AuthenticationScheme)
            //    .AddMicrosoftIdentityWebApp(Configuration.GetSection("AzureAd"));

            services.AddControllersWithViews(options =>
            {
                var policy = new AuthorizationPolicyBuilder()
                    .RequireAuthenticatedUser()
                    .Build();
                options.Filters.Add(new AuthorizeFilter(policy));
            });
            //services.AddRazorPages()
            //     .AddMicrosoftIdentityUI();
            // Configuration

            services.AddOptions();
            services.Configure<WopiHostOptions>(Configuration);
            //services.Configure<WopiOptions>(Configuration.GetSection("Wopi"));

            services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme).AddJwtBearer(options =>
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
            });

            services.AddSession();

            services.AddSingleton(Configuration);

            //Add HttpClientFactory Middleware
            services.AddHttpClient();

            /*services.AddHttpClient<IDiscoveryFileProvider, HttpDiscoveryFileProvider>(client =>
            {
                client.BaseAddress = new Uri(Configuration[$"Wopi:ClientUrl"]);
            });*/
            //services.AddSingleton<IDiscoverer, WopiDiscoverer>();

            services.AddScoped<IWopiStorageProvider, WopiFileSystemProvider>();
            services.AddScoped<IHttpClientHelperBase, HttpClientHelperBase>();
            services.AddScoped<IAzureBlobStorage, AzureBlobStorage>();

            services.AddLogging(loggingBuilder =>
            {
                loggingBuilder.AddConsole();//Configuration.GetSection("Logging")
                loggingBuilder.AddDebug();
            });
        }

        /// <summary>
        /// Configure is called after ConfigureServices is called.
        /// </summary>
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                //app.UseExceptionHandler("/Home/Error");
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }
            //app.UseHttpsRedirection();
            app.UseSession();

            //Add Authorization header
            app.Use(async (context, next) =>
            {
                var token = context.Session.GetString("Token");
                if (!string.IsNullOrEmpty(token))
                {
                    context.Request.Headers.Add("Authorization", "Bearer " + token);
                }
                await next();
            });
            //Unauthorized redirect to /Account/Login
            app.UseStatusCodePages(async context =>
            {
                var response = context.HttpContext.Response;
                if (response.StatusCode == (int)HttpStatusCode.Unauthorized || response.StatusCode == (int)HttpStatusCode.Forbidden)
                {
                    response.Redirect("/Account/Login");
                }
            });
            app.UseStaticFiles();

            app.UseRouting();

            app.UseAuthentication();
            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllerRoute(
                    name: "default",
                    pattern: "{controller=Home}/{action=Index}/{id?}");
                //endpoints.MapRazorPages();
            });
        }
    }
}