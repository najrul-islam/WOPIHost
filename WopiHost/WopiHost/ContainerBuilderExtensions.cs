﻿using System;
using System.Runtime.Loader;
using Autofac;
using Microsoft.Extensions.PlatformAbstractions;
using WopiHost.Abstractions;

namespace WopiHost
{
    public static class ContainerBuilderExtensions
    {
        /*public static void AddFileProvider(this ContainerBuilder builder, WopiHostOptions options)
        {
            var providerAssembly = options.WopiFileProviderAssemblyName;
            // Load file provider
#if NET461
			var assembly = AppDomain.CurrentDomain.Load(new System.Reflection.AssemblyName(providerAssembly));
#endif

#if NETCOREAPP2_0
            // Load file provider
            var path = PlatformServices.Default.Application.ApplicationBasePath;
            var assembly = System.Runtime.Loader.AssemblyLoadContext.Default.LoadFromAssemblyPath(path + "\\" + providerAssembly + ".dll");
#endif
#if NETCOREAPP2_2
            // Load file provider
            var path = PlatformServices.Default.Application.ApplicationBasePath;
            var assembly = System.Runtime.Loader.AssemblyLoadContext.Default.LoadFromAssemblyPath(path + "\\" + providerAssembly + ".dll");
#endif
            builder.RegisterAssemblyTypes(assembly).AsImplementedInterfaces();
        }


        public static void AddCobalt(this ContainerBuilder builder)
        {
            // Load cobalt when running under the full .NET Framework
#if NET461
			var cobaltAssembly = AppDomain.CurrentDomain.Load(new System.Reflection.AssemblyName("WopiHost.Cobalt"));
			builder.RegisterAssemblyTypes(cobaltAssembly).AsImplementedInterfaces();
#endif
        }*/
        public static void AddFileProvider(this ContainerBuilder builder, string storageProviderAssemblyName)
        {
            // Load file provider
            //TODO: load by name? AssemblyLoadContext.Default.LoadFromAssemblyName
            //TODO: Unloadability https://docs.microsoft.com/en-us/dotnet/standard/assembly/unloadability
            var assembly = AssemblyLoadContext.Default.LoadFromAssemblyPath($"{AppContext.BaseDirectory}\\{storageProviderAssemblyName}.dll");
            builder.RegisterAssemblyTypes(assembly).AsImplementedInterfaces();
        }


        public static void AddCobalt(this ContainerBuilder builder)
        {
            // Load Cobalt            
            var cobaltAssembly = AssemblyLoadContext.Default.LoadFromAssemblyPath($"{AppContext.BaseDirectory}\\WopiHost.Cobalt.dll");
            builder.RegisterAssemblyTypes(cobaltAssembly).AsImplementedInterfaces();
        }
    }
}
