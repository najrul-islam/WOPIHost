namespace WopiHost.Abstractions
{
    /// <summary>
    /// Configuration class for WopiHost.Core
    /// </summary>
    public class WopiHostOptions
    {
        //wopi-docs
        public string WopiRootPath { get; set; }
        //other-docs
        public string WopiOtherDocsPath { get; set; }

        public string WopiFileProviderAssemblyName { get; set; }

        public string WebRootPath { get; set; }

        public string Region { get; set; }

        public string LocalStorage { get; set; }

        public bool IsUseLocalStorage { get; set; }

        //Azure Configure
        public string StorageAccountName { get; set; }

        public string StorageKey { get; set; }

        public string StorageAccountUrl { get; set; }

        public string StorageConnectionString { get; set; }

        public string WopiClientUrl { get; set; }

        public string WopiO365ClientUrl { get; set; }
        
        public string WopiHostUrl { get; set; }

        public string RedisConnectionString { get; set; }
        public string WopiHostContext { get; set; }
        public Jwt Jwt { get; set; }
    }

    public class Jwt
    {
        public string Key { get; set; }
        public string Issuer { get; set; }
        public string Audience { get; set; }
    }
}
