using Microsoft.Extensions.Caching.Memory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using WopiHost.Discovery.Enumerations;

namespace WopiHost.Discovery
{
    public class WopiDiscoverer : IDiscoverer
    {
        private const string ELEMENT_NET_ZONE = "net-zone";
        private const string ELEMENT_APP = "app";
        private const string ELEMENT_ACTION = "action";
        private const string ATTR_NET_ZONE_NAME = "name";
        private const string ATTR_ACTION_EXTENSION = "ext";
        private const string ATTR_ACTION_NAME = "name";
        private const string ATTR_ACTION_URL = "urlsrc";
        private const string ATTR_ACTION_REQUIRES = "requires";
        private const string ATTR_APP_NAME = "name";
        private const string ATTR_APP_FAVICON = "favIconUrl";
        private const string ATTR_VAL_COBALT = "cobalt";
        private const string PROOF_KEY = "proof-key";
        private const string PROOF_KEY_CACHE = "WopiProofKeyCache";

        private IDiscoveryFileProvider DiscoveryFileProvider { get; }

        public NetZoneEnum NetZone { get; }
        private readonly IMemoryCache _memoryCache;

        public WopiDiscoverer(IDiscoveryFileProvider discoveryFileProvider
            , NetZoneEnum netZone = NetZoneEnum.Any
            , IMemoryCache memoryCache = null)
        {
            DiscoveryFileProvider = discoveryFileProvider;
            NetZone = netZone;
            _memoryCache = memoryCache;
        }

        private async Task<IEnumerable<XElement>> GetAppsAsync()
        {
            return (await DiscoveryFileProvider.GetDiscoveryXmlAsync())
                .Elements(ELEMENT_NET_ZONE)
                .Where(ValidateNetZone)
                .Elements(ELEMENT_APP);
        }
        private async Task<XElement> GetProofKeyAsync()
        {
            return (await DiscoveryFileProvider.GetDiscoveryXmlAsync())
                .Elements(PROOF_KEY).FirstOrDefault();
        }

        private bool ValidateNetZone(XElement e)
        {
            if (NetZone != NetZoneEnum.Any)
            {
                var netZoneString = (string)e.Attribute(ATTR_NET_ZONE_NAME);
                netZoneString = netZoneString.Replace("-", "");
                bool success = Enum.TryParse(netZoneString, true, out NetZoneEnum netZone);
                return success && (netZone == NetZone);
            }
            return true;
        }

        public async Task<bool> SupportsExtensionAsync(string extension)
        {
            var query = (await GetAppsAsync()).Elements()
                .FirstOrDefault(e => (string)e.Attribute(ATTR_ACTION_EXTENSION) == extension);
            return query != null;
        }

        public async Task<bool> SupportsActionAsync(string extension, WopiActionEnum action)
        {
            string actionString = action.ToString().ToLower();

            var query = (await GetAppsAsync()).Elements().Where(e => (string)e.Attribute(ATTR_ACTION_EXTENSION) == extension && (string)e.Attribute(ATTR_ACTION_NAME) == actionString);

            return query.Any();
        }

        public async Task<IEnumerable<string>> GetActionRequirementsAsync(string extension, WopiActionEnum action)
        {
            string actionString = action.ToString().ToLower();

            var query = (await GetAppsAsync()).Elements().Where(e => (string)e.Attribute(ATTR_ACTION_EXTENSION) == extension && (string)e.Attribute(ATTR_ACTION_NAME) == actionString).Select(e => e.Attribute(ATTR_ACTION_REQUIRES).Value.Split(','));

            return query.FirstOrDefault();
        }

        public async Task<bool> RequiresCobaltAsync(string extension, WopiActionEnum action)
        {
            var requirements = await GetActionRequirementsAsync(extension, action);
            return requirements != null && requirements.Contains(ATTR_VAL_COBALT);
        }

        public async Task<string> GetUrlTemplateAsync(string extension, WopiActionEnum action)
        {
            string actionString = action.ToString().ToLower();

            var query = (await GetAppsAsync()).Elements().Where(e => (string)e.Attribute(ATTR_ACTION_EXTENSION) == extension && (string)e.Attribute(ATTR_ACTION_NAME) == actionString).Select(e => e.Attribute(ATTR_ACTION_URL).Value);

            return query.FirstOrDefault();
        }

        public async Task<string> GetApplicationNameAsync(string extension)
        {
            var query = (await GetAppsAsync()).Where(e => e.Descendants(ELEMENT_ACTION).Any(d => (string)d.Attribute(ATTR_ACTION_EXTENSION) == extension)).Select(e => e.Attribute(ATTR_APP_NAME).Value);

            return query.FirstOrDefault();
        }

        public async Task<string> GetApplicationFavIconAsync(string extension)
        {
            var query = (await GetAppsAsync()).Where(e => e.Descendants(ELEMENT_ACTION).Any(d => (string)d.Attribute(ATTR_ACTION_EXTENSION) == extension)).Select(e => e.Attribute(ATTR_APP_FAVICON).Value);

            return query.FirstOrDefault();
        }

        async Task<Uri> IDiscoverer.GetApplicationFavIconAsync(string extension)
        {
            var query = (await GetAppsAsync()).Where(e => e.Descendants(ELEMENT_ACTION).Any(d => (string)d.Attribute(ATTR_ACTION_EXTENSION) == extension)).Select(e => e.Attribute(ATTR_APP_FAVICON).Value);
            var result = query.FirstOrDefault();
            return result is not null ? new Uri(result) : null;
        }

        public async Task<WopiProof> GetWopiProof()
        {
            WopiProof wopiProof = null;
            // Check cache for this data
            bool isCache = _memoryCache?.TryGetValue(PROOF_KEY_CACHE, out wopiProof) ?? false;
            if (!isCache)
            {
                XElement proof = await GetProofKeyAsync();
                wopiProof = new WopiProof()
                {
                    value = proof.Attribute("value").Value,
                    modulus = proof.Attribute("modulus").Value,
                    exponent = proof.Attribute("exponent").Value,
                    oldvalue = proof.Attribute("oldvalue").Value,
                    oldmodulus = proof.Attribute("oldmodulus").Value,
                    oldexponent = proof.Attribute("oldexponent").Value
                };
                // Add to cache for 20min
                _memoryCache?.Set(PROOF_KEY_CACHE, wopiProof, DateTimeOffset.UtcNow.AddMinutes(20));
            }
            return await Task.FromResult(wopiProof);
        }
    }
}
