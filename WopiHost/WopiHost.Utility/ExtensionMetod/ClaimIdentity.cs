using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using WopiHost.Abstractions;
using WopiHost.Utility.ViewModel;

namespace WopiHost.Utility.ExtensionMetod
{
    public static class ClaimIdentity
    {
        public static async Task<WopiClaimTypesVm> GetClaimInfoFromClaimIdentity(this IPrincipal identity)
        {
            ClaimsIdentity claimsIdentity = identity.Identity as ClaimsIdentity;
            var res = new WopiClaimTypesVm
            {
                Email = claimsIdentity?.FindFirst(WopiClaimTypes.Email)?.Value,
                UserFullName = claimsIdentity?.FindFirst(WopiClaimTypes.UserFullName)?.Value,
                Name = claimsIdentity?.FindFirst(WopiClaimTypes.Name)?.Value,
                WopiSrc = claimsIdentity?.FindFirst(WopiClaimTypes.WopiSrc)?.Value,
                StorageName = claimsIdentity?.FindFirst(WopiClaimTypes.StorageName)?.Value,

            };
            return await Task.FromResult(res);
        }
    }
}
