using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Security.Claims;
using System.Security.Principal;

namespace WopiHost.Abstractions
{
    public static class ClaimIdentityExtention
    {
        /// <summary>
        /// Get UserId of Current Logged User as Int32
        /// </summary>
        /// <param name="identity"></param>
        /// <returns></returns>
        public static int GetUserIdFromClaimIdentity(this IPrincipal identity)
        {
            ClaimsIdentity claimsIdentity = identity.Identity as ClaimsIdentity;
            Claim claim = claimsIdentity?.FindFirst(WopiClaimTypes.UserId.ToString());
            return Convert.ToInt32(claim?.Value);
        }

        /// <summary>
        /// Get CompanyId of Current Logged User as Int32
        /// </summary>
        /// <param name="identity"></param>
        /// <returns></returns>
        public static int GetCompanyIdFromClaimIdentity(this IPrincipal identity)
        {
            ClaimsIdentity claimsIdentity = identity.Identity as ClaimsIdentity;
            Claim claim = claimsIdentity?.FindFirst(WopiClaimTypes.CompanyId.ToString());
            return Convert.ToInt32(claim?.Value);
        }
        
        /// <summary>
        /// Get Email of Current Logged User as string
        /// </summary>
        /// <returns></returns>
        public static string GetEmailFromClaimIdentity(this IPrincipal identity)
        {
            ClaimsIdentity claimsIdentity = identity.Identity as ClaimsIdentity;
            Claim claim = claimsIdentity?.FindFirst(WopiClaimTypes.Email.ToString());
            return claim?.Value;
        }

        /// <summary>
        /// Get CompanyName of Current Logged User as string
        /// </summary>
        /// <returns></returns>
        public static string GetCompanyNameFromClaimIdentity(this IPrincipal identity)
        {
            ClaimsIdentity claimsIdentity = identity.Identity as ClaimsIdentity;
            Claim claim = claimsIdentity?.FindFirst(WopiClaimTypes.CompanyName.ToString());
            return claim?.Value;
        }

        /// <summary>
        /// Get UserName of Current Logged User as string
        /// </summary>
        /// <returns></returns>
        public static string GetUserNameFromClaimIdentity(this IPrincipal identity)
        {
            ClaimsIdentity claimsIdentity = identity.Identity as ClaimsIdentity;
            Claim claim = claimsIdentity?.FindFirst(WopiClaimTypes.UserName.ToString());
            return claim?.Value;
        }

        /// <summary>
        /// Get UserName of Current Logged User as string
        /// </summary>
        /// <returns></returns>
        public static string GetUserFullNameFromClaimIdentity(this IPrincipal identity)
        {
            ClaimsIdentity claimsIdentity = identity.Identity as ClaimsIdentity;
            Claim claim = claimsIdentity?.FindFirst(WopiClaimTypes.UserFullName.ToString());
            return claim?.Value;
        }
        
       
        /// <summary>
        /// Get CompanyUserStorage of Current Logged User as List<UserCompanyStorageClaim>
        /// </summary>
        /// <returns></returns>
        public static UserCommonClaims GetCommonClaimFromClaimIdentity(this IPrincipal identity)
        {
            ClaimsIdentity claimsIdentity = identity.Identity as ClaimsIdentity;
            var res = new UserCommonClaims
            {
                CompanyId = Convert.ToInt32(claimsIdentity?.FindFirst(WopiClaimTypes.CompanyId)?.Value ?? "0"),
                UserId = Convert.ToInt32(claimsIdentity?.FindFirst(WopiClaimTypes.UserId)?.Value ?? "0"),
                StorageName = claimsIdentity?.FindFirst(WopiClaimTypes.StorageName)?.Value,
                Email = claimsIdentity?.FindFirst(WopiClaimTypes.Email)?.Value,
                UserFullName = claimsIdentity?.FindFirst(WopiClaimTypes.UserFullName)?.Value,
                LanguageCountryId = Convert.ToInt32(claimsIdentity?.FindFirst(WopiClaimTypes.LanguageCountryId)?.Value ?? "0")
            };
            return res;
        }
    }

    public class UserCommonClaims
    {
        public int UserId { get; set; }
        public int CompanyId { get; set; }
        public string UserFullName { get; set; }
        public string StorageName { get; set; }
        public string Email { get; set; }
        public int LanguageCountryId { get; set; }
    }
}
