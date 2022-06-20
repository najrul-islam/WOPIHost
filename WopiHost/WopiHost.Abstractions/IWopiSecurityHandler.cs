﻿using System.Security.Claims;
using Microsoft.IdentityModel.Tokens;

namespace WopiHost.Abstractions
{
    /// <summary>
    /// Performs security-related actions.
    /// </summary>
    public interface IWopiSecurityHandler
    {
        /// <summary>
        /// Generates authorization token for the given value.
        /// </summary>
        /// <returns>Authorization token</returns>
        SecurityToken GenerateAccessToken(int userId, string email, string userFullName, WopiDocsTypeEnum wopiDocsType = WopiDocsTypeEnum.WopiDocs, WopiVersionableTypeEnum isVersionable = WopiVersionableTypeEnum.True, string fileOriginalName = "", string wopiSrc = "");

        ClaimsPrincipal GetPrincipal(string token);


        bool IsAuthorized(ClaimsPrincipal principal, string resourceId, WopiAuthorizationRequirement operation);

        /// <summary>
        /// Converts the security token to a Base64 string.
        /// </summary>
        string WriteToken(SecurityToken token);

        long GenerateAccessToken_TTL();
    }
}
