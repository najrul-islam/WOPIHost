using Microsoft.IdentityModel.Tokens;
using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Security.Claims;
using System.Text;
using WopiHost.Abstractions;

namespace WopiHost.FileSystemProvider
{
    public class WopiSecurityHandler : IWopiSecurityHandler
    {
        private readonly JwtSecurityTokenHandler tokenHandler = new();
        private SymmetricSecurityKey _key = null;

        private SymmetricSecurityKey Key
        {
            get
            {
                if (_key == null)
                {
                    //RandomNumberGenerator rng = RandomNumberGenerator.Create();
                    //byte[] key = new byte[128];
                    //rng.GetBytes(key);
                    //var key = Encoding.ASCII.GetBytes("secretKeysecretKeysecretKey123"/* + new Random(DateTime.Now.Millisecond).Next(1,999)*/);
                    var key = Encoding.ASCII.GetBytes("32a57bf0-api-47ab-hoxro-3d87d3b3a47f");
                    _key = new SymmetricSecurityKey(key);
                }
                return _key;
            }
        }

        //TODO: abstract
        private readonly Dictionary<string, ClaimsPrincipal> UserDatabase = new Dictionary<string, ClaimsPrincipal>
        {
            {"Anonymous",new ClaimsPrincipal(
                new ClaimsIdentity(new List<Claim>
                {
                    new Claim(WopiClaimTypes.UserId, "12345"),
                    new Claim(WopiClaimTypes.UserFullName, "Anonymous"),
                    new Claim(ClaimTypes.Name, "Anonymous"),
                    new Claim(WopiClaimTypes.Email, "anonymous@domain.com"),
                    new Claim(WopiClaimTypes.UserPermissions, (WopiUserPermissions.UserCanWrite | WopiUserPermissions.UserCanRename | WopiUserPermissions.UserCanAttend | WopiUserPermissions.UserCanPresent).ToString()),
                    new Claim(WopiClaimTypes.WopiDocsType, ""),
                    new Claim(WopiClaimTypes.IsVersionable, "True")
                })
            ) }
        };

        public SecurityToken GenerateAccessToken(int userId, string email, string userFullName, WopiDocsTypeEnum wopiDocsType = WopiDocsTypeEnum.WopiDocs, WopiVersionableTypeEnum isVersionable = WopiVersionableTypeEnum.True, string fileOriginalName = "", string wopiSrc = "")
        {
            var userDb = UserDatabase.FirstOrDefault(x => x.Key == email);
            bool ok = UserDatabase.TryGetValue(email, out ClaimsPrincipal user);
            if (!ok)
            {
                UserDatabase.Add(
                    email, new ClaimsPrincipal(
                        new ClaimsIdentity(new List<Claim>
                        {
                            new Claim(WopiClaimTypes.UserId, userId.ToString()),
                            new Claim(WopiClaimTypes.UserFullName, userFullName),
                            new Claim(ClaimTypes.Name, userFullName),
                            new Claim(WopiClaimTypes.Email, email),
                            new Claim(WopiClaimTypes.FileOriginalName, fileOriginalName),
                            new Claim(WopiClaimTypes.UserPermissions, (WopiUserPermissions.UserCanWrite | WopiUserPermissions.UserCanRename | WopiUserPermissions.UserCanAttend | WopiUserPermissions.UserCanPresent).ToString()),
                            new Claim(WopiClaimTypes.WopiDocsType, wopiDocsType.ToString()),
                            new Claim(WopiClaimTypes.IsVersionable, isVersionable.ToString()),
                            new Claim(WopiClaimTypes.WopiSrc, wopiSrc),
                        })
                    ));
                user = UserDatabase[email];
            }


            var tokenDescriptor = new SecurityTokenDescriptor
            {
                Subject = user.Identities.FirstOrDefault(),
                Expires = DateTime.UtcNow.AddHours(24),
                //Expires = DateTime.UtcNow.AddDays(1),
                SigningCredentials = new SigningCredentials(Key, SecurityAlgorithms.HmacSha256)
            };
            return tokenHandler.CreateToken(tokenDescriptor);
        }

        public ClaimsPrincipal GetPrincipal(string tokenString)
        {
            //TODO: https://github.com/aspnet/Security/tree/master/src/Microsoft.AspNetCore.Authentication.JwtBearer

            var tokenValidation = new TokenValidationParameters
            {
                ValidateAudience = false,
                ValidateIssuer = false,
                ValidateActor = false,
                ValidateLifetime = true,
                ValidateIssuerSigningKey = true,
                IssuerSigningKey = Key
            };

            try
            {
                // Try to validate the token
                var principal = tokenHandler.ValidateToken(tokenString, tokenValidation, out SecurityToken token);
                return principal;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public bool IsAuthorized(ClaimsPrincipal principal, string resourceId, WopiAuthorizationRequirement operation)
        {
            return true;
        }

        /// <summary>
        /// Converts the security token to a Base64 string.
        /// </summary>
        public string WriteToken(SecurityToken token)
        {
            return tokenHandler.WriteToken(token);
        }

        public long GenerateAccessToken_TTL()
        {
            return ((DateTimeOffset)DateTime.UtcNow.AddHours(24)).ToUnixTimeMilliseconds();//Timestamp in Milliseconds
        }
    }
}
