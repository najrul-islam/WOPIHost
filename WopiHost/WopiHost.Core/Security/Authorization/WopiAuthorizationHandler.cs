﻿using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using WopiHost.Abstractions;

namespace WopiHost.Core.Security.Authorization
{
    /// <summary>
    /// Performs resource-based authorization.
    /// </summary>
    public class WopiAuthorizationHandler : AuthorizationHandler<WopiAuthorizationRequirement, FileResource>
    {
        public IWopiSecurityHandler SecurityHandler { get; }

        public WopiAuthorizationHandler(IWopiSecurityHandler securityHandler)
        {
            SecurityHandler = securityHandler;
        }

        protected override Task HandleRequirementAsync(AuthorizationHandlerContext context, WopiAuthorizationRequirement requirement, FileResource resource)
        {
            if (SecurityHandler.IsAuthorized(context.User, resource.FileId, requirement))
            {
                context.Succeed(requirement);
            }
            else
            {
                context.Fail();
            }
            return Task.CompletedTask;
        }
    }
}
