using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WopiHost.Discovery.Enumerations;

namespace WopiHost.Utility.ViewModel
{
    public class WopiClaimTypesVm
    {
        public string UserId { get; set; }
        public string UserName { get; set; }
        public string Name { get; set; }
        public string CompanyName { get; set; }
        public string CompanyId { get; set; }
        public string Email { get; set; }
        public string StorageName { get; set; }
        public string Rights { get; set; }
        public string TimeZoneId { get; set; }
        public string SessionId { get; set; }
        public string DateTimeFormat { get; set; }
        public string UserFullName { get; set; }
        public string CompanyGuid { get; set; }
        public string RoleId { get; set; }
        public string LanguageCountryId { get; set; }
        public string BranchId { get; set; }
        public string BranchGuid { get; set; }
        public string UserTypeId { get; set; }
        public string IsRedisCacheEnable { get; set; }
        public string UserPermissions { get; set; }
        public string WopiDocsType { get; set; }
        public string IsVersionable { get; set; }
        public string FileOriginalName { get; set; }


        public string WopiSrc { get; set; }
        public string UrlSr { get; set; }
        public WopiActionEnum WopiActionEnum { get; set; }
    }
}
