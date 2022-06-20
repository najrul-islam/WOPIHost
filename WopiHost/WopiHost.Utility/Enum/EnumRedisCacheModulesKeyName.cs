using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WopiHost.Utility.Enum
{
    public enum EnumRedisCacheModulesKeyName
    {
        GetRightOrderUserWiseList,
        GetMatterDocumentApps,
        SyncTableValueCompanyContactsByCardId,
        GetRightsUserPreference,
        WOPILock, //Use for maintain WOPI Locks
        WOPIMattterDocumentEdit //Use for maintain MatterDocument Edit Locks
    }
}
