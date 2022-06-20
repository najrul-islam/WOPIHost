using System;
using System.Collections.Generic;
using System.Text;

namespace WopiHost.Utility.Common
{
    public static class WopiHeaders
    {

        public const string WopiOverride = "X-WOPI-Override";
        public const string Lock = "X-WOPI-Lock";
        public const string OldLock = "X-WOPI-OldLock";
        public const string LockFailureReason = "X-WOPI-LockFailureReason";
        public const string LockedByOtherInterface = "X-WOPI-LockedByOtherInterface";

        public const string SuggestedTarget = "X-WOPI-SuggestedTarget";
        public const string RelativeTarget = "X-WOPI-RelativeTarget";
        public const string OverwriteRelativeTarget = "X-WOPI-OverwriteRelativeTarget";
        public const string CorrelationId = "X-WOPI-CorrelationID";
        public const string MaxExpectedSize = "X-WOPI-MaxExpectedSize";
        public const string WopiSrc = "X-WOPI-WopiSrc";
        public const string EcosystemOperation = "X-WOPI-EcosystemOperation";
        public const string WopiItemVersion = "X-WOPI-ItemVersion";
        public const string Proof = "X-WOPI-Proof";
        public const string Proof_Old = "X-WOPI-ProofOld";
        public const string Time_Stamp = "X-WOPI-TimeStamp";
        //Modify
        public const string StorageName = "X-WOPI-StorageName";
        public const string SourceBlob = "X-WOPI-Azure-Blob";
        public const string SourceTemplateURL = "X-WOPI-Azure-Blob-Template";
        public const string SignatureUrl = "X-WOPI-Azure-Blob-Signature-Template";
        public const string NewBlob = "X-WOPI-Azure-New-Blob";

        public const string EnableCoAuthoring = "X-WOPI-EnableCoAuthoring";
        public const string Action = "X-WOPI-Action";
        public const string Extenssion = "X-WOPI-Extenssion";
        public const string FileName = "X-WOPI-FileName";
        public const string FileOriginalName = "X-WOPI-FileOriginalName";
        public const string UserName = "X-WOPI-UserName";
        public const string UserFullName = "X-WOPI-UserFullName";
        public const string Override = "X-WOPI-Override";
        //public const string BlankTemplateId = "X-WOPI-TemplateId";
        public const string OperationType = "X-WOPI-OperationType";
        public const string NewDocumentName = "X-WOPI-NewDocumentName";
        //public const string SearchFileName = "X-WOPI-SearchFileName";
        public const string Checked = "X-WOPI-Checked-Unlock";
        public const string IsMatterDocument = "X-WOPI-IsMatterDocument";
        public const string MatterDocumentId = "X-WOPI-MatterDocumentId";
        public const string ParentId = "X-WOPI-ParentId";
        public const string VersionNo = "X-WOPI-VersionNo";
        public const string UI_LLCC = "X-WOPI-UI_LLCC";
        public const string SkipFieldIds = "X-WOPI-SkipFieldIds";
        public const string Authorization = "Authorization";
    }
}
