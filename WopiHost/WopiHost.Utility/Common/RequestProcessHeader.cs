using Microsoft.AspNetCore.Http;

namespace WopiHost.Utility.Common
{
    public class RequestProcessHeader
    {
        public RequestProcessHeader(HttpRequest request)
        {
            Authorization = request.Headers[WopiHeaders.Authorization];
            WopiSrc = request.Headers[WopiHeaders.WopiSrc];
            Action = request.Headers[WopiHeaders.Action];
            StorageName = request.Headers[WopiHeaders.StorageName];
            Checked = request.Headers[WopiHeaders.Checked];
            SourceBlob = request.Headers[WopiHeaders.SourceBlob];
            SourceTemplateURL = request.Headers[WopiHeaders.SourceTemplateURL];
            SignatureUrl = request.Headers[WopiHeaders.SignatureUrl];
            NewBlob = request.Headers[WopiHeaders.NewBlob];
            OperationType = request.Headers[WopiHeaders.OperationType];
            NewDocumentName = request.Headers[WopiHeaders.NewDocumentName];
            FileName = request.Headers[WopiHeaders.FileName];
            FileOriginalName = request.Headers[WopiHeaders.FileOriginalName];
            Extenssion = request.Headers[WopiHeaders.Extenssion].ToString().Replace(".", "");
            Override = request.Headers[WopiHeaders.Override];
            UserName = request.Headers[WopiHeaders.UserName];
            UserFullName = request.Headers[WopiHeaders.UserFullName];
            UI_LLCC = request.Headers[WopiHeaders.UI_LLCC];
            SkipFieldIds = request.Headers[WopiHeaders.SkipFieldIds];
            IsMatterDocument = request.Headers[WopiHeaders.IsMatterDocument];
            ParentId = request.Headers[WopiHeaders.ParentId];
            VersionNo = request.Headers[WopiHeaders.VersionNo];
            MatterDocumentId = request.Headers[WopiHeaders.MatterDocumentId];
        }
        public string Authorization { get; set; }
        public string WopiSrc { get; set; }
        public string Action { get; set; }
        public string StorageName { get; }
        public string Checked { get; }
        public string SourceBlob { get; }
        public string SourceTemplateURL { get; }
        public string SignatureUrl { get; }
        public string NewBlob { get; }
        public string OperationType { get; }
        public string NewDocumentName { get; }
        public string FileName { get; set; }
        public string FileOriginalName { get; set; } = string.Empty;
        public string Extenssion { get; }
        public string Override { get; }
        public string UserName { get; }
        public string UserFullName { get; }
        public string UI_LLCC { get; }
        public string SkipFieldIds { get; }
        public string IsMatterDocument { get; }
        public string ParentId { get; }
        public string VersionNo { get; }
        public string MatterDocumentId { get; }
    }
}
