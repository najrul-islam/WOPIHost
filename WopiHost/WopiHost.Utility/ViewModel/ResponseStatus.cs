using System;
using System.Collections.Generic;

namespace WopiHost.Utility.ViewModel
{
    public class ResponseStatus
    {
        public int StatusCode { get; set; }
        public string Message { get; set; }
        public ResponseStatus()
        {
            StatusCode = 200;
            Message = "Success";
        }
        /*public ResponseStatus(Exception exception, string blob)
        {
            StatusCode = exception.RequestInformation.HttpStatusCode;
            Message = exception.RequestInformation.HttpStatusMessage.Replace("blob", $"blob '{blob}'");
        }*/

    }
    public class WopiUrlResponse
    {
        public string Url { get; set; }
        public string ViewUrl { get; set; }
        public string EditUrl { get; set; }
        public int StatusCode { get; set; } = 200;
        public string Message { get; set; } = "Success";
        public string FileName { get; set; }
    }
    public class NewContentTemplateResponse : WopiUrlResponse
    {
        public string NewFileName { get; set; }
    }
    public class NewLetterResponse : NewContentTemplateResponse
    {
        public string Content { get; set; }
        public List<GlobalDataVm> UnMergeGlobalDataList { get; set; } = new List<GlobalDataVm>();
        public List<UnAttachedCardVM> ListUnAttachedCardVm { get; set; } = new List<UnAttachedCardVM>();
    }
}
