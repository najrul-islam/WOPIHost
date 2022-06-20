using System;
using System.Collections.Generic;
using System.Text;

namespace WopiHost.Utility.ViewModel
{
    public class MatterStageActivityMailMergeResponse
    {
        public List<DocumentContent> DocumentContents { get; set; }
    }
    public class MatterStageActivityMailMergeRequest : MatterStageActivityMailMergeResponse
    {
        public List<GlobalDataVm> GDatas { get; set; }
    }
    public class DocumentContent : WopiUrlResponse
    {
        public string ContentBlob { get; set; }
        public string TemplateBlob { get; set; }
        public int WorkFlowMatterStageActivityId { get; set; }
        public int MatterId { get; set; }
        public int LibraryContentId { get; set; }
        public string Name { get; set; }
        //public string FileName { get; set; }
        public string Blob { get; set; }
        public string Details { get; set; }
    }
}
