using System;
using System.Collections.Generic;
using System.Text;

namespace WopiHost.Utility.ViewModel
{
    public class WopiResponseUnMailMergeVM
    {
        public WopiResponseUnMailMergeVM()
        {
            ContentDocumnets = new List<ContentDocument>();
        }
        public List<ContentDocument> ContentDocumnets { get; set; }
    }
    public class WopiRequestUnMailMergeVM
    {
        public List<ContentDocument> ContentDocumnets { get; set; }
        public List<GlobalDataVm> ListGlobalDataVm { get; set; }
    }
    public class ContentDocument : WopiUrlResponse
    {
        public int ContentDocumentId { get; set; }
        public string Bolb { get; set; }
        public int CompanyId { get; set; }
        public List<GlobalDataVm> ListUnMergeGlobalDataVm { get; set; }
    }
    public class GlobalDataVm
    {
        public int GlobalDataId { get; set; }
        public string Name { get; set; }
        public string InitialValue { get; set; }
        public string InputTypeValues { get; set; }
        public string InitialValueMultiple { get; set; }
        public int CompanyId { get; set; }
        public int FieldId { get; set; }
        public int FieldGroupId { get; set; }
        public int? CardId { get; set; }
        public string CardName { get; set; }
        public int? MatterId { get; set; }
        public string DataType { get; set; }
        public int? CardScreenDetailsId { get; set; }
        public string InputType { get; set; }
        public int ObjectId { get; set; }
        public int ObjectTypeId { get; set; }
        public string Format { get; set; }
        public string LabelName { get; set; }
    }
    public class UnAttachedCardVM
    {
        public int CardId { get; set; }
        public string CardName { get; set; }
        public int FieldId { get; set; }
        public string Name { get; set; }
    }
    public class InitialValueMultiple
    {
        public int GlobalDataId { get; set; }
        public string InitialValue { get; set; }
        public string CardName { get; set; }
        public int CardId { get; set; }
        public int CardScreenDetailsId { get; set; }
        public string DisplayCardName { get; set; }
        public bool IsLead { get; set; }
    }
}
