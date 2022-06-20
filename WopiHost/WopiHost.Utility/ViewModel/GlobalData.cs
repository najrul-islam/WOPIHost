using System;
using System.Collections.Generic;
using System.Text;

namespace WopiHost.Utility.ViewModel
{
    public class GlobalData
    {
        public int FieldId { get; set; }
        public string Name { get; set; }
        public string InitialValue { get; set; }
    }
    public class UnMergerdGlobalDataResult
    {
        public string url { get; set; }
        public List<GlobalData> UnMargeGlobalData { get; set; }
    }
}
