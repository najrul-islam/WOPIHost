using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WopiHost.Utility.StaticData
{
    public class InvoiceTemplateStaticData
    {
    }

    public static class InvoiceTemplateHeaderData
    {
        public const string Fee_Header = "Fee_Header";
        public const string Fee_Header_NoTax = "Fee_Header_NoTax";
        public const string Fee_Header_Tax = "Fee_Header_Tax";

        public const string Disbursements_Header = "Disbursements_Header";
        public const string Disbursements_Header_NoTax = "Disbursements_Header_NoTax";
        public const string Disbursements_Header_Tax = "Disbursements_Header_Tax";

        public const string Expenses_Header = "Expenses_Header";
        public const string Expenses_Header_NoTax = "Expenses_Header_NoTax";
        public const string Expenses_Header_Tax = "Expenses_Header_Tax";
        //Claimbale
        public const string Expenses_Header_Claimbale = "Expenses_Header_Claimbale";
        public const string Expenses_Header_NoTax_Claimbale = "Expenses_Header_NoTax_Claimbale";
        public const string Expenses_Header_Tax_Claimbale = "Expenses_Header_Tax_Claimbale";
        //NonClaimbale
        public const string Expenses_Header_NonClaimbale = "Expenses_Header_NonClaimbale";
        public const string Expenses_Header_NoTax_NonClaimbale = "Expenses_Header_NoTax_NonClaimbale";
        public const string Expenses_Header_Tax_NonClaimbale = "Expenses_Header_Tax_NonClaimbale";
    }
    public static class InvoiceTemplateHeaderValue
    {
        public const string Fee_Header = "Fee";
        public const string Fee_Header_NoTax = "Fee Without Tax";
        public const string Fee_Header_Tax = "Fee With Tax";

        public const string Disbursements_Header = "Disbursements";
        public const string Disbursements_Header_NoTax = "Disbursements Without Tax";
        public const string Disbursements_Header_Tax = "Disbursements With Tax";

        public const string Expenses_Header = "Expenses";
        public const string Expenses_Header_NoTax = "Expenses Without Tax";
        public const string Expenses_Header_Tax = "Expenses With Tax";
        //NonClaimbale
        public const string Expenses_Header_Claimbale = "Expenses Claimbale";
        public const string Expenses_Header_NoTax_Claimbale = "Expenses Without Tax Claimbale";
        public const string Expenses_Header_Tax_Claimbale = "Expenses With Tax Claimbale";
        //NonClaimbale
        public const string Expenses_Header_NonClaimbale = "Expenses NonClaimbale";
        public const string Expenses_Header_NoTax_NonClaimbale = "Expenses Without Tax NonClaimbale";
        public const string Expenses_Header_Tax_NonClaimbale = "Expenses With Tax NonClaimbale";
    }
}
