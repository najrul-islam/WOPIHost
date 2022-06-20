using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WopiHost.Utility.Enum
{
    public enum EnumCategoryTypeGroup
    {
        Fees = 1,
        Expense = 2,
        Disbursement = 3
    }
    public enum EnumTaxType
    {
        All = 1,
        WithoutTax = 2,
        WithTax = 3
    }
    public enum EnumClaimbaleType
    {
        Claimbale = 1,
        NonClaimbale = 2
    }
    public enum EnumBillingActionTypes
    {
        Letter = 1,
        PhoneIn = 2,
        PhoneOut = 3,
        Email = 4,
        Memo = 6,
        SMS = 8,
        Fees = 10,
        Expense = 11,
        Disbursement = 12,
        Claimbale = 13,
        Anticipated = 14,
        Incurred = 15,
        Slip = 16
    }
}
