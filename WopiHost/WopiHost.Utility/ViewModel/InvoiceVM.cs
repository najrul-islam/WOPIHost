using System;
using System.Collections.Generic;
using System.Text;

namespace WopiHost.Utility.ViewModel
{
    public class InvoiceVM
    {
        public Invoice Invoice { get; set; }
    }
    public class Invoice
    {
        public Invoice()
        {
            InvoiceDetails = new HashSet<InvoiceDetails>();
            DisbursementDetails = new HashSet<DisbursementDetails>();
            ExpenseDetails = new HashSet<ExpenseDetails>();
        }
        public int InvoiceId { get; set; }
        public string InvoiceGuid { get; set; }
        public int CardId { get; set; }
        public int CardScreenDetailsId { get; set; }
        public string TransactionNumber { get; set; }
        public string DisplayCardName { get; set; }
        public int ProductsTypeQuantity { get; set; }
        public decimal Amount { get; set; }
        public decimal? Tax { get; set; }
        public decimal Total { get; set; }
        public int? CurrencyId { get; set; }
        public DateTime Date { get; set; }
        public DateTime? DueDate { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime ModifiedDate { get; set; }
        public int? CreatedBy { get; set; }
        public int? ModifiedBy { get; set; }
        public bool IsDeleted { get; set; }
        public bool IsActive { get; set; }
        public int? StatusId { get; set; }
        public int? BranchId { get; set; }
        public int CompanyId { get; set; }
        public int? OriginId { get; set; }
        public string SerialNumber { get; set; }
        public int? SystemSerialNumber { get; set; }
        public int? SubsidiaryType { get; set; }
        public int MatterId { get; set; }
        public int? TemplateId { get; set; }
        public int? InvoiceTypeId { get; set; }
        public string ReferenceNo { get; set; }
        public int? DirectReverseInvoiceId { get; set; }
        public int PaymentStatusId { get; set; }
        public int ReceiptBankId { get; set; }
        public int AccountTypeId { get; set; }
        public string ProviderInvoiceGuid { get; set; }
        public int? ProviderTypeId { get; set; }
        public int? ProviderStatusId { get; set; }
        public DateTime? ProcessDate { get; set; }
        public virtual ICollection<InvoiceDetails> InvoiceDetails { get; set; }
        public virtual ICollection<DisbursementDetails> DisbursementDetails { get; set; }
        public virtual ICollection<ExpenseDetails> ExpenseDetails { get; set; }
    }
    public class InvoiceDetails
    {
        public int InvoiceDetailsId { get; set; }
        public string InvoiceDetailsGuid { get; set; }
        public int InvoiceId { get; set; }
        public int? ProductId { get; set; }
        public string ProductDescription { get; set; }
        public int? ProductQuantity { get; set; }
        public decimal? ProductUnitPrice { get; set; }
        public decimal Amount { get; set; }
        public int? TaxId { get; set; }
        public decimal TaxAmount { get; set; }
        public decimal Total { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime ModifiedDate { get; set; }
        public bool IsDeleted { get; set; }
        public bool IsActive { get; set; }
        public int? StatusId { get; set; }
        public string MatterDocumentIds { get; set; }
        public int? MatterId { get; set; }
        public int? CardId { get; set; }
        public int? CardScreenDetailsId { get; set; }
        public string DisplayCardName { get; set; }
        public bool IsDisbursement { get; set; }
        public int? CategoryTypeGroupId { get; set; }

        public int? ObjectTypeId { get; set; }
        public int? ObjectId { get; set; }
        public int? LaActivityId { get; set; }
        public int? TaskId { get; set; }
        public string ActivityDescription { get; set; }
        public int? DocumentId { get; set; }
        public DateTime? InvoiceDetailsDate { get; set; }
        public bool? IsBillable { get; set; }
        public decimal? Hours { get; set; }
        public int? UserId { get; set; }
        public int? DirectReverseInvoiceDetailsId { get; set; }
        public int? CompanyId { get; set; }
        public int? ActionTypeId { get; set; }

    }
    public class DisbursementDetails
    {
        public int DisbursementDetailsId { get; set; }
        public int InvoiceId { get; set; }
        public string DisbursementDetailsGuid { get; set; }
        public int? ProductId { get; set; }
        public string ProductDescription { get; set; }
        public int? ProductQuantity { get; set; }
        public decimal? ProductUnitPrice { get; set; }
        public decimal Amount { get; set; }
        public int? TaxId { get; set; }
        public decimal TaxAmount { get; set; }
        public decimal Total { get; set; }
        public decimal DrAmount { get; set; }
        public decimal DrAmountTax { get; set; }
        public decimal CrAmount { get; set; }
        public decimal CrAmountTax { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime ModifiedDate { get; set; }
        public bool IsDeleted { get; set; }
        public bool IsActive { get; set; }
        public int? StatusId { get; set; }
        public int? CardId { get; set; }
        public int? CardScreenDetailsId { get; set; }
        public int? DisbursementAccountListId { get; set; }
        public string DisplayCardName { get; set; }
        public bool IsDisbursement { get; set; }
        public int? CategoryTypeGroupId { get; set; }
        public int? ExpenseTaskId { get; set; }
        public string ActivityDescription { get; set; }
        public string ReferenceNo { get; set; }
        public DateTime? DisbursementDetailsDate { get; set; }
        public bool? IsBillable { get; set; }
        public string UploadedInvoiceNo { get; set; }
        public string UploadedInvoiceDetails { get; set; }
        public string UploadedInvoiceReference { get; set; }
        public decimal? Hours { get; set; }
        public int? UserId { get; set; }
        public int? MatterId { get; set; }
        public bool IsLoose { get; set; }
        public int CompanyId { get; set; }
        public int? DirectReverseDisbursementDetailsId { get; set; }
    }
    public class ExpenseDetails
    {
        public int ExpenseDetailsId { get; set; }
        public int InvoiceId { get; set; }
        public string ExpenseDetailsGuid { get; set; }
        public int? ProductId { get; set; }
        public string ProductDescription { get; set; }
        public int? ProductQuantity { get; set; }
        public decimal? ProductUnitPrice { get; set; }
        public decimal Amount { get; set; }
        public int? TaxId { get; set; }
        public decimal TaxAmount { get; set; }
        public decimal Total { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime ModifiedDate { get; set; }
        public bool IsDeleted { get; set; }
        public bool IsActive { get; set; }
        public int? StatusId { get; set; }
        public int? CardId { get; set; }
        public int? CardScreenDetailsId { get; set; }
        public string DisplayCardName { get; set; }
        public bool IsDisbursement { get; set; }
        public int? CategoryTypeGroupId { get; set; }
        public bool? IsRecoverable { get; set; }
        public int? ExpenseTaskId { get; set; }
        public string ActivityDescription { get; set; }
        public DateTime? ExpenseDetailsDate { get; set; }
        public bool? IsBillable { get; set; }
        public decimal? Hours { get; set; }
        public int? UserId { get; set; }
        public int? DirectReverseExpenseDetailsId { get; set; }
        public int? CompanyId { get; set; }
    }

    public class InvoiceTemplateMergeVM
    {
        public Invoice InvoiceVM { get; set; }
        public List<GlobalDataVm> GDatas { get; set; }
    }
}
