using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using WopiHost.Utility.Common;
using System.Collections.Generic;
using WopiHost.Utility.ViewModel;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using WopiHost.Utility.Enum;
using WopiHost.Utility.StaticData;

namespace WopiHost.Utility.XMLProcess
{
    public class ProcessInvoiceTemplateMerge
    {
        public static WopiUrlResponse InvoiceTemplateMerge(string path, Invoice invoiceVM, List<GlobalDataVm> gDatas)
        {
            WopiUrlResponse wopiUrl = new();
            string innerFullText = "";
            using (var stream = File.Open(path, FileMode.Open))
            {
                using (var doc = WordprocessingDocument.Open(stream, true))
                {
                    innerFullText = doc.MainDocumentPart.Document.InnerText;
                    MargeInvoice(doc.MainDocumentPart.Document, invoiceVM, gDatas);
                    doc.Save();
                    doc.Close();
                }
                stream.Close();
            }
            #region Invoice Data Merge
            if (!string.IsNullOrEmpty(innerFullText))
            {
                Regex regex = new(InvoiceMailMergeRegex.InvoiceMailMergePattern, RegexOptions.Compiled);
                IEnumerable<string> matchList = regex.Matches(innerFullText).Cast<Match>().Select(x => x.Value).Distinct();
                if (matchList?.Count() > 0)
                {
                    //search and replace
                    //ProcessBulkXMLMailMerge.MailMergeOpenXmlRegexInvoice(path, matchList, gDatas);
                    ProcessBulkXMLMailMerge.MailMergeWithGlobalData(path, regex, InvoiceMailMergeRegex.InvoiceMailMergePattern, InvoiceMailMergeRegex.StartText, InvoiceMailMergeRegex.EndText, gDatas, new NewLetterResponse());
                }
            }
            #endregion
            return wopiUrl;
        }
        public static bool MargeInvoice(OpenXmlElement xmlElement, Invoice invoiceVM, List<GlobalDataVm> gDatas)
        {
            try
            {
                //Header data
                //InvoiceTemplateStaticData invoiceTemplateStaticData;
                #region Line Item Identifire Columns
                //Fees
                Dictionary<string, string> modelFeessColumnDictionary = new()
                {
                    { "Fee_Description", nameof(InvoiceDetails.ProductDescription) },
                    { "Fee_Quantity", nameof(InvoiceDetails.ProductQuantity) },
                    { "Fee_UnitPrice", nameof(InvoiceDetails.ProductUnitPrice) },
                    { "Fee_NetAmount", nameof(InvoiceDetails.Amount) },
                    { "Fee_TaxAmount", nameof(InvoiceDetails.TaxAmount) },
                    { "Fee_GrossAmount", nameof(InvoiceDetails.Total) }
                };
                //Disbursment
                Dictionary<string, string> modelDisbursementColumnDictionary = new()
                {
                    { "RecoverableCosts_Description", nameof(InvoiceDetails.ProductDescription) },
                    { "RecoverableCosts_Quantity", nameof(InvoiceDetails.ProductQuantity) },
                    { "RecoverableCosts_UnitPrice", nameof(InvoiceDetails.ProductUnitPrice) },
                    { "RecoverableCosts_NetAmount", nameof(InvoiceDetails.Amount) },
                    { "RecoverableCosts_TaxAmount", nameof(InvoiceDetails.TaxAmount) },
                    { "RecoverableCosts_GrossAmount", nameof(InvoiceDetails.Total) }
                };
                //Expense
                Dictionary<string, string> modelExpenseColumnDictionary = new()
                {
                    { "Expense_Description", nameof(InvoiceDetails.ProductDescription) },
                    { "Expense_Quantity", nameof(InvoiceDetails.ProductQuantity) },
                    { "Expense_UnitPrice", nameof(InvoiceDetails.ProductUnitPrice) },
                    { "Expense_NetAmount", nameof(InvoiceDetails.Amount) },
                    { "Expense_TaxAmount", nameof(InvoiceDetails.TaxAmount) },
                    { "Expense_GrossAmount", nameof(InvoiceDetails.Total) }
                };
                #endregion

                #region Total Row Identifire Columns

                List<InvoiceDetails> feessDetails = invoiceVM?.InvoiceDetails?.Where(x => x.CategoryTypeGroupId == (int)EnumCategoryTypeGroup.Fees)?.ToList();
                List<InvoiceDetails> disbursementDetails = invoiceVM?.InvoiceDetails?.Where(x => x.CategoryTypeGroupId == (int)EnumCategoryTypeGroup.Disbursement)?.ToList();
                List<InvoiceDetails> expenseDetails = invoiceVM?.InvoiceDetails?.Where(x => x.CategoryTypeGroupId == (int)EnumCategoryTypeGroup.Expense)?.ToList();

                //Total ALL
                Dictionary<string, string> modelTotalColumnDictionary = new()
                {
                    { "Fee_Quantity_Total", feessDetails?.Sum(x => x.ProductQuantity)?.ToString() ?? "0" },
                    { "Fee_UnitPrice_Total", feessDetails?.Sum(x => x.ProductUnitPrice)?.ToString() ?? "0" },
                    { "Fee_NetAmount_Total", feessDetails?.Sum(x => x.Amount).ToString() ?? "0" },
                    { "Fee_TaxAmount_Total", feessDetails?.Sum(x => x.TaxAmount).ToString() ?? "0" },
                    { "Fee_GrossAmount_Total", feessDetails?.Sum(x => x.Total).ToString() ?? "0" },
                    { "FeeTax_Quantity_Total", feessDetails?.Sum(x => x.ProductQuantity)?.ToString() ?? "0" },
                    { "FeeTax_UnitPrice_Total", feessDetails?.Sum(x => x.ProductUnitPrice)?.ToString() ?? "0" },
                    { "FeeTax_NetAmount_Total", feessDetails?.Sum(x => x.Amount).ToString() ?? "0" },
                    { "FeeTax_TaxAmount_Total", feessDetails?.Sum(x => x.TaxAmount).ToString() ?? "0" },
                    { "FeeTax_GrossAmount_Total", feessDetails?.Sum(x => x.Total).ToString() ?? "0" },
                    { "FeeNoTax_Quantity_Total", feessDetails?.Sum(x => x.ProductQuantity)?.ToString() ?? "0" },
                    { "FeeNoTax_UnitPrice_Total", feessDetails?.Sum(x => x.ProductUnitPrice)?.ToString() ?? "0" },
                    { "FeeNoTax_NetAmount_Total", feessDetails?.Sum(x => x.Amount).ToString() ?? "0" },
                    { "FeeNoTax_TaxAmount_Total", feessDetails?.Sum(x => x.TaxAmount).ToString() ?? "0" },
                    { "FeeNoTax_GrossAmount_Total", feessDetails?.Sum(x => x.Total).ToString() ?? "0" },

                    { "RecoverableCosts_Quantity_Total", disbursementDetails?.Sum(x => x.ProductQuantity)?.ToString() ?? "0" },
                    { "RecoverableCosts_UnitPrice_Total", disbursementDetails?.Sum(x => x.ProductUnitPrice)?.ToString() ?? "0" },
                    { "RecoverableCosts_NetAmount_Total", disbursementDetails?.Sum(x => x.Amount).ToString() ?? "0" },
                    { "RecoverableCosts_TaxAmount_Total", disbursementDetails?.Sum(x => x.TaxAmount).ToString() ?? "0" },
                    { "RecoverableCosts_GrossAmount_Total", disbursementDetails?.Sum(x => x.Total).ToString() ?? "0" },
                    { "RecoverableCostsTax_Quantity_Total", disbursementDetails?.Sum(x => x.ProductQuantity)?.ToString() ?? "0" },
                    { "RecoverableCostsTax_UnitPrice_Total", disbursementDetails?.Sum(x => x.ProductUnitPrice)?.ToString() ?? "0" },
                    { "RecoverableCostsTax_NetAmount_Total", disbursementDetails?.Sum(x => x.Amount).ToString() ?? "0" },
                    { "RecoverableCostsTax_TaxAmount_Total", disbursementDetails?.Sum(x => x.TaxAmount).ToString() ?? "0" },
                    { "RecoverableCostsTax_GrossAmount_Total", disbursementDetails?.Sum(x => x.Total).ToString() ?? "0" },
                    { "RecoverableCostsNoTax_Quantity_Total", disbursementDetails?.Sum(x => x.ProductQuantity)?.ToString() ?? "0" },
                    { "RecoverableCostsNoTax_UnitPrice_Total", disbursementDetails?.Sum(x => x.ProductUnitPrice)?.ToString() ?? "0" },
                    { "RecoverableCostsNoTax_NetAmount_Total", disbursementDetails?.Sum(x => x.Amount).ToString() ?? "0" },
                    { "RecoverableCostsNoTax_TaxAmount_Total", disbursementDetails?.Sum(x => x.TaxAmount).ToString() ?? "0" },
                    { "RecoverableCostsNoTax_GrossAmount_Total", disbursementDetails?.Sum(x => x.Total).ToString() ?? "0" },

                    { "Expense_Quantity_Total", expenseDetails?.Sum(x => x.ProductQuantity)?.ToString() ?? "0" },
                    { "Expense_UnitPrice_Total", expenseDetails?.Sum(x => x.ProductUnitPrice)?.ToString() ?? "0" },
                    { "Expense_NetAmount_Total", expenseDetails?.Sum(x => x.Amount).ToString() ?? "0" },
                    { "Expense_TaxAmount_Total", expenseDetails?.Sum(x => x.TaxAmount).ToString() ?? "0" },
                    { "Expense_GrossAmount_Total", expenseDetails?.Sum(x => x.Total).ToString() ?? "0" },
                    { "ExpenseTax_Quantity_Total", expenseDetails?.Sum(x => x.ProductQuantity)?.ToString() ?? "0" },
                    { "ExpenseTax_UnitPrice_Total", expenseDetails?.Sum(x => x.ProductUnitPrice)?.ToString() ?? "0" },
                    { "ExpenseTax_NetAmount_Total", expenseDetails?.Sum(x => x.Amount).ToString() ?? "0" },
                    { "ExpenseTax_TaxAmount_Total", expenseDetails?.Sum(x => x.TaxAmount).ToString() ?? "0" },
                    { "ExpenseTax_GrossAmount_Total", expenseDetails?.Sum(x => x.Total).ToString() ?? "0" },
                    { "ExpenseNoTax_Quantity_Total", expenseDetails?.Sum(x => x.ProductQuantity)?.ToString() ?? "0" },
                    { "ExpenseNoTax_UnitPrice_Total", expenseDetails?.Sum(x => x.ProductUnitPrice)?.ToString() ?? "0" },
                    { "ExpenseNoTax_NetAmount_Total", expenseDetails?.Sum(x => x.Amount).ToString() ?? "0" },
                    { "ExpenseNoTax_TaxAmount_Total", expenseDetails?.Sum(x => x.TaxAmount).ToString() ?? "0" },
                    { "ExpenseNoTax_GrossAmount_Total", expenseDetails?.Sum(x => x.Total).ToString() ?? "0" },

                };

                //Total WithoutTax
                Dictionary<string, string> modelTotalColumnDictionaryWithoutTax = new()
                {
                    { "FeeNoTax_Quantity_Total", feessDetails?.Where(x => x.TaxAmount == 0)?.Sum(x => x.ProductQuantity)?.ToString() ?? "0" },
                    { "FeeNoTax_UnitPrice_Total", feessDetails?.Where(x => x.TaxAmount == 0)?.Sum(x => x.ProductUnitPrice)?.ToString() ?? "0" },
                    { "FeeNoTax_NetAmount_Total", feessDetails?.Where(x => x.TaxAmount == 0)?.Sum(x => x.Amount).ToString() ?? "0" },
                    { "FeeNoTax_TaxAmount_Total", feessDetails?.Where(x => x.TaxAmount == 0)?.Sum(x => x.TaxAmount).ToString() ?? "0" },
                    { "FeeNoTax_GrossAmount_Total", feessDetails?.Where(x => x.TaxAmount == 0)?.Sum(x => x.Total).ToString() ?? "0" },

                    { "RecoverableCostsNoTax_UnitPrice_Total", disbursementDetails?.Where(x => x.TaxAmount == 0)?.Sum(x => x.ProductUnitPrice)?.ToString() ?? "0" },
                    { "RecoverableCostsNoTax_Quantity_Total", disbursementDetails?.Where(x => x.TaxAmount == 0)?.Sum(x => x.ProductQuantity)?.ToString() ?? "0" },
                    { "RecoverableCostsNoTax_NetAmount_Total", disbursementDetails?.Where(x => x.TaxAmount == 0)?.Sum(x => x.Amount).ToString() ?? "0" },
                    { "RecoverableCostsNoTax_TaxAmount_Total", disbursementDetails?.Where(x => x.TaxAmount == 0)?.Sum(x => x.TaxAmount).ToString() ?? "0" },
                    { "RecoverableCostsNoTax_GrossAmount_Total", disbursementDetails?.Where(x => x.TaxAmount == 0)?.Sum(x => x.Total).ToString() ?? "0" },

                    { "ExpenseNoTax_Quantity_Total", expenseDetails?.Where(x => x.TaxAmount == 0)?.Sum(x => x.ProductQuantity)?.ToString() ?? "0" },
                    { "ExpenseNoTax_UnitPrice_Total", expenseDetails?.Where(x => x.TaxAmount == 0)?.Sum(x => x.ProductUnitPrice)?.ToString() ?? "0" },
                    { "ExpenseNoTax_NetAmount_Total", expenseDetails?.Where(x => x.TaxAmount == 0)?.Sum(x => x.Amount).ToString() ?? "0" },
                    { "ExpenseNoTax_TaxAmount_Total", expenseDetails?.Where(x => x.TaxAmount == 0)?.Sum(x => x.TaxAmount).ToString() ?? "0" },
                    { "ExpenseNoTax_GrossAmount_Total", expenseDetails?.Where(x => x.TaxAmount == 0)?.Sum(x => x.Total).ToString() ?? "0" }
                };
                //Total WithTax
                Dictionary<string, string> modelTotalColumnDictionaryWithTax = new()
                {
                    { "FeeTax_UnitPrice_Total", feessDetails?.Where(x => x.TaxAmount > 0)?.Sum(x => x.ProductUnitPrice)?.ToString() ?? "0" },
                    { "FeeTax_NetAmount_Total", feessDetails?.Where(x => x.TaxAmount > 0)?.Sum(x => x.Amount).ToString() ?? "0" },
                    { "FeeTax_TaxAmount_Total", feessDetails?.Where(x => x.TaxAmount > 0)?.Sum(x => x.TaxAmount).ToString() ?? "0" },
                    { "FeeTax_Quantity_Total", feessDetails?.Where(x => x.TaxAmount > 0)?.Sum(x => x.ProductQuantity)?.ToString() ?? "0" },
                    { "FeeTax_GrossAmount_Total", feessDetails?.Where(x => x.TaxAmount > 0)?.Sum(x => x.Total).ToString() ?? "0" },

                    { "RecoverableCostsTax_Quantity_Total", disbursementDetails?.Where(x => x.TaxAmount > 0)?.Sum(x => x.ProductQuantity)?.ToString() ?? "0" },
                    { "RecoverableCostsTax_UnitPrice_Total", disbursementDetails?.Where(x => x.TaxAmount > 0)?.Sum(x => x.ProductUnitPrice)?.ToString() ?? "0" },
                    { "RecoverableCostsTax_NetAmount_Total", disbursementDetails?.Where(x => x.TaxAmount > 0)?.Sum(x => x.Amount).ToString() ?? "0" },
                    { "RecoverableCostsTax_TaxAmount_Total", disbursementDetails?.Where(x => x.TaxAmount > 0)?.Sum(x => x.TaxAmount).ToString() ?? "0" },
                    { "RecoverableCostsTax_GrossAmount_Total", disbursementDetails?.Where(x => x.TaxAmount > 0)?.Sum(x => x.Total).ToString() ?? "0" },

                    { "ExpenseTax_Quantity_Total", expenseDetails?.Where(x => x.TaxAmount > 0)?.Sum(x => x.ProductQuantity)?.ToString() ?? "0" },
                    { "ExpenseTax_UnitPrice_Total", expenseDetails?.Where(x => x.TaxAmount > 0)?.Sum(x => x.ProductUnitPrice)?.ToString() ?? "0" },
                    { "ExpenseTax_NetAmount_Total", expenseDetails?.Where(x => x.TaxAmount > 0)?.Sum(x => x.Amount).ToString() ?? "0" },
                    { "ExpenseTax_TaxAmount_Total", expenseDetails?.Where(x => x.TaxAmount > 0)?.Sum(x => x.TaxAmount).ToString() ?? "0" },
                    { "ExpenseTax_GrossAmount_Total", expenseDetails?.Where(x => x.TaxAmount > 0)?.Sum(x => x.Total).ToString() ?? "0" }
                };

                InsertTotalToGlobalData(modelTotalColumnDictionary, gDatas);
                InsertTotalToGlobalData(modelTotalColumnDictionaryWithoutTax, gDatas);
                InsertTotalToGlobalData(modelTotalColumnDictionaryWithTax, gDatas);
                #endregion
                //add transaction details into table row
                List<W.Table> tables = xmlElement.Descendants<W.Table>().ToList();
                //table loop
                foreach (var table in tables)
                {
                    W.TableRow tableTotalRow = null;
                    List<W.TableRow> tablePropertyRows = new();
                    //Get first first header row
                    W.TableRow tableFirstRow = table?.Elements<W.TableRow>()?.FirstOrDefault();
                    W.TableCell tableFirstRowCells = tableFirstRow?.Elements<W.TableCell>()?.FirstOrDefault();
                    string dictionaryFirstColumnHeader = tableFirstRowCells?.InnerText;
                    EnumCategoryTypeGroup invoiceType;
                    EnumTaxType taxType = EnumTaxType.All;

                    //table row loop
                    foreach (var tableRow in table.Elements<W.TableRow>())
                    {
                        List<KeyValuePair<string, string>> dictionaryColumnHeader = new();
                        Dictionary<int, W.TableCell> firstRowCellProp = new();
                        List<W.TableRow> tableRows = table.Elements<W.TableRow>().ToList();
                        List<W.TableCell> firstRowCells = tableRow.Elements<W.TableCell>().ToList();
                        int count = 1;
                        foreach (var cell in firstRowCells)
                        {
                            //heading row add into dictionary
                            dictionaryColumnHeader.Add(new KeyValuePair<string, string>(cell.InnerText, cell.InnerText.Replace(InvoiceMailMergeRegex.StartText, "").Replace(InvoiceMailMergeRegex.EndText, "")));
                            firstRowCellProp.Add(count, cell);
                            count++;
                        }
                        if (firstRowCells.Count > 0)
                        {
                            #region remove total row from table
                            if (dictionaryColumnHeader.Select(x => x.Value).Intersect(modelTotalColumnDictionary.Keys)?.Count() > 0)
                            {
                                tableTotalRow = tableRow.CloneNode(true) as W.TableRow;
                                tableRow.Remove();
                            }
                            #endregion

                            #region Feess
                            if (dictionaryColumnHeader.Select(x => x.Value).Intersect(modelFeessColumnDictionary.Keys)?.Count() > 0)
                            {
                                if (dictionaryFirstColumnHeader == $"[[{InvoiceTemplateHeaderData.Fee_Header}]]")
                                {
                                    //build table row
                                    MakeTableRowData(feessDetails, table, tableRow, firstRowCellProp, firstRowCells, dictionaryColumnHeader, modelFeessColumnDictionary);
                                }
                                else if (dictionaryFirstColumnHeader == $"[[{InvoiceTemplateHeaderData.Fee_Header_NoTax}]]")
                                {
                                    var feessDetailsWithOutTax = feessDetails?.Where(x => x.TaxAmount == 0)?.ToList();
                                    //build table row
                                    MakeTableRowData(feessDetailsWithOutTax, table, tableRow, firstRowCellProp, firstRowCells, dictionaryColumnHeader, modelFeessColumnDictionary);
                                }
                                else if (dictionaryFirstColumnHeader == $"[[{InvoiceTemplateHeaderData.Fee_Header_Tax}]]")
                                {
                                    var feessDetailsWithTax = feessDetails?.Where(x => x.TaxAmount > 0)?.ToList();
                                    //build table row
                                    MakeTableRowData(feessDetailsWithTax, table, tableRow, firstRowCellProp, firstRowCells, dictionaryColumnHeader, modelFeessColumnDictionary);
                                }
                                tablePropertyRows.Add(tableRow);
                            }
                            #endregion

                            #region Disbursement
                            else if (dictionaryColumnHeader.Select(x => x.Value).Intersect(modelDisbursementColumnDictionary.Keys)?.Count() > 0)
                            {
                                if (dictionaryFirstColumnHeader == $"[[{InvoiceTemplateHeaderData.Disbursements_Header}]]")
                                {
                                    //build table row
                                    MakeTableRowData(disbursementDetails, table, tableRow, firstRowCellProp, firstRowCells, dictionaryColumnHeader, modelDisbursementColumnDictionary);
                                }
                                else if (dictionaryFirstColumnHeader == $"[[{InvoiceTemplateHeaderData.Disbursements_Header_NoTax}]]")
                                {
                                    var disbursementDetailsWithOutTax = disbursementDetails?.Where(x => x.TaxAmount == 0)?.ToList();
                                    //build table row
                                    MakeTableRowData(disbursementDetailsWithOutTax, table, tableRow, firstRowCellProp, firstRowCells, dictionaryColumnHeader, modelDisbursementColumnDictionary);
                                }
                                else if (dictionaryFirstColumnHeader == $"[[{InvoiceTemplateHeaderData.Disbursements_Header_Tax}]]")
                                {
                                    var disbursementDetailsWithTax = disbursementDetails?.Where(x => x.TaxAmount > 0)?.ToList();
                                    //build table row
                                    MakeTableRowData(disbursementDetailsWithTax, table, tableRow, firstRowCellProp, firstRowCells, dictionaryColumnHeader, modelDisbursementColumnDictionary);
                                }
                                tablePropertyRows.Add(tableRow);
                            }
                            #endregion

                            #region Expense
                            else if (dictionaryColumnHeader.Select(x => x.Value).Intersect(modelExpenseColumnDictionary.Keys)?.Count() > 0)
                            {
                                if (dictionaryFirstColumnHeader == $"[[{InvoiceTemplateHeaderData.Expenses_Header}]]")
                                {
                                    //build table row
                                    MakeTableRowData(expenseDetails, table, tableRow, firstRowCellProp, firstRowCells, dictionaryColumnHeader, modelExpenseColumnDictionary);
                                }
                                else if (dictionaryFirstColumnHeader == $"[[{InvoiceTemplateHeaderData.Expenses_Header_NoTax}]]")
                                {
                                    var expenseDetailsWithOutTax = expenseDetails?.Where(x => x.TaxAmount == 0)?.ToList();
                                    //build table row
                                    MakeTableRowData(expenseDetailsWithOutTax, table, tableRow, firstRowCellProp, firstRowCells, dictionaryColumnHeader, modelExpenseColumnDictionary);
                                }
                                else if (dictionaryFirstColumnHeader == $"[[{InvoiceTemplateHeaderData.Expenses_Header_Tax}]]")
                                {
                                    var expenseDetailsWithTax = expenseDetails?.Where(x => x.TaxAmount > 0)?.ToList();
                                    //build table row
                                    MakeTableRowData(expenseDetailsWithTax, table, tableRow, firstRowCellProp, firstRowCells, dictionaryColumnHeader, modelExpenseColumnDictionary);
                                }
                                tablePropertyRows.Add(tableRow);
                            }
                            #endregion

                            #region Header Name

                            //Fees Header Name Change [[Fees_Header]] => Fees
                            else if (dictionaryColumnHeader.Select(x => x.Key).Any(x => x == $"[[{InvoiceTemplateHeaderData.Fee_Header}]]"))
                            {
                                invoiceType = EnumCategoryTypeGroup.Fees;
                                taxType = EnumTaxType.All;
                                MakeTableHeaderData(feessDetails?.Count ?? 0, table, firstRowCells, InvoiceTemplateHeaderData.Fee_Header, InvoiceTemplateHeaderValue.Fee_Header);
                            }
                            //[[Fees_Header_NoTax]]
                            else if (dictionaryColumnHeader.Select(x => x.Key).Any(x => x == $"[[{InvoiceTemplateHeaderData.Fee_Header_NoTax}]]"))
                            {
                                invoiceType = EnumCategoryTypeGroup.Fees;
                                taxType = EnumTaxType.WithoutTax;
                                int feessDetailsWithOutTaxCount = feessDetails?.Where(x => x.TaxAmount == 0)?.Count() ?? 0;
                                MakeTableHeaderData(feessDetailsWithOutTaxCount, table, firstRowCells, InvoiceTemplateHeaderData.Fee_Header_NoTax, InvoiceTemplateHeaderValue.Fee_Header_NoTax);
                            }
                            //[[Fees_Header_Tax]]
                            else if (dictionaryColumnHeader.Select(x => x.Key).Any(x => x == $"[[{InvoiceTemplateHeaderData.Fee_Header_Tax}]]"))
                            {
                                invoiceType = EnumCategoryTypeGroup.Fees;
                                taxType = EnumTaxType.WithTax;
                                int feessDetailsWithTaxCount = feessDetails?.Where(x => x.TaxAmount > 0)?.Count() ?? 0;
                                MakeTableHeaderData(feessDetailsWithTaxCount, table, firstRowCells, InvoiceTemplateHeaderData.Fee_Header_Tax, InvoiceTemplateHeaderValue.Fee_Header_Tax);
                            }

                            //Disbursements Header Name Change [[Disbursements_Header]] => Disbursements
                            else if (dictionaryColumnHeader.Select(x => x.Key).Any(x => x == $"[[{InvoiceTemplateHeaderData.Disbursements_Header}]]"))
                            {
                                invoiceType = EnumCategoryTypeGroup.Disbursement;
                                taxType = EnumTaxType.All;
                                MakeTableHeaderData(disbursementDetails?.Count ?? 0, table, firstRowCells, InvoiceTemplateHeaderData.Disbursements_Header, InvoiceTemplateHeaderValue.Disbursements_Header);
                            }
                            else if (dictionaryColumnHeader.Select(x => x.Key).Any(x => x == $"[[{InvoiceTemplateHeaderData.Disbursements_Header_NoTax}]]"))
                            {
                                invoiceType = EnumCategoryTypeGroup.Disbursement;
                                taxType = EnumTaxType.WithoutTax;
                                int disbursementDetailsWithOutTaxCount = disbursementDetails?.Where(x => x.TaxAmount == 0)?.Count() ?? 0;
                                MakeTableHeaderData(disbursementDetailsWithOutTaxCount, table, firstRowCells, InvoiceTemplateHeaderData.Disbursements_Header_NoTax, InvoiceTemplateHeaderValue.Disbursements_Header_NoTax);
                            }
                            else if (dictionaryColumnHeader.Select(x => x.Key).Any(x => x == $"[[{InvoiceTemplateHeaderData.Disbursements_Header_Tax}]]"))
                            {
                                invoiceType = EnumCategoryTypeGroup.Disbursement;
                                taxType = EnumTaxType.WithTax;
                                int disbursementDetailsWithTaxCount = disbursementDetails?.Where(x => x.TaxAmount > 0)?.Count() ?? 0;
                                MakeTableHeaderData(disbursementDetailsWithTaxCount, table, firstRowCells, InvoiceTemplateHeaderData.Disbursements_Header_Tax, InvoiceTemplateHeaderValue.Disbursements_Header_Tax);
                            }
                            //Expenses Header Name Change [[Expenses_Header]] => Expenses
                            else if (dictionaryColumnHeader.Select(x => x.Key).Any(x => x == $"[[{InvoiceTemplateHeaderData.Expenses_Header}]]"))
                            {
                                invoiceType = EnumCategoryTypeGroup.Expense;
                                taxType = EnumTaxType.All;
                                MakeTableHeaderData(expenseDetails?.Count ?? 0, table, firstRowCells, InvoiceTemplateHeaderData.Expenses_Header, InvoiceTemplateHeaderValue.Expenses_Header);
                            }
                            else if (dictionaryColumnHeader.Select(x => x.Key).Any(x => x == $"[[{InvoiceTemplateHeaderData.Expenses_Header_NoTax}]]"))
                            {
                                invoiceType = EnumCategoryTypeGroup.Expense;
                                taxType = EnumTaxType.WithoutTax;
                                int expenseDetailsWithOutTaxCount = expenseDetails?.Where(x => x.TaxAmount == 0)?.Count() ?? 0;
                                MakeTableHeaderData(expenseDetailsWithOutTaxCount, table, firstRowCells, InvoiceTemplateHeaderData.Expenses_Header_NoTax, InvoiceTemplateHeaderValue.Expenses_Header_NoTax);
                            }
                            else if (dictionaryColumnHeader.Select(x => x.Key).Any(x => x == $"[[{InvoiceTemplateHeaderData.Expenses_Header_Tax}]]"))
                            {
                                invoiceType = EnumCategoryTypeGroup.Expense;
                                taxType = EnumTaxType.WithTax;
                                int expenseDetailsWithTaxCount = expenseDetails?.Where(x => x.TaxAmount > 0)?.Count() ?? 0;
                                MakeTableHeaderData(expenseDetailsWithTaxCount, table, firstRowCells, InvoiceTemplateHeaderData.Expenses_Header_Tax, InvoiceTemplateHeaderValue.Expenses_Header_Tax);
                            }
                            #endregion
                        }
                    }
                    if (tablePropertyRows?.Count > 0)
                    {
                        foreach (var tablePropertyRow in tablePropertyRows)
                        {
                            tablePropertyRow.Remove();
                        }
                    }
                    //Insert Total row end of the table
                    if (tableTotalRow is not null)
                    {
                        foreach (var cell in tableTotalRow.Elements<W.TableCell>())
                        {
                            string cellText = cell.InnerText;
                            KeyValuePair<string, string> model = new();
                            if (taxType == EnumTaxType.All)//All
                            {
                                model = modelTotalColumnDictionary.FirstOrDefault(x => x.Key.ToLower() == cellText.Replace("[", "").Replace("]", "")?.Trim().ToString().ToLower());
                            }
                            else if (taxType == EnumTaxType.WithoutTax)//WithoutTax
                            {
                                model = modelTotalColumnDictionaryWithoutTax.FirstOrDefault(x => x.Key.ToLower() == cellText.Replace("[", "").Replace("]", "")?.Trim().ToString().ToLower());
                            }
                            else if (taxType == EnumTaxType.WithTax)//WithTax
                            {
                                model = modelTotalColumnDictionaryWithTax.FirstOrDefault(x => x.Key.ToLower() == cellText.Replace("[", "").Replace("]", "")?.Trim().ToString().ToLower());
                            }
                            var cellRunProp = cell.Elements<W.Paragraph>().FirstOrDefault().Elements<W.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true);
                            var cellParagraphProp = cell.Elements<W.Paragraph>().FirstOrDefault()?.ParagraphProperties?.CloneNode(true);
                            if (model.Value is not null)
                            {
                                cell.RemoveAllChildren<W.Paragraph>();
                                var newPara = new W.Paragraph(new W.Run(new W.Text(model.Value ?? cellText)) { RunProperties = (W.RunProperties)cellRunProp?.CloneNode(true) ?? new W.RunProperties() }) { ParagraphProperties = (W.ParagraphProperties)cellParagraphProp ?? new W.ParagraphProperties() };
                                cell.Append(newPara);
                            }
                        }
                        table.AppendChild(tableTotalRow);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        private static void MakeTableRowData(List<InvoiceDetails> invoiceDetails, W.Table table, W.TableRow tableRow, Dictionary<int, W.TableCell> firstRowCellProp, List<W.TableCell> firstRowCells, List<KeyValuePair<string, string>> dictionaryColumnHeader, Dictionary<string, string> modelColumnDictionary)
        {
            int count = 0;
            if (invoiceDetails?.Count > 0)
            {
                foreach (var item in invoiceDetails)
                {
                    var itemProp = item.GetType().GetProperties();
                    W.TableRow itemRow = new() { TableRowProperties = (W.TableRowProperties)tableRow?.TableRowProperties?.CloneNode(true) ?? new W.TableRowProperties() };
                    count = 1;
                    foreach (var dic in dictionaryColumnHeader)
                    {
                        var modelDic = modelColumnDictionary.FirstOrDefault(x => x.Key.ToLower() == dic.Value.Trim().ToString().ToLower());
                        var text = itemProp.FirstOrDefault(x => x.Name.ToLower() == modelDic.Value?.ToString()?.ToLower())?.GetValue(item)?.ToString();
                        var cellParagraphProp = firstRowCellProp.FirstOrDefault(x => x.Key == count).Value.Elements<W.Paragraph>().FirstOrDefault()?.ParagraphProperties?.CloneNode(true);
                        var cellRunProp = firstRowCellProp.FirstOrDefault(x => x.Key == count).Value.Elements<W.Paragraph>().FirstOrDefault().Elements<W.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true);
                        W.TableCell tableCell = new();
                        tableCell.Append(firstRowCellProp.FirstOrDefault(x => x.Key == count).Value.TableCellProperties?.CloneNode(true));
                        tableCell.Append(new W.Paragraph(new W.Run(new W.Text(text ?? "")) { RunProperties = (W.RunProperties)cellRunProp ?? new W.RunProperties() }) { ParagraphProperties = (W.ParagraphProperties)cellParagraphProp ?? new W.ParagraphProperties() });
                        itemRow.Append(tableCell);
                        count++;
                    }
                    table.AppendChild(itemRow);
                }
                foreach (var cell in firstRowCells)
                {
                    string cellText = cell.InnerText;
                    var model = modelColumnDictionary.FirstOrDefault(x => x.Key.ToLower() == cellText.Replace("[", "").Replace("]", "")?.Trim().ToString().ToLower());
                    var cellRunProp = cell.Elements<W.Paragraph>().FirstOrDefault().Elements<W.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true);
                    var cellParagraphProp = cell.Elements<W.Paragraph>().FirstOrDefault()?.ParagraphProperties?.CloneNode(true);
                    if (model.Value is not null)
                    {
                        cell.RemoveAllChildren<W.Paragraph>();
                        var newPara = new W.Paragraph(new W.Run(new W.Text(model.Value ?? cellText)) { RunProperties = (W.RunProperties)cellRunProp?.CloneNode(true) ?? new W.RunProperties() }) { ParagraphProperties = (W.ParagraphProperties)cellParagraphProp ?? new W.ParagraphProperties() };
                        cell.Append(newPara);
                    }
                }
            }
        }
        private static void MakeTableHeaderData(int invoiceDetailsCount, W.Table table, List<W.TableCell> firstRowCells, string invoiceTemplateHeaderData, string invoiceTemplateHeaderValue)
        {
            //invoiceTemplateHeaderValue = "";
            if (invoiceDetailsCount == 0)
            {
                table.Remove();
            }
            else
            {
                foreach (var cell in firstRowCells)
                {
                    string cellText = cell.InnerText;
                    if (cellText?.Trim() == $"[[{invoiceTemplateHeaderData}]]")
                    {
                        var runProp = cell.Elements<W.Paragraph>().FirstOrDefault().Elements<W.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true);
                        cell.RemoveAllChildren<W.Paragraph>();
                        var newPara = new W.Paragraph(new W.Run(new W.Text(invoiceTemplateHeaderValue)) { RunProperties = (W.RunProperties)runProp?.CloneNode(true) ?? new W.RunProperties() });
                        cell.Append(newPara);
                    }
                }
            }
        }

        private static void InsertTotalToGlobalData(Dictionary<string, string> modelInvoiceDictionary, List<GlobalDataVm> gDatas)
        {
            foreach (var item in modelInvoiceDictionary)
            {
                gDatas.Add(new GlobalDataVm() { Name = item.Key, InitialValue = item.Value });
            }
        }
    }
}
