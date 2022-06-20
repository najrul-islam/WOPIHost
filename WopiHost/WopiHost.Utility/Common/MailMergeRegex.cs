using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace WopiHost.Utility.Common
{
    public static class MailMergeRegex
    {
        /// <summary>
        /// Mailmerge starting text
        /// </summary>
        public static readonly string StartText = "{{";
        /// <summary>
        /// Mailmerge ending text
        /// </summary>
        public static readonly string EndText = "}}";
        /// <summary>
        /// Mailmerge full Regex pattern as string like {{.*?}}
        /// </summary>
        public static readonly string MailMergePattern = $"{StartText}.*?{EndText}";
        /// <summary>
        /// Mailmerge Regex
        /// </summary>
        public static readonly Regex Regex = new Regex(MailMergePattern);
        /// <summary>
        /// Mailmerge address Regex pattern like [\r\n]+
        /// </summary>
        public static readonly string MailMergeAddressPattern = "[\r\n]+";
        /// <summary>
        /// Symbol on template where content will insert
        /// </summary>
        public static readonly string MailMergeTemplateSymbol = "♥";
        /// <summary>
        /// Signature Start id
        /// </summary>
        public static readonly string MailMergeSignatureStartKey = "hxrSignatureStart";
        /// <summary>
        /// Signature End id
        /// </summary>
        public static readonly string MailMergeSignatureEndKey = "hxrSignatureEnd";
    }
    public static class InvoiceMailMergeRegex
    {
        /// <summary>
        /// InvoiceMailmerge starting text
        /// </summary>
        public static readonly string StartText = "[[";
        /// <summary>
        /// InvoiceMailmerge ending text
        /// </summary>
        public static readonly string EndText = "]]";
        /// <summary>
        /// InvoiceMailmerge full Regex pattern as string like [[.*?]]
        /// </summary>
        public static readonly string InvoiceMailMergePattern = @"\[\[.*?]]";
        /// <summary>
        /// InvoiceMailmerge Regex
        /// </summary>
        public static readonly Regex Regex = new Regex(InvoiceMailMergePattern);
    }
}
