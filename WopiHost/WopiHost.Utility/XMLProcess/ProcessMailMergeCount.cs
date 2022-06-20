using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using WopiHost.Utility.Common;
using WopiHost.Utility.ViewModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
 
namespace WopiHost.Utility.XMLProcess
{
    public static class ProcessMailMergeCount
    {
        private static bool IsMatterSpecific = false;
        public static void DocumentMailMergeCount(string documentPath, List<GlobalDataVm> gDatas, List<GlobalDataVm> unMergeGlobalData, bool isMatterSpecific)
        {
            IsMatterSpecific = isMatterSpecific;
            try
            {
                using (var doc = WordprocessingDocument.Open(documentPath, false))
                {
                    var mdp = doc.MainDocumentPart;
                    if (gDatas != null && mdp != null)
                    {
                        Regex regex = new Regex(MailMergeRegex.MailMergePattern, RegexOptions.Compiled);
                        MailMergeMainDocumentPart(mdp, regex, gDatas, unMergeGlobalData);
                    }
                }
            }
            catch (Exception)
            {

            }
        }
        public static void MailMergeMainDocumentPart(MainDocumentPart mdp, Regex regex, List<GlobalDataVm> gDatas, List<GlobalDataVm> unMergeGlobalData)
        {
            if (mdp.HeaderParts?.Count() > 0)
            {
                var headers = mdp.HeaderParts.ToList();
                for (int i = 0; i < headers.Count; i++)
                {
                    MailMergeDocument(headers[i].Header, regex, gDatas, unMergeGlobalData);
                }
            }
            MailMergeDocument(mdp.Document, regex, gDatas, unMergeGlobalData);
            if (mdp.FooterParts?.Count() > 0)
            {
                var footers = mdp.FooterParts.ToList();
                for (int i = 0; i < footers.Count; i++)
                {
                    MailMergeDocument(footers[i].Footer, regex, gDatas, unMergeGlobalData);
                }
            }
        }
        public static void MailMergeDocument(OpenXmlElement element, Regex regex, List<GlobalDataVm> gDatas, List<GlobalDataVm> unMergeGlobalData)
        {
            var fullTextConetnt = element?.InnerText;
            if (!string.IsNullOrEmpty(fullTextConetnt))
            {
                //var match = regex.Match(paragraphs[i].InnerText);
                if (regex.Match(fullTextConetnt).Success)
                {
                    var courlyContent = Regex.Matches(fullTextConetnt, $@"({MailMergeRegex.MailMergePattern})").Cast<Match>().Select(m => m.Value).Distinct();
                    foreach (var item in courlyContent)
                    {
                        string token = item.Substring(MailMergeRegex.StartText.Length, item.Length - (MailMergeRegex.StartText.Length + MailMergeRegex.EndText.Length));//Get item without {}
                        GlobalDataVm gData = gDatas.FirstOrDefault(x => x.Name.ToLowerInvariant() == System.Web.HttpUtility.HtmlDecode(token?.ToLowerInvariant()));
                        try
                        {
                            if (IsMatterSpecific)
                            {
                                MatterStageActivityCount(token, gData, unMergeGlobalData);
                            }
                            else
                            {
                                UnMailMergeList(token, gData, unMergeGlobalData);
                            }
                        }
                        catch (Exception)
                        {
                            if (gData == null)
                            {
                                unMergeGlobalData.Add(new GlobalDataVm
                                {
                                    Name = token
                                });
                            }
                            else
                            {
                                unMergeGlobalData.Add(gData);
                            }
                        }
                    }
                }
            }
        }
        /// <summary>
        /// For matter stage activity Mailmerge count
        /// </summary>
        /// <param name="token"></param>
        /// <param name="gData"></param>
        /// <param name="unMergeGlobalData"></param>
        private static void MatterStageActivityCount(string token, GlobalDataVm gData, List<GlobalDataVm> unMergeGlobalData)
        {
            //List<InitialValueMultiple> initialValueMultiple = JsonConvert.DeserializeObject<List<InitialValueMultiple>>(gData?.InitialValueMultiple ?? "");
            if (string.IsNullOrEmpty(gData?.InitialValue))//if not found in global data
            {
                if (gData == null)
                {
                    unMergeGlobalData.Add(new GlobalDataVm
                    {
                        Name = token
                    });
                }
                else
                {
                    unMergeGlobalData.Add(gData);
                }
            }
            /*else if (initialValueMultiple?.Count > 1)
            {
                if (initialValueMultiple.Any(x => string.IsNullOrEmpty(x.InitialValue)))
                {
                    unMergeGlobalData.Add(gData);
                }
            }*/
        }

        /// <summary>
        /// Check Mailmerge Exist or Not (Used in Hoxro Cloud)
        /// </summary>
        /// <param name="token"></param>
        /// <param name="gData"></param>
        /// <param name="unMergeGlobalData"></param>
        private static void UnMailMergeList(string token, GlobalDataVm gData, List<GlobalDataVm> unMergeGlobalData)
        {
            if (gData == null)//if not found in global data
            {
                unMergeGlobalData.Add(new GlobalDataVm
                {
                    Name = token
                });
            }
        }
    }
}
