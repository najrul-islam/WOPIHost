using System;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using WopiHost.Utility.ViewModel;
using System.Text.RegularExpressions;
using OpenXmlPowerTools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WopiHost.Utility.Common;
using System.IO;
using System.Threading.Tasks;

namespace WopiHost.Utility.XMLProcess
{
    public class ProcessBulkXMLMailMerge
    {
        public static async Task<NewLetterResponse> MergeDocuments(string templatePath, string contentPath, string signaturePath, string newFilePath, List<GlobalDataVm> gDatas, string skipFieldIds = "")
        {
            NewLetterResponse res = new();
            try
            {
                #region Merge Template and Content Using DocumentBuilder Process

                #region Get content insert position

                //int? inserFirstPositionIndex = 0;
                int? insertLastPositionIndex = 0;
                using (var saveDoc = WordprocessingDocument.Open(templatePath, false))
                {
                    //find content insert position
                    /*inserFirstPositionIndex = saveDoc.MainDocumentPart.Document.Body.Elements<Paragraph>()
                        .Select((x, i) => new { x.InnerText, Index = i })
                        .FirstOrDefault(x => x.InnerText == MailMergeRegex.MailMergeTemplateSymbol)
                        ?.Index;
                    saveDoc?.Close();*/
                    //find last content insert position
                    insertLastPositionIndex = saveDoc.MainDocumentPart.Document.Body.Elements<Paragraph>()
                        .Select((x, i) => new { x.InnerText, Index = i })
                        .LastOrDefault()
                        ?.Index;
                    saveDoc?.Close();
                }
                #endregion

                #region Merge content and template

                List<Source> sources1 = new List<Source>()
                        {
                            new Source(new WmlDocument(contentPath), 0, 0 , false),//keep content style
                            //new Source(new WmlDocument(templatePath), 0 , (int)inserFirstPositionIndex - 1, true),//take before ♥ from template
                            new Source(new WmlDocument(templatePath), 0 , 0, true),//take header part
                            new Source(new WmlDocument(contentPath), false),// insert full content
                            //new Source(new WmlDocument(templatePath), (int)inserFirstPositionIndex + 1, true),//take after ♥ from template
                            new Source(new WmlDocument(templatePath), (int)insertLastPositionIndex + 1, true),//take footer part
                        };
                DocumentBuilder.BuildDocument(sources1, newFilePath);
                #endregion

                #region Convert html signature to docx

                if (Path.GetExtension(signaturePath)?.ToLowerInvariant() == StaticData.Extensions.Html)
                {
                    //html to docx
                    string docxSignaturePath = $"{Path.GetDirectoryName(signaturePath)}\\{Guid.NewGuid()}.docx";
                    FileInfo signaturePathInfo = new(signaturePath);
                    ParseDocument.HtmlToDocx(signaturePathInfo, docxSignaturePath);
                    signaturePath = docxSignaturePath;
                }
                #endregion

                #region Signature Merge
                if (File.Exists(newFilePath) && File.Exists(signaturePath) && !string.IsNullOrEmpty(signaturePath) && Path.GetExtension(signaturePath)?.ToLowerInvariant() == StaticData.Extensions.Docx)
                {
                    string newSignaturePath = $"{Path.GetDirectoryName(signaturePath)}\\{Guid.NewGuid()}{Path.GetExtension(signaturePath)}";
                    //copy signature
                    File.Copy(signaturePath, newSignaturePath);
                    //int? insertSignaturePositionIndex;
                    using (var document = WordprocessingDocument.Open(newSignaturePath, true))
                    {
                        //insert custom id into start position
                        document.MainDocumentPart.Document.Body.Descendants<Paragraph>().FirstOrDefault()
                        ?.SetAttribute(new OpenXmlAttribute(MailMergeRegex.MailMergeSignatureStartKey, null, MailMergeRegex.MailMergeSignatureStartKey)
                        /*{
                            LocalName = MailMergeRegex.MailMergeSignatureStartKey,
                            Value = MailMergeRegex.MailMergeSignatureStartKey
                        }*/
                        );
                        //insert custom id into end position
                        document.MainDocumentPart.Document.Body.Descendants<Paragraph>().LastOrDefault()
                       ?.SetAttribute(new OpenXmlAttribute(MailMergeRegex.MailMergeSignatureEndKey, null, MailMergeRegex.MailMergeSignatureEndKey)
                       /*{
                           LocalName = MailMergeRegex.MailMergeSignatureStartKey,
                           Value = MailMergeRegex.MailMergeSignatureStartKey
                       }*/
                       );
                        document?.Close();
                    }

                    //merge content and template
                    List<Source> sources2 = new List<Source>()
                        {
                            new Source(new WmlDocument(newFilePath), 0, 0, false ),// keep header/footer
                            new Source(new WmlDocument(newSignaturePath), 0, 0 , false),//keep signature style
                            new Source(new WmlDocument(newFilePath), 0, false ),// (int)insertSignaturePositionIndex - 1, true),//take before [MailMergeSignatureSymbol] from signature template
                            new Source(new WmlDocument(newSignaturePath), false),//{ InsertId = "hxr-sign-start-id"},// insert full signature
                            //new Source(new WmlDocument(newFilePath), (int)insertSignaturePositionIndex + 1, true),//take after [MailMergeSignatureSymbol] from signature template
                        };
                    DocumentBuilder.BuildDocument(sources2, newFilePath);
                    //delete copied signature
                    File.Delete(newSignaturePath);
                }
                #endregion

                #region MailMerge Using Search and Replace
                if (gDatas is not null)
                {
                    /*string innerFullText = "";
                    using (var newDoc = WordprocessingDocument.Open(newFilePath, false))
                    {
                        innerFullText = newDoc.MainDocumentPart.Document.InnerText;
                    }
                    if (!string.IsNullOrEmpty(innerFullText))
                    {
                        Regex regex = new Regex(MailMergeRegex.MailMergePattern, RegexOptions.Compiled);
                        IEnumerable<string> matchList = regex.Matches(innerFullText).Cast<Match>().Select(x => x.Value).Distinct();
                        if (matchList?.Count() > 0)
                        {
                            //search and replace
                            MailMergeOpenXmlRegex(newFilePath, matchList, gDatas);
                            //highlight yellow color
                            MailMergeManualHighlight(newFilePath, regex, gDatas, res);
                        }
                    }*/
                    Regex regex = new Regex(MailMergeRegex.MailMergePattern, RegexOptions.Compiled);
                    MailMergeWithGlobalData(newFilePath, regex, MailMergeRegex.MailMergePattern, MailMergeRegex.StartText, MailMergeRegex.EndText, gDatas, res, skipFieldIds: skipFieldIds);
                }

                #endregion

                #endregion


                #region Merge Template and Content Using AltChunk Process

                /*
                 using (var saveDoc = WordprocessingDocument.Open(tempPath, false))
                {
                var mdp = saveDoc.MainDocumentPart;
                var templateParagraph = mdp.Document.Body.Elements<W.Paragraph>().Where(x => x.InnerText.Contains("♥")).FirstOrDefault();
                if (templateParagraphIndex == null)
                {
                    res.StatusCode = 403;
                    res.Message = "Your template format is not correct.";
                    return res;
                }*/
                /*var chunkId = $"AltChunk{Guid.NewGuid()}";
                var run = mdp.Document.Body.Elements<W.Run>().ToList();
                using (var contentStream = File.Open(contentPath, FileMode.Open))
                {
                    var chunk = mdp.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, chunkId);
                    chunk.FeedData(contentStream);
                }
                var alterChunk = new AltChunk { Id = chunkId };
                templateParagraph.InsertBeforeSelf(alterChunk);
                templateParagraph.Remove();
                saveDoc.Save();*/
                //ProcessInterOpDocument.RepairAndSave(savePath, savePath);
                #endregion

                #region Merge Template and Content Using Manula Process
                //if (!string.IsNullOrEmpty(contentPath))
                //{
                //    using (var contentStream = File.Open(contentPath, FileMode.Open))
                //    {
                //        using (var content = WordprocessingDocument.Open(contentStream, true))
                //        {
                //            //Copy whole content and image without last part <sectPr> XML
                //            foreach (var e in content.MainDocumentPart.Document.Body.Elements().Where(x => x.LocalName != "sectPr"))
                //            {
                //                var clonedElement = e.CloneNode(true);
                //                #region Copy Image from content docx
                //                clonedElement.Descendants<Blip>().ToList().ForEach(blip =>
                //                    {
                //                        var imgEmbed = saveDoc.CopyImage(blip.Embed, content);
                //                        blip.Embed = imgEmbed;
                //                    });
                //                clonedElement.Descendants<ImageData>().ToList().ForEach(imageData =>
                //                {
                //                    var imgRelationShipId = saveDoc.CopyImage(imageData.RelationshipId, content);
                //                    imageData.RelationshipId = imgRelationShipId;
                //                });
                //                #endregion
                //                #region Copy Hyperlink from content docx
                //                clonedElement.Descendants<Hyperlink>().ToList().ForEach(hyperLink =>
                //                {
                //                    var hyperLinkRelationShip = content.MainDocumentPart.HyperlinkRelationships.FirstOrDefault(x => x.Id == hyperLink.Id?.Value);
                //                    string newRelationshipsId = $"R{Guid.NewGuid().ToString().Replace("-", "")}";//New RelationshipsId
                //                    if (hyperLinkRelationShip != null)
                //                    {
                //                        saveDoc.MainDocumentPart.AddHyperlinkRelationship(hyperLinkRelationShip.Uri, true, newRelationshipsId);
                //                        hyperLink.Id = newRelationshipsId;//Replace Old RelationshipsId
                //                    }
                //                });
                //                #endregion
                //                /*clonedElement.Descendants<Comment>().ToList().ForEach(comment =>
                //                {
                //                    var contentComment = comment;
                //                });*/

                //                templateParagraph.InsertBeforeSelf(clonedElement);
                //            }

                //            //Copy Numbering format, bullet point from content etc
                //            CopyNumberAndBulltePoint(saveDoc, content);

                //            //Copy Style format from content etc
                //            CopyStyleDefinitionsPart(saveDoc, content);
                //            //Copy NameSpace from content
                //            /*foreach (var item in mdp.Document.NamespaceDeclarations)
                //            {
                //                if (!content.MainDocumentPart.Document.NamespaceDeclarations.Any(x => x.Key == item.Key))
                //                {
                //                    content.MainDocumentPart.Document.AddNamespaceDeclaration(item.Key, item.Value);
                //                }
                //            }*/

                //            //MergeTwoDocument(doc, content);
                //        }
                //    }
                //    templateParagraph.Remove();
                //    saveDoc.Save();
                //}
                //    if (gDatas != null)
                //    {
                //        Regex regex = new Regex(MailMergeRegex.MailMergePattern, RegexOptions.Compiled);
                //        MailMergeMainDocumentPart(mdp, regex, gDatas, res.UnMergeGlobalDataList);
                //    }
                //}
                #endregion
            }
            catch (Exception ex)
            {
                res.Message = ex.Message;
                if (ex.Message.Contains("Could not find file."))
                {
                    res.StatusCode = 404;
                }
                else
                {
                    res.StatusCode = 500;
                }
            }
            return await Task.FromResult(res);
        }

        #region Marge and Highlight
        public static void MailMergeWithGlobalData(string filePath, Regex regex, string regexPattern, string regexStartText, string regexEndText, List<GlobalDataVm> gDatas, NewLetterResponse res, bool highlightMissingData = true, string skipFieldIds = "")
        {
            if (gDatas != null)
            {
                string innerFullText = "";
                using (var newDoc = WordprocessingDocument.Open(filePath, false))
                {
                    innerFullText = newDoc.MainDocumentPart.Document.InnerText;
                    newDoc?.Close();
                }
                if (!string.IsNullOrEmpty(innerFullText))
                {
                    //Regex regex = new Regex(MailMergeRegex.MailMergePattern, RegexOptions.Compiled);
                    IEnumerable<string> matchList = regex.Matches(innerFullText).Cast<Match>().Select(x => x.Value).Distinct();
                    if (matchList?.Count() > 0)
                    {
                        //search and replace
                        MailMergeOpenXmlRegex(filePath, regexStartText, regexEndText, matchList, gDatas, skipFieldIds);
                        //highlight yellow color
                        if (highlightMissingData)
                        {
                            MailMergeManualHighlight(filePath, regex, regexPattern, regexStartText, regexEndText, gDatas, res, skipFieldIds);
                        }
                    }
                }
            }
        }
        static void MailMergeManualHighlight(string newFilePath, Regex regex, string regexPattern, string regexStartText, string regexEndText, List<GlobalDataVm> gDatas, NewLetterResponse res, string skipFieldIds = "")
        {
            using (var newDoc = WordprocessingDocument.Open(newFilePath, true))
            {
                var mdp = newDoc.MainDocumentPart;
                if (gDatas != null)
                {
                    MailMergeHighlightMainDocumentPart(mdp, regex, regexPattern, regexStartText, regexEndText, gDatas, res.UnMergeGlobalDataList, skipFieldIds);
                }
                newDoc.Save();
            }
        }
        static void MailMergeHighlightMainDocumentPart(MainDocumentPart mdp, Regex regex, string regexPattern, string regexStartText, string regexEndText, List<GlobalDataVm> gDatas, List<GlobalDataVm> unMergeGlobalData, string skipFieldIds = "")
        {
            if (mdp.HeaderParts?.Count() > 0)
            {
                var headers = mdp.HeaderParts.ToList();
                for (int i = 0; i < headers.Count; i++)
                {
                    MailMergeHighlight(headers[i].Header, regex, regexPattern, regexStartText, regexEndText, gDatas, unMergeGlobalData, skipFieldIds);
                }
            }
            MailMergeHighlight(mdp.Document, regex, regexPattern, regexStartText, regexEndText, gDatas, unMergeGlobalData, skipFieldIds);
            if (mdp.FooterParts?.Count() > 0)
            {
                var footers = mdp.FooterParts.ToList();
                for (int i = 0; i < footers.Count; i++)
                {
                    MailMergeHighlight(footers[i].Footer, regex, regexPattern, regexStartText, regexEndText, gDatas, unMergeGlobalData);
                }
            }
        }
        static void MailMergeHighlight(OpenXmlElement element, Regex regex, string regexPattern, string regexStartText, string regexEndText, List<GlobalDataVm> gDatas, List<GlobalDataVm> unMergeGlobalData, string skipFieldIds = "")
        {
            var paragraphs = element.Descendants<Paragraph>().Where(x => regex.Match(x.InnerText).Success).ToList();
            //var paragraphs = element.Descendants<W.Paragraph>().ToList();
            for (int i = 0; i < paragraphs.Count; i++)
            {
                //var match = regex.Match(paragraphs[i].InnerText);
                //if (match.Success)
                //{
                List<OpenXmlElement> allRun = new();
                //first run to get properties
                var firstRun = paragraphs[i].Descendants<Run>().FirstOrDefault();
                string para = paragraphs[i].InnerText.ToString();
                //var listMatches = Regex.Split(para, @"(\{.*?})");
                var listMatches = Regex.Split(para, $@"({regexPattern})");
                //var listMatches = regex.Split(para);
                foreach (var item in listMatches)
                {
                    var modifiedRun = new Run();
                    modifiedRun.AppendChild(firstRun.RunProperties?.CloneNode(true) ?? new RunProperties());
                    var tokenMatch = regex.Match(item);
                    if (tokenMatch.Success)
                    {
                        string token = tokenMatch.Value[regexStartText.Length..^regexEndText.Length];
                        //token = colorMatch.Value.[MailMergeRegex.StartText.Length..^MailMergeRegex.EndText.Length];
                        var gData = gDatas.FirstOrDefault(f => f.Name?.ToLowerInvariant() == System.Web.HttpUtility.HtmlDecode(token?.ToLowerInvariant()));
                        var rText = gData == null ? tokenMatch.Value : gData.InitialValue;

                        if (string.IsNullOrEmpty(gData?.InitialValue?.Trim()))//Highlight color(if not found in global data)
                        {
                            allRun.Add(new ProofError() { Type = ProofingErrorValues.SpellStart });//Spell check Start
                            modifiedRun.RunProperties.Highlight = new Highlight() { Val = HighlightColorValues.Yellow };
                            //colorRun.AppendChild(new W.RunProperties(new W.Highlight() { Val = HighlightColorValues.Yellow }));
                            modifiedRun.AppendChild(new Text(item));
                            //add to UnMergeGlobalData list
                            if (!unMergeGlobalData.Any(x => x.Name == token))
                            {
                                GlobalDataVm originalFieldGData = new();
                                if (Regex.IsMatch(token, @"(\w)(\.)(\w)+(\.)(\w)"))
                                {
                                    string originalField = token;
                                    var splitStr = token.Split('.');
                                    originalField = $"{splitStr[0]}.{splitStr[2]}";
                                    originalFieldGData = gDatas.FirstOrDefault(f => f.Name?.ToLowerInvariant() == System.Web.HttpUtility.HtmlDecode(originalField?.ToLowerInvariant()));
                                }
                                unMergeGlobalData.Add(gData ?? new GlobalDataVm() { Name = token, FieldId = originalFieldGData?.FieldId ?? 0, InputType = originalFieldGData?.InputType ?? "TextBox" });
                            }
                            allRun.Add(modifiedRun);
                            allRun.Add(new ProofError() { Type = ProofingErrorValues.SpellEnd });//Spell check End
                        }
                        else//Normar color 
                        {
                            //If gData is a address and contains NewLine(\n) then make it multiple line
                            if (Regex.Match(rText, MailMergeRegex.MailMergeAddressPattern).Success)
                            {
                                var multiAdrress = Regex.Split(rText, MailMergeRegex.MailMergeAddressPattern);
                                int count = 0;
                                foreach (var address in multiAdrress)
                                {
                                    count++;
                                    modifiedRun.AppendChild(new Text(address));
                                    if (count < multiAdrress.Length)
                                    {
                                        modifiedRun.AppendChild(new Break());
                                    }
                                }
                            }
                            else
                            {
                                modifiedRun.AppendChild(new Text(rText));
                            }
                            allRun.Add(modifiedRun);
                        }
                    }
                    else
                    {
                        modifiedRun.AppendChild(new Text(item));
                        allRun.Add(modifiedRun);
                    }
                }
                paragraphs[i].RemoveAllChildren<Run>();
                paragraphs[i].Append(allRun);
                //}
            }
        }
        #endregion

        #region Merge with Remove Highlight
        public static void MailMergeWithGlobalDataAndRemoveHighlight(string filePath, Regex regex, string regexPattern, string regexStartText, string regexEndText, List<GlobalDataVm> gDatas, NewLetterResponse res, string skipFieldIds = "")
        {
            if (gDatas != null)
            {
                string innerFullText = "";
                using (var newDoc = WordprocessingDocument.Open(filePath, false))
                {
                    innerFullText = newDoc.MainDocumentPart.Document.InnerText;
                    newDoc?.Close();
                }
                if (!string.IsNullOrEmpty(innerFullText))
                {
                    //Regex regex = new Regex(MailMergeRegex.MailMergePattern, RegexOptions.Compiled);
                    IEnumerable<string> matchList = regex.Matches(innerFullText).Cast<Match>().Select(x => x.Value).Distinct();
                    if (matchList?.Count() > 0)
                    {
                        //remove highlight yellow color
                        MailMergeManualHighlightRemove(filePath, regex, regexPattern, regexStartText, regexEndText, gDatas, res, skipFieldIds);
                        //search and replace
                        MailMergeOpenXmlRegex(filePath, regexStartText, regexEndText, matchList, gDatas, skipFieldIds);
                    }
                }
            }
        }
        static void MailMergeManualHighlightRemove(string newFilePath, Regex regex, string regexPattern, string regexStartText, string regexEndText, List<GlobalDataVm> gDatas, NewLetterResponse res, string skipFieldIds = "")
        {
            using (var newDoc = WordprocessingDocument.Open(newFilePath, true))
            {
                var mdp = newDoc.MainDocumentPart;
                if (gDatas != null)
                {
                    MailMergeHighlightRemoveMainDocumentPart(mdp, regex, regexPattern, regexStartText, regexEndText, gDatas, res.UnMergeGlobalDataList, skipFieldIds);
                }
                newDoc.Save();
            }
        }
        static void MailMergeHighlightRemoveMainDocumentPart(MainDocumentPart mdp, Regex regex, string regexPattern, string regexStartText, string regexEndText, List<GlobalDataVm> gDatas, List<GlobalDataVm> unMergeGlobalData, string skipFieldIds = "")
        {
            if (mdp.HeaderParts?.Count() > 0)
            {
                var headers = mdp.HeaderParts.ToList();
                for (int i = 0; i < headers.Count; i++)
                {
                    MailMergeRemoveHighlight(headers[i].Header, regex, regexPattern, regexStartText, regexEndText, gDatas, unMergeGlobalData, skipFieldIds);
                }
            }
            MailMergeRemoveHighlight(mdp.Document, regex, regexPattern, regexStartText, regexEndText, gDatas, unMergeGlobalData, skipFieldIds);
            if (mdp.FooterParts?.Count() > 0)
            {
                var footers = mdp.FooterParts.ToList();
                for (int i = 0; i < footers.Count; i++)
                {
                    MailMergeRemoveHighlight(footers[i].Footer, regex, regexPattern, regexStartText, regexEndText, gDatas, unMergeGlobalData, skipFieldIds);
                }
            }
        }
        //remove highlight
        static void MailMergeRemoveHighlight(OpenXmlElement element, Regex regex, string regexPattern, string regexStartText, string regexEndText, List<GlobalDataVm> gDatas, List<GlobalDataVm> unMergeGlobalData, string skipFieldIds = "")
        {
            //var skipFieldIdsInt = skipFieldIds.Split(",").Select(x => { _ = int.TryParse(x, out int result); return result; });
            var skipFieldIdsInt = skipFieldIds.Split(",");
            var paragraphs = element.Descendants<Paragraph>().Where(x => regex.Match(x.InnerText).Success).ToList();
            for (int i = 0; i < paragraphs.Count; i++)
            {
                List<OpenXmlElement> allRun = new();
                var firstRun = paragraphs[i].Descendants<Run>().FirstOrDefault();
                string para = paragraphs[i].InnerText.ToString();
                //var listMatches = Regex.Split(para, @"(\{.*?})");
                var listMatches = Regex.Split(para, $@"({regexPattern})");
                //var listMatches = regex.Split(para);
                foreach (var item in listMatches)
                {
                    var modifiedRun = new Run();
                    modifiedRun.AppendChild(firstRun.RunProperties?.CloneNode(true) ?? new RunProperties());
                    var colorMatch = regex.Match(item);
                    if (colorMatch.Success)
                    {
                        string token = colorMatch.Value[MailMergeRegex.StartText.Length..^MailMergeRegex.EndText.Length];
                        //token = colorMatch.Value.[MailMergeRegex.StartText.Length..^MailMergeRegex.EndText.Length];
                        var gData = gDatas.FirstOrDefault(f => f.Name?.ToLowerInvariant() == System.Web.HttpUtility.HtmlDecode(token?.ToLowerInvariant()));
                        var rText = string.IsNullOrEmpty(gData?.InitialValue?.Trim()) ? colorMatch.Value : gData.InitialValue;
                        #region match char at (like Client.PostCode/1)
                        if (Regex.IsMatch(token, @"/[0-9]+$"))
                        {
                            rText = GetCharAtValue(token, gDatas, out gData);
                            //gData = new GlobalDataVm() { InitialValue = rText };
                        }
                        #endregion
                        //bool? skipField = skipFieldIdsInt?.Any(x => x > 0 && gData?.FieldId > 0 && x == gData?.FieldId) ?? false;
                        bool? skipField = skipFieldIdsInt?.Any(x => !string.IsNullOrEmpty(x) && gData?.FieldId > 0 && x == gData?.Name) ?? false;

                        if (!string.IsNullOrEmpty(gData?.InitialValue?.Trim()) || skipField == true)// Remove Highlight color(if found in global data)
                        {
                            //If gData is a address and contains NewLine(\n) then make it multiple line
                            if (Regex.Match(rText, MailMergeRegex.MailMergeAddressPattern).Success)
                            {
                                var multiAdrress = Regex.Split(rText, MailMergeRegex.MailMergeAddressPattern);
                                int count = 0;
                                foreach (var address in multiAdrress)
                                {
                                    count++;
                                    modifiedRun.AppendChild(new Text(address));
                                    if (count < multiAdrress.Length)
                                    {
                                        modifiedRun.AppendChild(new Break());
                                    }
                                }
                            }
                            else
                            {
                                //allRun.Add(new ProofError() { Type = ProofingErrorValues.SpellStart });//Spell check Start
                                modifiedRun.RunProperties.Highlight = new Highlight() { Val = HighlightColorValues.None };
                                modifiedRun.AppendChild(new Text(skipField == true? "" : item)); //also skipField 
                                //allRun.Add(new ProofError() { Type = ProofingErrorValues.SpellEnd });//Spell check End
                            }
                            allRun.Add(modifiedRun);
                        }
                        else//Yellow color 
                        {
                            //allRun.Add(new ProofError() { Type = ProofingErrorValues.SpellStart });//Spell check Start
                            modifiedRun.RunProperties.Highlight = new Highlight() { Val = HighlightColorValues.Yellow };
                            modifiedRun.AppendChild(new Text(rText));
                            //allRun.Add(new ProofError() { Type = ProofingErrorValues.SpellEnd });//Spell check End
                            //add to UnMergeGlobalData list
                            if (!unMergeGlobalData.Any(x => x.Name == token))
                            {
                                unMergeGlobalData.Add(gData ?? new GlobalDataVm() { Name = token });
                            }
                            allRun.Add(modifiedRun);
                        }
                    }
                    else
                    {
                        modifiedRun.AppendChild(new Text(item));
                        //remove normal text highlight
                        if (!regex.Match(item).Success)
                        {
                            modifiedRun.RunProperties.Highlight = new Highlight() { Val = HighlightColorValues.None };
                        }
                        allRun.Add(modifiedRun);
                    }
                }
                paragraphs[i].RemoveAllChildren<Run>();
                paragraphs[i].Append(allRun);
                //}
            }
        }
        #endregion

        //merge missing data with search and replace method
        public static void MailMergeOpenXmlRegex(string filePath, string regexStartText, string regexEndText, IEnumerable<string> matchList, List<GlobalDataVm> gDatas, string skipFieldIds = "")
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                foreach (var item in matchList)
                {
                    string token = item.Substring(regexStartText.Length, item.Length - (regexStartText.Length + regexEndText.Length));
                    string tokenValue = gDatas.FirstOrDefault(x => x.Name?.ToLowerInvariant() == token?.ToLowerInvariant())?.InitialValue?.Trim();

                    #region match charAt (like Client.PostCode/1)
                    if (Regex.IsMatch(token, @"/[0-9]+$"))
                    {
                        tokenValue = GetCharAtValue(token, gDatas, out GlobalDataVm gData);
                    }
                    #endregion

                    //if value is not empty and is not address and is not multiline
                    if (!string.IsNullOrEmpty(tokenValue) && !Regex.Match(tokenValue, MailMergeRegex.MailMergeAddressPattern).Success)
                    {
                        Regex regex = new($"{Regex.Escape(item)}", RegexOptions.Compiled);
                        IEnumerable<XElement> content = wordDoc.MainDocumentPart.GetXDocument().Descendants(W.p).Where(x => x.Value.Contains(item));
                        OpenXmlRegex.Replace(content, regex, tokenValue, null);
                    }

                }
                wordDoc.MainDocumentPart.PutXDocument();
                wordDoc?.Close();
            }
        }
        public static void MailMergeOpenXmlRegexInvoice(string filePath, IEnumerable<string> matchList, List<GlobalDataVm> gDatas)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                foreach (var item in matchList)
                {
                    string token = item.Substring(InvoiceMailMergeRegex.StartText.Length, item.Length - (InvoiceMailMergeRegex.StartText.Length + InvoiceMailMergeRegex.EndText.Length));
                    string tokenValue = gDatas.FirstOrDefault(x => x.Name?.ToLowerInvariant() == token?.ToLowerInvariant())?.InitialValue?.Trim();
                    //if value is not empty and is not address and is not multiline
                    if (!string.IsNullOrEmpty(tokenValue) && !Regex.Match(tokenValue, MailMergeRegex.MailMergeAddressPattern).Success)
                    {
                        Regex regex = new($"{Regex.Escape(item)}", RegexOptions.Compiled);
                        IEnumerable<XElement> content = wordDoc.MainDocumentPart.GetXDocument().Descendants(W.p).Where(x => x.Value.Contains(item));
                        OpenXmlRegex.Replace(content, regex, tokenValue, null);
                    }
                }
                wordDoc.MainDocumentPart.PutXDocument();
            }
        }

        public static NewLetterResponse ReplaceSignatureOnExistingFile(string filePath, string signaturePath)
        {
            NewLetterResponse res = new();
            try
            {
                //covert to docx if html signature
                if (Path.GetExtension(signaturePath)?.ToLowerInvariant() == StaticData.Extensions.Html)
                {
                    //html to docx
                    string docxSignaturePath = $"{Path.GetDirectoryName(signaturePath)}\\{Guid.NewGuid()}{StaticData.Extensions.Docx}";
                    FileInfo signaturePathInfo = new(signaturePath);
                    ParseDocument.HtmlToDocx(signaturePathInfo, docxSignaturePath);
                    signaturePath = docxSignaturePath;
                }
                if (File.Exists(filePath) && File.Exists(signaturePath) && Path.GetExtension(signaturePath)?.ToLowerInvariant() == StaticData.Extensions.Docx)
                {
                    int? signatureStartPosition;
                    int? signatureEndPosition;
                    //int documentLastPosition;
                    using (var document = WordprocessingDocument.Open(filePath, false))
                    {
                        signatureStartPosition = document.MainDocumentPart.Document.Body.Elements<OpenXmlElement>()
                            .Select((x, i) => new { x.ExtendedAttributes?.FirstOrDefault(att => att.LocalName == MailMergeRegex.MailMergeSignatureStartKey).LocalName, Index = i })
                            .FirstOrDefault(x => x.LocalName == MailMergeRegex.MailMergeSignatureStartKey)
                            ?.Index;
                        signatureEndPosition = document.MainDocumentPart.Document.Body.Elements<OpenXmlElement>()
                            .Select((x, i) => new { x.ExtendedAttributes?.FirstOrDefault(att => att.LocalName == MailMergeRegex.MailMergeSignatureEndKey).LocalName, Index = i })
                            .FirstOrDefault(x => x.LocalName == MailMergeRegex.MailMergeSignatureEndKey)
                            ?.Index;
                        /*documentLastPosition = document.MainDocumentPart.Document.Body.Elements<Paragraph>()
                            .Select((x, Index) => (x, Index)).LastOrDefault().Index;*/
                        document?.Close();
                    }

                    string newSignaturePath = $"{Path.GetDirectoryName(signaturePath)}\\{Guid.NewGuid()}{Path.GetExtension(signaturePath)}";
                    //copy signature
                    File.Copy(signaturePath, newSignaturePath);
                    //int? insertSignaturePositionIndex;
                    using (var document = WordprocessingDocument.Open(newSignaturePath, true))
                    {
                        //insert custom id into start position
                        document.MainDocumentPart.Document.Body.Descendants<Paragraph>().FirstOrDefault()
                        ?.SetAttribute(new OpenXmlAttribute(MailMergeRegex.MailMergeSignatureStartKey, null, MailMergeRegex.MailMergeSignatureStartKey));
                        //insert custom id into end position
                        document.MainDocumentPart.Document.Body.Descendants<Paragraph>().LastOrDefault()
                       ?.SetAttribute(new OpenXmlAttribute(MailMergeRegex.MailMergeSignatureEndKey, null, MailMergeRegex.MailMergeSignatureEndKey));
                        document?.Close();
                    }

                    if (signatureStartPosition != null && signatureEndPosition != null)
                    {
                        List<Source> sources = new()
                        {
                            new Source(new WmlDocument(filePath), 0, 0, true ),// keep header/footer
                            new Source(new WmlDocument(newSignaturePath), 0, 0 , false),//keep signature style
                            new Source(new WmlDocument(filePath), 0 , (int)signatureStartPosition, true), //take before signature
                            new Source(new WmlDocument(newSignaturePath), false),// insert full signature
                            new Source(new WmlDocument(filePath), (int)signatureEndPosition + 1, true),
                            //new Source(new WmlDocument(filePath), documentLastPosition + 2, true),
                        };
                        DocumentBuilder.BuildDocument(sources, filePath);
                    }
                    else
                    {
                        List<Source> sources = new()
                        {
                            new Source(new WmlDocument(filePath), 0, 0, false ),// keep header/footer
                            new Source(new WmlDocument(newSignaturePath), 0, 0 , false),//keep signature style
                            new Source(new WmlDocument(filePath), 0, false ), //insert full docx
                            new Source(new WmlDocument(newSignaturePath), false),// insert signature at last position
                        };
                        DocumentBuilder.BuildDocument(sources, filePath);
                    }
                    File.Delete(newSignaturePath);
                }
            }
            catch (Exception ex)
            {
                res.Message = ex.Message;
                if (ex.Message.Contains("Could not find file."))
                {
                    res.StatusCode = 404;
                }
                else
                {
                    res.StatusCode = 500;
                }
                return res;
            }
            return res;
        }
        //(like Client.PostCode/1)
        private static string GetCharAtValue(string token, List<GlobalDataVm> gDatas, out GlobalDataVm gData)
        {
            string tokenValue = null;
            var charToken = token.Split('/')?.FirstOrDefault();
            var charAtToken = token.Split('/')?.LastOrDefault();
            _ = int.TryParse(charAtToken, out int charAtTokenNumb);
            gData = gDatas.FirstOrDefault(x => x.Name?.ToLowerInvariant() == charToken?.ToLowerInvariant());
            if (!string.IsNullOrEmpty(gData?.InitialValue?.Trim()) && charAtTokenNumb > 0 && charAtTokenNumb <= gData?.InitialValue?.Length)
            {
                tokenValue = gData?.InitialValue[charAtTokenNumb - 1].ToString();
            }
            return tokenValue;
        }
    }
}
