using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using WopiHost.Utility.ViewModel;

namespace WopiHost.Core.Models
{
    /*public class GlobalData
    {
        public int FieldId { get; set; }
        public string Name { get; set; }
        public string InitialValue { get; set; }
    }
    public class UnMergerdGlobalDataResult
    {
        public string url { get; set; }
        public List<GlobalData> UnMargeGlobalData { get; set; }
    }*/
    public static class Converter
    {
        //public static void WriteMailMerge(List<GlobalData> gDatas, string location)
        //{
        //    // Mail Merge
        //    try
        //    {
        //        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(location, true))
        //        {
        //            string docText = null;
        //            using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
        //            {
        //                docText = sr.ReadToEnd();
        //            }
        //            Regex regexText = new Regex("{.*?}");
        //            var fieldNames = new List<string>();
        //            string fieldName = string.Empty;
        //            MatchCollection matched = regexText.Matches(docText);
        //            //Field field;
        //            string token = string.Empty;
        //            foreach (Match match in matched)
        //            {
        //                token = match.Value.Substring(1, match.Value.Length - 2);
        //                Regex tokenRegex = new Regex("<.*?>");
        //                token = tokenRegex.Replace(token, "");
        //                var gData = gDatas.FirstOrDefault(f => f.Name == System.Web.HttpUtility.HtmlDecode(token));
        //                if (gData != null)
        //                {
        //                    docText = docText.Replace(match.Value, gData.InitialValue);
        //                }
        //                else
        //                {
        //                    //readText(match.Value);
        //                    //Check for image
        //                    if (!match.Value.Contains("-"))
        //                    {
        //                        docText = docText.Replace(match.Value, readText(match.Value));
        //                    }
        //                }
        //            }

        //            using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
        //            {
        //                sw.Write(docText);
        //            }
        //        }
        //    }
        //    catch (Exception exp)
        //    {
        //        throw exp;
        //    }
        //}
        public static void WriteMailMerge(List<GlobalData> gDatas, string location, string templatePath)
        {
            // Mail Merge
            try
            {
                Regex regex = new Regex("{.*?}", RegexOptions.Compiled);
                //MergeDocument(location, templatePath);
                if (gDatas != null && gDatas.Count() > 0)
                {
                    using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(location, true))
                    {

                        //grab the header parts and replace tags there
                        foreach (HeaderPart headerPart in wordDocument.MainDocumentPart.HeaderParts)
                        {
                            ReplaceParagraphParts(headerPart.Header, regex, gDatas);
                        }
                        //now do the document
                        ReplaceParagraphParts(wordDocument.MainDocumentPart.Document, regex, gDatas);
                        //now replace the footer parts
                        foreach (FooterPart footerPart in wordDocument.MainDocumentPart.FooterParts)
                        {
                            ReplaceParagraphParts(footerPart.Footer, regex, gDatas);
                        }
                        wordDocument.Save();
                    }
                    //ValidateWordDocument(location);
                }

            }
            catch (Exception exp)
            {
                throw exp;
            }
        }
        public static void ValidateWordDocument(string filepath)
        {
            using (WordprocessingDocument wordprocessingDocument =
            WordprocessingDocument.Open(filepath, true))
            {
                try
                {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    int count = 0;
                    foreach (ValidationErrorInfo error in
                        validator.Validate(wordprocessingDocument))
                    {
                        count++;
                        Console.WriteLine("Error " + count);
                        Console.WriteLine("Description: " + error.Description);
                        Console.WriteLine("ErrorType: " + error.ErrorType);
                        Console.WriteLine("Node: " + error.Node);
                        Console.WriteLine("Path: " + error.Path.XPath);
                        Console.WriteLine("Part: " + error.Part.Uri);
                        Console.WriteLine("-------------------------------------------");
                    }

                    Console.WriteLine("count={0}", count);
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                wordprocessingDocument.Close();
            }
        }
        private static void MergeDocument(string location, string templatePath)
        {
            Regex regexText = new Regex(@"\[content\]", RegexOptions.Compiled);
            //using (WordprocessingDocument wordTemplateDocument = WordprocessingDocument.Open(templatePath, true))
            //{
            //    using (WordprocessingDocument wordContentDocument = WordprocessingDocument.Open(location, true))
            //    {
            //        ReplaceParagraphParts(wordTemplateDocument.MainDocumentPart.Document, regexText, wordContentDocument.MainDocumentPart.Document);

            //    }
            //}
            //System.IO.File.Copy(templatePath, location, true);

            using (WordprocessingDocument wordContentDocument = WordprocessingDocument.Open(location, true))
            {
                //ReplaceParagraphParts(wordTemplateDocument.MainDocumentPart.Document, regexText, wordContentDocument.MainDocumentPart.Document);
                ReplaceParagraphParts(templatePath, regexText, wordContentDocument.MainDocumentPart.Document);
            }

            System.IO.File.Copy(templatePath, location, true);


        }


        private static string readText(string founditem)

        {

            Regex regexText = new Regex("<w:t>.*?</w:t>");
            MatchCollection matched = regexText.Matches(founditem);
            string token = string.Empty;
            token = string.Join(string.Empty, matched.Cast<Match>().Select(m => m.Value).ToArray());
            token = "<w:rPr><w:highlight w:val = 'lightGray' /></w:rPr>" + token.Replace("</w:t><w:t>", "");
            if (matched != null)
            {
                foreach (string str in matched.Cast<Match>().Select(m => m.Value).ToArray())
                {
                    founditem = founditem.Replace(str, string.Empty);
                }
            }
            Regex regexPrp = new Regex("<w:rPr>.*?</w:rPr>");
            MatchCollection matchedPrp = regexPrp.Matches(founditem);
            string matchPrp = string.Empty;
            if (matchedPrp != null)
            {
                var matchPrpArray = matchedPrp.Cast<Match>().Select(m => m.Value).ToArray();
                if (matchPrpArray != null && matchPrpArray.Count() > 0)
                {
                    matchPrp = matchPrpArray[0];
                }
            }


            int Place = founditem.IndexOf(matchPrp);
            if (Place <= 0)
            {
                //for no tag only {} assume that opening tag is there <w:t>
                founditem = "</w:t></w:r><w:r w:rsidRPr='6EDAA289' w:rsidR='6EDAA289'>" + "<w:rPr><w:highlight w:val = 'lightGray' /></w:rPr><w:t>" + founditem;
            }
            else
            {
                founditem = founditem.Remove(Place, matchPrp.Length).Insert(Place, token);
            }

            //founditem = founditem.Replace(matchPrp, token);
            return founditem;
        }



        private static void ReplaceParagraphParts(OpenXmlElement element, Regex regex, List<GlobalData> gDatas)
        {
            foreach (var paragraph in element.Descendants<Paragraph>())
            {
                Match match = regex.Match(paragraph.InnerText);

                if (match.Success)
                {
                    string token = string.Empty;
                    token = match.Value.Substring(1, match.Value.Length - 2);
                    Regex tokenRegex = new Regex("<.*?>");
                    token = tokenRegex.Replace(token, "");
                    var gData = gDatas.FirstOrDefault(f => f.Name == System.Web.HttpUtility.HtmlDecode(token));
                    var rText = gData == null ? match.Value : gData.InitialValue;
                    Run newRun = new Run();
                    newRun.AppendChild(new Text(paragraph.InnerText.Replace(match.Value, rText)));
                    //remove any child runs
                    paragraph.RemoveAllChildren<Run>();
                    //add the newly created run
                    paragraph.AppendChild(newRun);
                }
            }
        }
        private static void ReplaceParagraphParts(OpenXmlElement element, Regex regex, Document rText)
        {
            foreach (var paragraph in element.Descendants<Paragraph>())
            {
                Match match = regex.Match(paragraph.InnerText);

                if (match.Success)
                {
                    string token = string.Empty;
                    token = match.Value.Substring(1, match.Value.Length - 2);
                    Regex tokenRegex = new Regex("<.*?>");
                    token = tokenRegex.Replace(token, "");
                    Run newRun = new Run();
                    newRun.AppendChild(new Text(paragraph.InnerText.Replace(match.Value, rText.InnerText)));

                    //remove any child runs
                    paragraph.RemoveAllChildren<Run>();
                    //add the newly created run
                    paragraph.AppendChild(newRun);
                }
            }
        }
        public static void ReplaceParagraphParts(string location, Regex regex, Document rTextxml)
        {
            //regex = new Regex(@"\[.*?content.*\]", RegexOptions.Compiled);
            regex = new Regex(@"(?s)<w:p\b[^>]*>.*?</w:p>", RegexOptions.Compiled);
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(location, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                string matched = Regex.Matches(docText, @"(?s)<w:p\b[^>]*>.*?</w:p>").Cast<Match>().Select(x => x.Value).Where(z => z.Contains("content")).FirstOrDefault();
                if (!string.IsNullOrEmpty(matched))
                {
                    docText = docText.Replace(matched, rTextxml.Body.InnerXml);
                    using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(docText);
                    }
                }
            }
        }

        public static string highlightText(string Source, string Find, string Replace)
        {
            var styleCount = Source.Split(new[] { "<w:rPr>" }, StringSplitOptions.None);
            string result = string.Empty;
            int Place = Source.IndexOf(Find);
            if (Place == -1)
            {
                result = "<w:rPr><w:highlight w:val='lightGray'/></w:rPr><w:t>" + Source + "</w:t>";
            }
            else if (Place > 1)
            {

                result = Source.Remove(Place, Find.Length).Insert(Place, Replace);
            }

            return result;
        }
        public static string ReplaceFirstOccurrence(string Source, string Find, string Replace)
        {
            string result = string.Empty;
            int Place = Source.IndexOf(Find);
            if (Place == -1)
            {
                result = "<w:rPr><w:highlight w:val='lightGray'/></w:rPr><w:t>" + Source + "</w:t>";
            }
            else
            {

                result = Source.Remove(Place, Find.Length).Insert(Place, Replace);
            }

            return result;
        }
        public static string ToReplaceDots(this string item)
        {
            return item.Replace("\\.\\", "\\");
        }
        public static string AddSourceAddress(this string item, string sourceAddtess)
        {
            if (!string.IsNullOrEmpty(sourceAddtess))
            {
                string output = $"{item}\\{Path.GetFileName(sourceAddtess)}";
                return output;
            }
            return null;
        }
    }
}
