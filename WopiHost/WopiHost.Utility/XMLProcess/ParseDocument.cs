using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OpenXmlPowerTools;
using DocumentFormat.OpenXml;
using System.Drawing.Imaging;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WopiHost.Utility.XMLProcess
{
    public static class ParseDocument
    {
        #region Html To Docx
        public static void HtmlToDocx(string html, string saveDocxPath)
        {
            if (html?.Contains("<html>") != true)
            {
                html = $"<html><body>{html}</body></html>";
            }
            try
            {
                ConvertToDocx(html, saveDocxPath);
            }
            catch (Exception ex)
            {
                //throw ex;
            }

        }
        static public void HtmlToDocx(FileInfo htmlPath, string saveDocxPath)
        {
            string html = File.ReadAllText(htmlPath.FullName);
            if (html?.Contains("<html>") != true)
            {
                html = $"<html><body>{html}</body></html>";
                //File.WriteAllText(htmlPath, html);
            }
            try
            {
                ConvertToDocx(html, saveDocxPath);
            }
            catch (Exception ex)
            {
                //throw ex;
            }

        }
        private static void ConvertToDocx(string html, string destinationDir)
        {
            //bool s_ProduceAnnotatedHtml = true;
            //var sourceHtmlFi = new FileInfo(html);
            var sourceImageDi = new DirectoryInfo(destinationDir);

            var destDocx = new FileInfo(destinationDir);
            var destCss = new FileInfo(destinationDir.Replace(".docx", ".css"));
            //var annotatedHtml = new FileInfo(destinationDir.Replace(".docx", ".txt"));

            XElement xhtml = ReadAsXElement(html);

            string usedAuthorCss = HtmlToWmlConverter.CleanUpCss((string)xhtml.Descendants().FirstOrDefault(d => d.Name.LocalName.ToLower() == "style"));
            File.WriteAllText(destCss.FullName, usedAuthorCss);

            HtmlToWmlConverterSettings settings = HtmlToWmlConverter.GetDefaultSettings();
            // image references in HTML files contain the path to the subdir that contains the images, so base URI is the name of the directory
            // that contains the HTML files
            settings.BaseUriForImages = sourceImageDi.FullName;

            WmlDocument doc = HtmlToWmlConverter.ConvertHtmlToWml(defaultCss, usedAuthorCss, userCss, xhtml, settings, null, null);
            doc.SaveAs(destDocx.FullName);

            if (File.Exists(destCss?.FullName))
            {
                File.Delete(destCss?.FullName);
            }
            /*if (File.Exists(destCss?.FullName))
            {
                File.Delete(destCss?.FullName);
            }*/
        }
        private static XElement ReadAsXElement(string html)
        {
            XElement xhtml;
            try
            {
                xhtml = XElement.Parse(html);
            }
            //USE_HTMLAGILITYPACK
            catch (XmlException)
            {
                var sb = new StringBuilder();
                var stringWriter = new StringWriter(sb);
                var test = new HtmlAgilityPack.HtmlDocument();
                test.LoadHtml(html);
                test.OptionOutputAsXml = true;
                test.OptionCheckSyntax = true;
                test.OptionFixNestedTags = true;

                test.Save(stringWriter);
                sb.Replace("&amp;", "&");
                sb.Replace("&nbsp;", "\xA0");
                sb.Replace("&quot;", "\"");
                sb.Replace("&lt;", "~lt;");
                sb.Replace("&gt;", "~gt;");
                sb.Replace("&#", "~#");
                sb.Replace("&", "&amp;");
                sb.Replace("~lt;", "&lt;");
                sb.Replace("~gt;", "&gt;");
                sb.Replace("~#", "&#");
                xhtml = XElement.Parse(sb.ToString());
            }

            // HtmlToWmlConverter expects the HTML elements to be in no namespace, so convert all elements to no namespace.
            xhtml = (XElement)ConvertToNoNamespace(xhtml);
            return xhtml;
        }
        private static object ConvertToNoNamespace(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                return new XElement(element.Name.LocalName,
                    element.Attributes().Where(a => !a.IsNamespaceDeclaration),
                    element.Nodes().Select(n => ConvertToNoNamespace(n)));
            }
            return node;
        }

        static readonly string defaultCss =
       @"html, address,
blockquote,
body, dd, div,
dl, dt, fieldset, form,
frame, frameset,
h1, h2, h3, h4,
h5, h6, noframes,
ol, p, ul, center,
dir, hr, menu, pre { display: block; unicode-bidi: embed }
li { display: list-item }
head { display: none }
table { display: table }
tr { display: table-row }
thead { display: table-header-group }
tbody { display: table-row-group }
tfoot { display: table-footer-group }
col { display: table-column }
colgroup { display: table-column-group }
td, th { display: table-cell }
caption { display: table-caption }
th { font-weight: bolder; text-align: center }
caption { text-align: center }
body { margin: auto; }
h1 { font-size: 2em; margin: auto; }
h2 { font-size: 1.5em; margin: auto; }
h3 { font-size: 1.17em; margin: auto; }
h4, p,
blockquote, ul,
fieldset, form,
ol, dl, dir,
menu { margin: auto }
a { color: blue; }
h5 { font-size: .83em; margin: auto }
h6 { font-size: .75em; margin: auto }
h1, h2, h3, h4,
h5, h6, b,
strong { font-weight: bolder }
blockquote { margin-left: 40px; margin-right: 40px }
i, cite, em,
var, address { font-style: italic }
pre, tt, code,
kbd, samp { font-family: monospace }
pre { white-space: pre }
button, textarea,
input, select { display: inline-block }
big { font-size: 1.17em }
small, sub, sup { font-size: .83em }
sub { vertical-align: sub }
sup { vertical-align: super }
table { border-spacing: 2px; }
thead, tbody,
tfoot { vertical-align: middle }
td, th, tr { vertical-align: inherit }
s, strike, del { text-decoration: line-through }
hr { border: 1px inset }
ol, ul, dir,
menu, dd { margin-left: 40px }
ol { list-style-type: decimal }
ol ul, ul ol,
ul ul, ol ol { margin-top: 0; margin-bottom: 0 }
u, ins { text-decoration: underline }
br:before { content: ""\A""; white-space: pre-line }
center { text-align: center }
:link, :visited { text-decoration: underline }
:focus { outline: thin dotted invert }
/* Begin bidirectionality settings (do not change) */
BDO[DIR=""ltr""] { direction: ltr; unicode-bidi: bidi-override }
BDO[DIR=""rtl""] { direction: rtl; unicode-bidi: bidi-override }
*[DIR=""ltr""] { direction: ltr; unicode-bidi: embed }
*[DIR=""rtl""] { direction: rtl; unicode-bidi: embed }

";
        static readonly string userCss = @"";
        #endregion

        #region Docx To Html
        public static void DocxToHtml(string docxPath, string saveHtmlPath)
        {
            try
            {
                #region Docx to Html
                //FileStream docStream = new FileStream(docxPath, FileMode.Open, FileAccess.Read);
                ////Creates a new instance of WordDocument
                //WordDocument document = new WordDocument();
                //document.Open(docStream, FormatType.Docx);
                ////Creates an instance of memory stream
                //MemoryStream stream = new MemoryStream();
                ////Saves the Word document to MemoryStream
                //document.Save(stream, FormatType.Html);
                ////Closes the WordDocument instance
                //document.Close();
                //stream.Position = 0;
                //File.WriteAllBytes(saveHtmlPath, stream.ToArray());
                //stream?.Close();
                //fileStreamPath?.Close();
                #endregion
                //byte[] byteArray = File.ReadAllBytes(docxPath);
                //using (MemoryStream memoryStream = new MemoryStream())
                //{
                //    memoryStream.Write(byteArray, 0, byteArray.Length);
                //    using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
                //    {
                //        HtmlConverterSettings settings = new HtmlConverterSettings()
                //        {
                //            PageTitle = ""
                //        };
                //        XElement html = HtmlConverter.ConvertToHtml(doc, settings);
                //        File.WriteAllText(saveHtmlPath, html.ToStringNewLineOnAttributes());
                //        doc?.Close();
                //    }
                //    memoryStream?.Close();
                //}
                ConvertToHtml(docxPath, saveHtmlPath);
            }
            catch (Exception ex)
            {
                //throw ex;
            }

        }
        private static void ConvertToHtml(string file, string outputDirectory)
        {
            var fi = new FileInfo(file);
            //Console.WriteLine(fi.Name);
            byte[] byteArray = File.ReadAllBytes(fi.FullName);
            using (MemoryStream memoryStream = new())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(memoryStream, true))
                {
                    var destFileName = fi.FullName.Replace(".docx", ".html");
                    //var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
                    int imageCounter = 0;

                    var pageTitle = fi.FullName;
                    var part = wDoc.CoreFilePropertiesPart;
                    if (part != null)
                    {
                        pageTitle = (string)part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? fi.FullName;
                    }

                    // TODO: Determine max-width from size of content area.
                    WmlToHtmlConverterSettings settings = new()
                    {
                        //AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
                        AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
                        PageTitle = pageTitle,
                        FabricateCssClasses = true,
                        CssClassPrefix = "pt-",
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        ImageHandler = imageInfo =>
                        {
                            ++imageCounter;
                            string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                            ImageFormat imageFormat = null;
                            if (extension == "png")
                                imageFormat = ImageFormat.Png;
                            else if (extension == "gif")
                                imageFormat = ImageFormat.Gif;
                            else if (extension == "bmp")
                                imageFormat = ImageFormat.Bmp;
                            else if (extension == "jpeg")
                                imageFormat = ImageFormat.Jpeg;
                            else if (extension == "tiff")
                            {
                                // Convert tiff to gif.
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "x-wmf")
                            {
                                extension = "wmf";
                                imageFormat = ImageFormat.Wmf;
                            }

                            // If the image format isn't one that we expect, ignore it,
                            // and don't return markup for the link.
                            if (imageFormat == null)
                                return null;

                            string base64 = null;
                            try
                            {
                                using (MemoryStream ms = new())
                                {
                                    imageInfo.Bitmap.Save(ms, imageFormat);
                                    var ba = ms.ToArray();
                                    base64 = System.Convert.ToBase64String(ba);
                                }
                            }
                            catch (System.Runtime.InteropServices.ExternalException)
                            {
                                return null;
                            }

                            ImageFormat format = imageInfo.Bitmap.RawFormat;
                            ImageCodecInfo codec = ImageCodecInfo.GetImageDecoders().First(c => c.FormatID == format.Guid);
                            string mimeType = codec.MimeType;

                            string imageSource = string.Format("data:{0};base64,{1}", mimeType, base64);

                            XElement img = new(Xhtml.img,
                                new XAttribute(NoNamespace.src, imageSource),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }
                    };
                    XElement htmlElement = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

                    // Produce HTML document with <!DOCTYPE html > declaration to tell the browser
                    // we are using HTML5.
                    var html = new XDocument(
                        new XDocumentType("html", null, null, null),
                        htmlElement);

                    var htmlString = html.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
                    File.WriteAllText(destFileName, htmlString, Encoding.UTF8);
                }
            }
        }
        #endregion
    }
}
