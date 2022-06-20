using System.IO;
using System.Linq;
using WopiHost.Discovery.Enumerations;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WopiHost.Utility.XMLProcess
{
    public static class ProcessXmlDocument
    {
        /// <summary>
        /// Create new docx content/template based on user language
        /// </summary>
        /// <param name="savePath"></param>
        /// <param name="contentType"></param>
        /// <param name="UI_LLCC"></param>
        public static void CreateNewDocx(string savePath, string contentType, string UI_LLCC, string fontName)
        {
            var runPro = MakeNewRunProperties(UI_LLCC, fontName);
            // Create a document by supplying the savePath. 
            using (WordprocessingDocument newDocument = WordprocessingDocument.Create(savePath, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = newDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Document document = mainPart.Document;
                Body body = document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                if (contentType == WopiOperationTypeEnum.BlankTemplate.ToString())
                {
                    run.AppendChild(new Text("♥") { Space = SpaceProcessingModeValues.Preserve });
                    run.AppendChild(new Break());
                    Paragraph para2 = body.AppendChild(new Paragraph());
                    Run run2 = para.AppendChild(new Run());
                    run2.AppendChild(new Text("Blank Template"));
                    run2.PrependChild(runPro);
                }
                else
                {
                    run.AppendChild(new Text("Blank Content"));
                    Paragraph para2 = body.AppendChild(new Paragraph());
                    Run run2 = para.AppendChild(new Run());
                    run2.AppendChild(runPro);
                }
                newDocument.Close();
            }
        }
        /// <summary>
        /// Add Font and Language to existing document
        /// </summary>
        /// <param name="fileLocation"></param>
        /// <param name="UI_LLCC"></param>
        /// <param name="fontName"></param>
        static public void AddFontAndLanguage(string fileLocation, string UI_LLCC, string fontName)
        {
            if (File.Exists(fileLocation))
            {
                using (var document = WordprocessingDocument.Open(fileLocation, true))
                {
                    MainDocumentPart mainPart = document.MainDocumentPart;
                    var allRuns = mainPart.Document.Body.Descendants<Run>().ToList();
                    foreach (var item in allRuns)
                    {
                        //item.RunProperties = item.RunProperties ?? new W.RunProperties();
                        ModifyRunProperties(item.RunProperties = item.RunProperties ?? new RunProperties(), UI_LLCC, fontName);
                    }
                }
            }
        }
        static private OpenXmlElement MakeNewRunProperties(string UI_LLCC, string fontName)
        {
            fontName = fontName ?? "Calibri";
            UI_LLCC = UI_LLCC ?? "en-GB";
            RunProperties runPro = new RunProperties();
            runPro.Append(new Languages() { Val = UI_LLCC });
            if (fontName.Contains("Calibri Light (Headings)"))
            {
                runPro.Append(new RunFonts() { Ascii = fontName, HighAnsi = fontName, EastAsia = fontName, ComplexScript = fontName, AsciiTheme = ThemeFontValues.MajorAscii, HighAnsiTheme = ThemeFontValues.MajorAscii, EastAsiaTheme = ThemeFontValues.MajorAscii, ComplexScriptTheme = ThemeFontValues.MajorAscii });
            }
            else if (fontName.Contains("Calibri (Body)"))
            {
                runPro.Append(new RunFonts() { Ascii = fontName, HighAnsi = fontName, EastAsia = fontName, ComplexScript = fontName, AsciiTheme = ThemeFontValues.MinorAscii, HighAnsiTheme = ThemeFontValues.MinorAscii, EastAsiaTheme = ThemeFontValues.MinorAscii, ComplexScriptTheme = ThemeFontValues.MinorAscii });
            }
            else
            {
                runPro.Append(new RunFonts() { Ascii = fontName, HighAnsi = fontName, EastAsia = fontName, ComplexScript = fontName });
            }
            return runPro;
        }
        static private void ModifyRunProperties(RunProperties runPro, string UI_LLCC, string fontName)
        {
            fontName = fontName ?? "Calibri";
            UI_LLCC = UI_LLCC ?? "en-GB";
            runPro.Languages = new Languages() { Val = UI_LLCC };
            if (fontName.Contains("Calibri Light (Headings)"))
            {
                runPro.RunFonts = new RunFonts() { Ascii = fontName, HighAnsi = fontName, EastAsia = fontName, ComplexScript = fontName, AsciiTheme = ThemeFontValues.MajorAscii, HighAnsiTheme = ThemeFontValues.MajorAscii, EastAsiaTheme = ThemeFontValues.MajorAscii, ComplexScriptTheme = ThemeFontValues.MajorAscii };
            }
            else if (fontName.Contains("Calibri (Body)"))
            {
                runPro.RunFonts = new RunFonts() { Ascii = fontName, HighAnsi = fontName, EastAsia = fontName, ComplexScript = fontName, AsciiTheme = ThemeFontValues.MinorAscii, HighAnsiTheme = ThemeFontValues.MinorAscii, EastAsiaTheme = ThemeFontValues.MinorAscii, ComplexScriptTheme = ThemeFontValues.MinorAscii };
            }
            else
            {
                runPro.RunFonts = new RunFonts() { Ascii = fontName, HighAnsi = fontName, EastAsia = fontName, ComplexScript = fontName };
            }
        }
    }
}
