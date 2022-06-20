using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WopiHost.Utility.StaticData
{
    public static class Extensions
    {
        public const string Docx = ".docx";
        public const string Pdf = ".pdf";
        public const string Xlsx = ".xlsx";
        public const string Pptx = ".pptx";
        public const string Html = ".html";
        public const string Htm = ".htm";
        public const string Mov = ".mov";
        public const string Jpeg = ".jpeg";
        public const string Jpg = ".jpg";
        public const string Ico = ".ico";
        public const string Avi = ".avi";
        public const string Bmp = ".bmp";
        public const string Msg = ".msg";
        public const string eml = ".eml";
        public const string Mp3 = ".mp3";
        public const string Mp4 = ".mp4";
        public const string Tif = ".tif";
        public const string Tiff = ".tiff";
        public const string Png = ".png";
        public const string Wav = ".wav";
        public const string Gif = ".gif";
        public const string Xml = ".xml";
        public const string Txt = ".txt";
        public const string Csv = ".csv";
        public const string Rtf = ".rtf";

        public const string M4a = ".m4a";
        public const string Doc = ".doc";
        public const string Ppt = ".ppt";
        public const string Xls = ".xls";

    }
    public class DictionaryExtension
    {
        static readonly Dictionary<string, int> getExtensionId = new Dictionary<string, int>
                        {
                            {".docx", 1 },
                            {".pdf", 2 },
                            {".xlsx", 3 },
                            {".pptx", 4 },
                            {".html", 5 },
                            {".mov", 6 },
                            {".jpeg", 7 },
                            {".jpg", 8 },
                            {".avi", 9 },
                            {".bmp", 10 },
                            {".msg", 11 },
                            {".eml", 12 },
                            {".mp3", 13 },
                            {".mp4", 14 },
                            {".tif", 15 },
                            {".tiff", 16 },
                            {".png", 17 },
                            {".wav", 18 },
                            {".gif", 19 },
                            {".xml", 20 },
                            {".txt", 21 },
                            {".xls", 22 },
                            {".csv", 23 },
                            {".htm", 24 },
                            {".ico", 25 },
                            {".svg", 26 },
                            {".m4a", 27 },
                            {".doc", 28 },
                            {".mkv", 29 }
                        };
        /// <summary>
        /// the extensionName is including the period "."
        /// </summary>
        /// <param name="extensionName"></param>
        /// <returns></returns>
        public static int GetExtensionId(string extensionName)
        {
            getExtensionId.TryGetValue(extensionName?.ToLowerInvariant(), out int result);
            return result;
        }
    }
}
