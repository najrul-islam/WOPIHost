using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WopiHost.Utility.ExtensionMetod
{
    public static class GlobalExtensionMethod
    {
        /// <summary>
        /// Get file extension name including dot (like .docx)
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string GetExtension(this string fileName)
        {
            return Path.GetExtension(fileName)?.ToLower();
        }

        /// <summary>
        /// Get file name including extension
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string GetFileName(this string fileName)
        {
            return Path.GetFileName(fileName)?.ToLower();
        }
    }
}
