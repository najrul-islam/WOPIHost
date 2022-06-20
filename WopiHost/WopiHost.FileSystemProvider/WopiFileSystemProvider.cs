﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.Extensions.Options;
using WopiHost.Abstractions;

namespace WopiHost.FileSystemProvider
{
    /// <summary>
    /// Provides files and folders based on a base64-ecoded paths.
    /// </summary>
    public class WopiFileSystemProvider : IWopiStorageProvider
    {
        public IOptionsSnapshot<WopiHostOptions> WopiHostOptions { get; }

        //private string ROOT_PATH = @".\";
        private readonly string ROOT_PATH = "";

        /// <summary>
        /// Reference to the root container.
        /// </summary>
        public IWopiFolder RootContainerPointer => new WopiFolder(ROOT_PATH, EncodeIdentifier(ROOT_PATH));

        protected string WopiRootPath => WopiHostOptions.Value.WopiRootPath;
        protected string WopiOtherDocsPath => WopiHostOptions.Value.WopiOtherDocsPath;

        /// <summary>
        /// Gets root path of the web application (e.g. IHostingEnvironment.WebRootPath for .NET Core apps)
        /// </summary>
        protected string WebRootPath => WopiHostOptions.Value.WebRootPath;

        protected string WopiAbsolutePath
        {
            get { return Path.IsPathRooted(WopiRootPath) ? WopiRootPath : Path.Combine(WebRootPath, WopiRootPath); }
        }
        protected string WopiAbsoluteOtherDocsPath
        {
            get { return Path.IsPathRooted(WopiOtherDocsPath) ? WopiOtherDocsPath : Path.Combine(WebRootPath, WopiOtherDocsPath); }
        }

        public WopiFileSystemProvider(IOptionsSnapshot<WopiHostOptions> wopiHostOptions)
        {
            WopiHostOptions = wopiHostOptions;
        }

        /// <summary>
        /// Gets a file using an identifier.
        /// </summary>
        /// <param name="identifier">A base64-encoded file path.</param>
        public IWopiFile GetWopiFile(string identifier, WopiDocsTypeEnum wopiDocsType = WopiDocsTypeEnum.WopiDocs)
        {
            string filePath = DecodeIdentifier(identifier);
            return new WopiFile(Path.Combine(wopiDocsType == WopiDocsTypeEnum.WopiDocs ? WopiAbsolutePath : WopiAbsoluteOtherDocsPath, filePath), identifier);
        }

        /// <summary>
        /// Gets a folder using an identifier.
        /// </summary>
        /// <param name="identifier">A base64-encoded folder path.</param>
        public IWopiFolder GetWopiContainer(string identifier = "")
        {
            string folderPath = DecodeIdentifier(identifier);
            return new WopiFolder(Path.Combine(WopiAbsolutePath, folderPath), identifier);
        }

        /// <summary>
        /// Gets all files in a folder.
        /// </summary>
        /// <param name="identifier">A base64-encoded folder path.</param>
        public List<IWopiFile> GetWopiFiles(string identifier = "")
        {
            string folderPath = DecodeIdentifier(identifier);
            List<IWopiFile> files = new List<IWopiFile>();
            foreach (string path in Directory.GetFiles(Path.Combine(WopiAbsolutePath, folderPath)))
            {
                string filePath = Path.Combine(folderPath, Path.GetFileName(path));
                string fileId = EncodeIdentifier(filePath);
                files.Add(GetWopiFile(fileId));
            }
            return files;
        }

        /// <summary>
        /// Gets all subfolders of a folder.
        /// </summary>
        /// <param name="identifier">A base64-encoded folder path.</param>
        public List<IWopiFolder> GetWopiContainers(string identifier = "")
        {
            string folderPath = DecodeIdentifier(identifier);
            List<IWopiFolder> folders = new List<IWopiFolder>();
            foreach (string directory in Directory.GetDirectories(Path.Combine(WopiAbsolutePath, folderPath)))
            {
                var subfolderPath = "." + directory.Remove(0, directory.LastIndexOf(Path.DirectorySeparatorChar));
                string folderId = EncodeIdentifier(subfolderPath);
                folders.Add(GetWopiContainer(folderId));
            }
            return folders;
        }

        private string DecodeIdentifier(string identifier)
        {
            var bytes = Convert.FromBase64String(identifier);
            return Encoding.UTF8.GetString(bytes);
        }

        private string EncodeIdentifier(string path)
        {
            var bytes = Encoding.UTF8.GetBytes(path);
            return Convert.ToBase64String(bytes);
        }
    }
}
