using System;
using WopiHost.Utility.Common;

namespace WopiHost.Core.Models
{
    public class FileStorageInfo
    {
        public FileStorageInfo() { }
        public FileStorageInfo(RequestProcessHeader headers, string id, string path)
        {
            UserId = headers.UserName;
            Id = id;
            Blob = headers.SourceBlob;
            Container = headers.StorageName;
            IsOverride = Convert.ToBoolean(headers.Override);
            LocalPath = path;
            IsUploaded = false;
            FileName = headers.FileName;
            StartTime = DateTime.Now;
            NewFileName = headers.NewDocumentName;
            BlobNew = headers.NewBlob;
            BlobTemplate = headers.SourceTemplateURL;
        }
        public FileStorageInfo(string userId, string id, string blob, string fileName, string container, bool isOverRide, string localPath, bool isUploaded, string newBlob, string newFileName)
        {
            UserId = userId;
            Id = id;
            Blob = blob;
            Container = container;
            IsOverride = isOverRide;
            LocalPath = localPath;
            IsUploaded = isUploaded;
            FileName = fileName;
            StartTime = DateTime.Now;
            NewFileName = newFileName;
            BlobNew = newBlob;
        }
        public FileStorageInfo(string userId, string id, string blob, string fileName, string container, bool isOverRide, string localPath, bool isUploaded, string newBlob, string newFileName, string templateBlob)
        {
            UserId = userId;
            Id = id;
            Blob = blob;
            Container = container;
            IsOverride = isOverRide;
            LocalPath = localPath;
            IsUploaded = isUploaded;
            FileName = fileName;
            StartTime = DateTime.Now;
            NewFileName = newFileName;
            BlobNew = newBlob;
            BlobTemplate = templateBlob;
        }
        public bool IsUpDateDone { get; set; }
        public bool IsCancel { get; set; }
        public string UserId { get; set; }
        public string Id { get; set; }
        public string Blob { get; set; }
        public string BlobNew { get; set; }
        public string BlobTemplate { get; set; }
        public string Container { get; set; }
        public bool IsOverride { get; set; }
        public string LocalPath { get; set; }
        public bool IsUploaded { get; set; }
        public DateTime StartTime { get; set; }
        public string LockId { get; set; }
        public string OldLockId { get; set; }
        public DateTime DateCreated { get; set; }
        public string FileName { get; set; }
        public string NewFileName { get; set; }
        public string UpUserId { get; set; }
        public string UpSearchFileName { get; set; }
        public bool IsEditing { get; set; }
        public int MatterDocumentId { get; set; }
        public int ParentId { get; set; }
        public int VersionNo { get; set; }
    }
}
