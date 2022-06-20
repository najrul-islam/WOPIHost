using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Azure.Storage.Blobs;
using Azure.Storage.Sas;
using Microsoft.Extensions.Options;
using WopiHost.Abstractions;

namespace WopiHost.FileSystemProvider
{

    public interface IAzureBlobStorage
    {
        Task<bool> UploadAsync(string storageName, string blobName, string filePath);
        Task<bool> UploadAsync(string storageName, string blobName, Stream stream);
        Task<bool> UploadExistingAsync(string storageName, string blobName, string filePath);
        Task DownloadAsync(string storageName, string blobName, string path);
        Task<bool> ExistsAsync(string storageName, string blobName);
        Task<string> GetViewUrlWithNoExpirationDate(string storageName, string blobName);
    }

    public class AzureBlobStorage : IAzureBlobStorage
    {
        private readonly string _connectionString;
        private readonly string _accountName;
        private readonly string _storageKey;
        private readonly string _accountUrl;

        public AzureBlobStorage(IOptionsSnapshot<WopiHostOptions> wopiHostOptions)
        {
            _accountName = wopiHostOptions.Value.StorageAccountName;
            _storageKey = wopiHostOptions.Value.StorageKey;
            _accountUrl = wopiHostOptions.Value.StorageAccountUrl;
            _connectionString = wopiHostOptions.Value.StorageConnectionString;
        }
        #region Upload
        public async Task<bool> UploadAsync(string storageName, string blobName, Stream stream)
        {
            try
            {
                //validate blob
                await ValidateBlobName(blobName);
                BlobClient blobClient = await GetBlobClient(storageName, blobName);
                stream.Position = 0;
                await blobClient.UploadAsync(stream, true);
            }
            catch (Exception ex)
            {
                return await Task.FromResult(false);
            }
            return await Task.FromResult(true);
        }
        public async Task<bool> UploadAsync(string storageName, string blobName, string filePath)
        {
            try
            {
                //validate blob
                await ValidateBlobName(blobName);
                BlobClient blobClient = await GetBlobClient(storageName, blobName);
                //Upload
                using (FileStream fileStream = File.OpenRead(filePath))
                {
                    #region Set/Update Metadata
                    /*if (await blobClient.ExistsAsync())
                    {
                    }
                    var existingMetadat = await blobClient.GetPropertiesAsync();
                    fileStream.Position = 0;
                    Dictionary<string, string> metaData = new();
                    metaData.Add("LastModifiedDate", DateTime.UtcNow.ToString());
                    await blobClient.SetMetadataAsync(metaData);
                    await blobClient.UploadAsync(fileStream, 
                        new BlobUploadOptions() {
                            Metadata = metaData
                        });*/
                    #endregion

                    await blobClient.UploadAsync(fileStream, true);
                }
            }
            catch (Exception ex)
            {
                return await Task.FromResult(false);
            }
            return await Task.FromResult(true);
        }
        public async Task<bool> UploadExistingAsync(string storageName, string blobName, string filePath)
        {
            try
            {
                //validate blob
                await ValidateBlobName(blobName);
                BlobClient blobClient = await GetBlobClient(storageName, blobName);
                //Upload
                using (FileStream fileStream = File.OpenRead(filePath))
                {
                    fileStream.Position = 0;
                    await blobClient.UploadAsync(fileStream, true);
                }
            }
            catch (Exception ex)
            {
                return await Task.FromResult(false);
            }
            return await Task.FromResult(true);
        }
        #endregion

        #region Download
        public async Task DownloadAsync(string storageName, string blobName, string path)
        {
            //Download
            //BlobContainerClient container = new BlobContainerClient(_connectionString, storageName);
            BlobClient blobClient = await GetBlobClient(storageName, blobName);
            if (await blobClient.ExistsAsync())
            {
                await blobClient.DownloadToAsync(path);
            }
            else
            {
                throw new Exception("Blob Not Found.");
            }
        }
        public async Task<MemoryStream> DownloadStreamAsync(string storageName, string blobName)
        {
            //BlobContainerClient container = new BlobContainerClient(_connectionString, storageName);
            BlobClient blobClient = await GetBlobClient(storageName, blobName);
            //Download
            MemoryStream stream = new MemoryStream();
            await blobClient.DownloadToAsync(stream);
            return await Task.FromResult(stream);
        }
        public async Task<byte[]> DownloadByteAsync(string storageName, string blobName)
        {
            try
            {
               // BlobContainerClient container = new BlobContainerClient(_connectionString, storageName);
                BlobClient blobClient = await GetBlobClient(storageName, blobName);
                //Download
                MemoryStream stream = new MemoryStream();
                await blobClient.DownloadToAsync(stream);
                stream.Position = 0;
                byte[] bytes = stream.ToArray();
                stream?.Close();
                return await Task.FromResult(bytes);
            }
            catch (Exception ex)
            {
                return new byte[0];
            }
        }
        public async Task<string> DownloadContentAsTextAsync(string storageName, string blobName)
        {
            try
            {
                BlobClient blobClient = await GetBlobClient(storageName, blobName);
                using (MemoryStream stream = new MemoryStream())
                {
                    await blobClient.DownloadToAsync(stream);
                    stream.Position = 0;
                    StreamReader streamReader = new StreamReader(stream);
                    string contentText = streamReader.ReadToEnd();
                    streamReader?.Close();
                    stream?.Close();
                    return await Task.FromResult(contentText);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion
        public async Task<bool> ExistsAsync(string storageName, string blobName)
        {
            BlobClient blobClient = await GetBlobClient(storageName, blobName);
            return await blobClient.ExistsAsync();
        }
        public async Task<string> GetViewUrlWithNoExpirationDate(string storageName, string blobName)
        {
            try
            {
                BlobClient blobClient = await GetBlobClient(storageName, blobName);
                var sasToken = blobClient.GenerateSasUri(BlobSasPermissions.Read, new DateTimeOffset(new DateTime(9999, 12, 31)));
                return await Task.FromResult(sasToken.ToString());
            }
            catch (Exception ex)
            {
                return await Task.FromResult(string.Empty);
            }
        }

        #region Private
        private async Task<BlobContainerClient> GetBlobContainerClient(string storageName)
        {
            return await Task.FromResult(new BlobContainerClient(_connectionString, storageName));
        }
        private async Task<BlobClient> GetBlobClient(string storageName, string blobName)
        {
            BlobContainerClient container = await GetBlobContainerClient(storageName);
            return container.GetBlobClient(blobName);
        }
        /// <summary>
        /// Validate blob Name and Extension
        /// </summary>
        /// <param name="blobName"></param>
        /// <returns></returns>
        private async Task<bool> ValidateBlobName(string blobName)
        {
            if (string.IsNullOrEmpty(Path.GetFileNameWithoutExtension(blobName)) ||
                string.IsNullOrEmpty(Path.GetExtension(blobName)))
            {
                throw new Exception("Invalid file name.");
            }
            return await Task.FromResult(true);
        }
        #endregion
    }
}
