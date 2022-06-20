﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Claims;
using Cobalt;
using WopiHost.Abstractions;

namespace WopiHost.Cobalt
{
    public class CobaltProcessor : ICobaltProcessor
    {
        private CobaltFile GetCobaltFile(IWopiFile file, ClaimsPrincipal principal)
        {
            var disposal = new DisposalEscrow(file.Owner);
            CobaltFilePartitionConfig content = new CobaltFilePartitionConfig
            {
                IsNewFile = true,
                HostBlobStore = new TemporaryHostBlobStore(new TemporaryHostBlobStore.Config(), disposal, file.Identifier + @".Content"),
                cellSchemaIsGenericFda = true,
                CellStorageConfig = new CellStorageConfig(),
                Schema = CobaltFilePartition.Schema.ShreddedCobalt,
                PartitionId = FilePartitionId.Content
            };

            CobaltFilePartitionConfig coauth = new CobaltFilePartitionConfig
            {
                IsNewFile = true,
                HostBlobStore = new TemporaryHostBlobStore(new TemporaryHostBlobStore.Config(), disposal, file.Identifier + @".CoauthMetadata"),
                cellSchemaIsGenericFda = false,
                CellStorageConfig = new CellStorageConfig(),
                Schema = CobaltFilePartition.Schema.ShreddedCobalt,
                PartitionId = FilePartitionId.CoauthMetadata
            };

            CobaltFilePartitionConfig wacupdate = new CobaltFilePartitionConfig
            {
                IsNewFile = true,
                HostBlobStore = new TemporaryHostBlobStore(new TemporaryHostBlobStore.Config(), disposal, file.Identifier + @".WordWacUpdate"),
                cellSchemaIsGenericFda = false,
                CellStorageConfig = new CellStorageConfig(),
                Schema = CobaltFilePartition.Schema.ShreddedCobalt,
                PartitionId = FilePartitionId.WordWacUpdate
            };

            Dictionary<FilePartitionId, CobaltFilePartitionConfig> partitionConfigs = new Dictionary<FilePartitionId, CobaltFilePartitionConfig> { { FilePartitionId.Content, content }, { FilePartitionId.WordWacUpdate, wacupdate }, { FilePartitionId.CoauthMetadata, coauth } };


            var tempCobaltFile = new CobaltFile(disposal, partitionConfigs, new CobaltHostLockingStore(principal), null);

            if (file.Exists)
            {
                using (var stream = file.GetReadStream())
                {
                    var srcAtom = new AtomFromStream(stream);
                    Metrics o1;
                    tempCobaltFile.GetCobaltFilePartition(FilePartitionId.Content).SetStream(RootId.Default.Value, srcAtom, out o1);
                    tempCobaltFile.GetCobaltFilePartition(FilePartitionId.Content).GetStream(RootId.Default.Value).Flush();
                }
            }
            return tempCobaltFile;
        }



        public Stream GetFileStream(IWopiFile file, ClaimsPrincipal principal)
        {
            //TODO: use in filescontroller
            using (MemoryStream ms = new MemoryStream())
            {
                new GenericFda(GetCobaltFile(file, principal).CobaltEndpoint).GetContentStream().CopyTo(ms);
                return ms;
            }
        }


        public Action<Stream> ProcessCobalt(IWopiFile file, ClaimsPrincipal principal, byte[] newContent)
        {
            // Refactoring tip: there are more ways of initializing Atom
            AtomFromByteArray atomRequest = new AtomFromByteArray(newContent);
            RequestBatch requestBatch = new RequestBatch();

            object ctx;
            ProtocolVersion protocolVersion;

            requestBatch.DeserializeInputFromProtocol(atomRequest, out ctx, out protocolVersion);
            var cobaltFile = GetCobaltFile(file, principal);
            cobaltFile.CobaltEndpoint.ExecuteRequestBatch(requestBatch);

            if (requestBatch.Requests.Any(request => request is PutChangesRequest && request.PartitionId == FilePartitionId.Content))
            {
                using (var stream = file.GetWriteStream())
                {
                    new GenericFda(cobaltFile.CobaltEndpoint).GetContentStream().CopyTo(stream);
                }
            }
            var response = requestBatch.SerializeOutputToProtocol(protocolVersion);
            Action<Stream> copyToAction = s => { response.CopyTo(s); };
            return copyToAction;
        }
    }
}
