using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PowerPointTextExtractor.Models
{
    /// <summary>
    /// A record that contains the persist object identifier and the offset of the persist object.
    /// </summary>
    internal class PersistDirectoryEntry : Record
    {
        public uint PersistStartId { get; set; }
        public uint PersistCount { get; set; }
        public List<PersistOffsetEntry> PersistOffsets { get; set; }

        public PersistDirectoryEntry(BinaryReader reader)
        {
            PersistOffsets = [];
            Read(reader);
        }
        public override void Read(BinaryReader reader)
        {
            var persistIdAndCount = reader.ReadUInt32();
            PersistStartId = persistIdAndCount & 0x000FFFFFU; // First 20 bit of field "persistIdAndCount"
            PersistCount = persistIdAndCount >> 20 & 0x00000FFFU;  // Last 12 bit of field "persistIdAndCount"
            uint startId = PersistStartId;
            for (uint i = 0; i < PersistCount; i++)
            {
                var persistOffset = new PersistOffsetEntry()
                {
                    PersistId = startId++,
                    RgPersistOffset = reader.ReadUInt32()
                };
                PersistOffsets.Add(persistOffset);
            }
        }
    }
}
