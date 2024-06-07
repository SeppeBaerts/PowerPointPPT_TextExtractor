using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PowerPointTextExtractor.Models
{
    /// <summary>
    /// An atom record that specifies a Persist Object Directory.
    /// </summary>
    internal class PersistDirectoryAtom : Record
    {
        public List<PersistDirectoryEntry> PersistDirectoryEntries { get; set; }
        /// <summary>
        /// This will read the persist directory atom from the powerpoint document stream
        /// </summary>
        /// <param name="powerPointDocumentStream">The powerpointstream</param>
        public PersistDirectoryAtom(BinaryReader powerPointDocumentReader)
        {
            PersistDirectoryEntries = [];
            Read(powerPointDocumentReader);
        }
        public override void Read(BinaryReader reader)
        {
            base.Read(reader);
            var maxLength = reader.BaseStream.Position + Header.Size;
            do
            {
                PersistDirectoryEntries.Add(new(reader));
            } while (reader.BaseStream.Position < maxLength);
        }
    }
}
