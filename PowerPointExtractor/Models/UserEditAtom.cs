using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PowerPointTextExtractor.Models
{
    /// <summary>
    /// This atom is used to store information about an edit to the document.
    /// </summary>
    internal class UserEditAtom : Record
    {
        public UserEditAtom(BinaryReader binaryReader)
        {
            Read(binaryReader);
        }
        /// <summary>
        ///  A reference to the last slide that was viewed.
        /// </summary>
        public uint LastSlideIdRef { get; set; }
        /// <summary>
        /// should be 0x0000 and MUST be ignored. 
        /// </summary>
        public ushort Version { get; set; }
        public byte MinorVersion { get; set; }
        public byte MajorVersion { get; set; }
        /// <summary>
        /// The offset in bytes from the beginning of the Powerpoint document stream to a 
        /// UserEdit Record for the previous user edit.
        /// </summary>
        public uint OffsetLastEdit { get; set; }
        /// <summary>
        /// The offset in bytes from the beginning of the Powerpoint document 
        /// stream to the PersistDirectoryAtom for THIS user edit.
        /// </summary>
        public uint OffsetPersistDirectory { get; set; }

        public uint DocPersistIdRef { get; set; }
        public uint PersistIdSeed { get; set; }
        public ushort LastView { get; set; }
        public PersistDirectoryAtom PersistDirectory { get; set; }
        public override void Read(BinaryReader reader)
        {
            base.Read(reader);
            LastSlideIdRef = reader.ReadUInt32();
            Version = reader.ReadUInt16();
            MinorVersion = reader.ReadByte();
            MajorVersion = reader.ReadByte();
            OffsetLastEdit = reader.ReadUInt32();
            OffsetPersistDirectory = reader.ReadUInt32();
            DocPersistIdRef = reader.ReadUInt32();
            PersistIdSeed = reader.ReadUInt32();
            LastView = reader.ReadUInt16();
            reader.ReadUInt16();//Skip 2 unused bytes
            reader.BaseStream.Position = OffsetPersistDirectory;
            PersistDirectory = new(reader);
        }
    }
}
