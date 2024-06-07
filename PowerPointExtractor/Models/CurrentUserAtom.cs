using System;
using System.Collections.Generic;
using System.Text;

namespace PowerPointTextExtractor.Models
{
    internal class CurrentUserAtom : Record
    {

        public uint Size { get; set; }
        /// <summary>
        /// 0xE391C05F = not encrypted, 0xF3D1C4DF = encrypted
        /// </summary>
        public uint HeaderToken { get; set; }
        /// <summary>
        /// Offset in bytes to the UserEditAtom of the most recent user. 
        /// </summary>
        public uint OffsetToCurrentEdit { get; set; }
        public ushort LenUserName { get; set; }
        public ushort DocFileVersion { get; set; }
        public byte MajorVersion { get; set; }
        public byte MinorVersion { get; set; }
        public ushort NotImportant { get; set; }
        public string AnsiUserName { get; set; }
        /// <summary>
        /// Specifies the release version of the application, it MUST be one of two values:
        /// 0x00000008, which means it contains one or more main master slides;
        /// 0x00000009, which means it contains more than one main master slide that should not be used. 
        /// </summary>
        public uint RelVersion { get; set; }
        public bool IsEncrypted => HeaderToken == 4090610911; //(0xF3D1C4DF)
        public override void Read(BinaryReader reader)
        {
            base.Read(reader);
            Size = reader.ReadUInt32();
            HeaderToken = reader.ReadUInt32(); //not encrypted = 3817979999 encrypted = 4090610911
            OffsetToCurrentEdit = reader.ReadUInt32();
            LenUserName = reader.ReadUInt16();
            DocFileVersion = reader.ReadUInt16();
            MajorVersion = reader.ReadByte();
            MinorVersion = reader.ReadByte();
            NotImportant = reader.ReadUInt16();
            AnsiUserName = Encoding.ASCII.GetString(reader.ReadBytes(LenUserName));
            RelVersion = reader.ReadUInt32();
        }
    }
}
