using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PowerPointTextExtractor.Models
{
    /// <summary>
    /// A subtype of record that contains raw bytes of data
    /// </summary>
    internal class Atom : Record
    {
        public byte[] Data;
        public override void Read(BinaryReader reader)
        {
            base.Read(reader);
            Data = reader.ReadBytes((int)Header.Size);
        }
        public override string ToString()
        {
            return Header.ToString();
        }
    }
}
