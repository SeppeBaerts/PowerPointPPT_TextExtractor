using System;
using System.Collections.Generic;
using System.Text;

namespace PowerPointTextExtractor.Models
{
    /// <summary>
    /// A string atom that is mostly used for hyperlinks. 
    /// </summary>
    internal class RTCString : Atom
    {
        public RTCString(BinaryReader reader, RecordHeader header)
        {
            Header = header;
            Read(reader);
        }
        public override string ToString()
        {
            string content = Encoding.Unicode.GetString(Data);
            return content;
        }
    }
}
