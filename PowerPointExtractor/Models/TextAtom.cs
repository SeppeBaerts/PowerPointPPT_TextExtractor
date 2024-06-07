using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;

[assembly: InternalsVisibleTo("PPTExtractor")]
namespace PowerPointTextExtractor.Models
{
    /// <summary>
    /// a text atom that contains all the text in the powerpoint 
    /// </summary>
    /// 
    internal class TextAtom : Atom
    {
        internal TextAtom(byte[] data) { Data = data; }
        public TextAtom(BinaryReader reader, RecordHeader header)
        {
            Header = header;
            Read(reader);
        }
        public override string ToString()
        {
            bool isInUnicode = Data.Contains((byte)0x00); //Check if the string already is in unicode (this happens when there are non UTF-8 characters);
            if (isInUnicode)
                return Encoding.Unicode.GetString(Data);
            //If it's not in unicode => we're going to add the 0x00 bytes to make it unicode
            List<byte> newBytes = [];
            foreach (var b in Data)
            {
                newBytes.Add(b);
                newBytes.Add(0x00);
            }
            string content = Encoding.Unicode.GetString([.. newBytes]);
            return content.Replace("\r", "\n");
        }
    }
}
