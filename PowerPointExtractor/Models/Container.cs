using System;
using System.Collections.Generic;
using System.Text;

namespace PowerPointTextExtractor.Models
{
    /// <summary>
    /// A contianer that contains a list of records
    /// </summary>
    internal class Container : Record
    {
        public List<Record> Children;

        public override string ToString()
        {
            return Header.ToString();
        }
    }
}
