using System;
using System.Collections.Generic;
using System.Text;

namespace PowerPointTextExtractor.Models
{
    internal class PersistOffsetEntry
    {
        /// <summary>
        /// The persist object identifier.
        /// </summary>
        public uint PersistId { get; set; }

        public uint RgPersistOffset { get; set; }
    }
}
