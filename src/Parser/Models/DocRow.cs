using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocRow
    {
        [XmlElement("entry")]
        public List<DocEntry> Entries { get; } = new();
    }
}