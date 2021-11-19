using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class Codeline
    {
        [XmlElement("highlight")]
        public List<Highlight> Highlights { get; } = new();
    }
}