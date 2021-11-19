using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class Listing
    {
        [XmlElement("codeline")]
        public List<Codeline> Codelines { get; } = new();
    }
}