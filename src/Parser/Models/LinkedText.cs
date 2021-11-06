using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class LinkedText
    {
        [XmlText(typeof(string))]
        [XmlElement("ref", typeof(RefText))]
        public List<object> Type { get; } = new();
    }
}
