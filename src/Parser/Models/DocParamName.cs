using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocParamName
    {
        [XmlAttribute("direction")]
        public DoxParamDir Direction { get; set; }

        [XmlText(typeof(string))]
        [XmlElement("ref", typeof(RefText))]
        public List<object> Name { get; } = new();
    }
}
