using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class Description
    {
        [XmlElement("title")]
        public string? Title { get; set; }

        [XmlElement("para")]
        public List<DocPara> Para { get; } = new();
    }
}
