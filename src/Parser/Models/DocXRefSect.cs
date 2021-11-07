using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocXRefSect
    {
        [XmlAttribute("id")]
        public string Id { get; set; } = null!;

        [XmlElement("xreftitle")]
        public List<string> Title { get; } = new();

        [XmlElement("xrefdescription")]
        public Description Description { get; set; } = null!;
    }
}