using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{

    public class Compound
    {
        [XmlAttribute("refid")]
        public string RefId { get; set; } = null!;

        [XmlAttribute("kind")]
        public CompoundKind Kind { get; set; }

        [XmlElement("name")]
        public string Name { get; set; } = null!;
    }
}
