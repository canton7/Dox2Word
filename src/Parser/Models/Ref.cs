using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class Ref
    {
        [XmlAttribute("refid")]
        public string RefId { get; set; } = null!;

        [XmlText]
        public string Name { get; set; } = null!;
    }
}
