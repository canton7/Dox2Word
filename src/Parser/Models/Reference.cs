using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class Reference
    {
        [XmlAttribute("refid")]
        public string RefId { get; set; } = null!;

        [XmlText]
        public string Content { get; set; } = null!;
    }
}
