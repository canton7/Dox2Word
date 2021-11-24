using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class RefText
    {
        [XmlAttribute("refid")]
        public string RefId { get; set; } = null!;

        [XmlAttribute("kindref")]
        public string KindRef { get; set; } = null!;

        [XmlText]
        public string Name { get; set; } = null!;
    }
}
