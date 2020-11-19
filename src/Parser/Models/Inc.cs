using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class Inc
    {
        [XmlAttribute("refid")]
        public string? RefId { get; set; }

        [XmlAttribute("local")]
        public DoxBool IsLocal { get; set; }

        [XmlText]
        public string Filename { get; set; } = null!;
    }
}