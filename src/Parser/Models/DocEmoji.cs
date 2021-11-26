using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocEmoji
    {
        [XmlAttribute("name")]
        public string Name { get; set; } = null!;

        [XmlAttribute("unicode")]
        public string Unicode { get; set; } = null!;
    }
}
