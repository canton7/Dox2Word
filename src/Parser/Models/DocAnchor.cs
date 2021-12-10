using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocAnchor
    {
        [XmlAttribute("id")]
        public string Id { get; set; } = null!;
    }
}
