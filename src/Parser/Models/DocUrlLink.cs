using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocUrlLink : DocTitleCmdGroup
    {
        [XmlAttribute("url")]
        public string Url { get; set; } = null!;
    }
}
