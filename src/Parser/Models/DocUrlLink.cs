using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocUrlLink : DocPara
    {
        [XmlAttribute("url")]
        public string Url { get; set; } = null!;
    }
}
