using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocSimpleSect
    {
        [XmlAttribute("kind")]
        public DoxSimpleSectKind Kind { get; set; }

        [XmlElement("para")]
        public DocPara? Para { get; set; }
    }
}