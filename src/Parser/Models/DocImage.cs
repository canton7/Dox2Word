using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocImage
    {
        [XmlAttribute("name")]
        public string? Name { get; set; }

        [XmlText]
        public string Contents { get; set; } = null!;
    }

    public class Dot : DocImage { }
    public class DotFile : DocImage { }
}
