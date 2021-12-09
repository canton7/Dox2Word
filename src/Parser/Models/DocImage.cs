using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocImage
    {
        [XmlAttribute("name")]
        public string? Name { get; set; }

        [XmlAttribute("caption")]
        public string? Caption { get; set; }

        [XmlAttribute("type")]
        public string? Type { get; set; }

        [XmlAttribute("inline")]
        public DoxBool Inline { get; set; }

        [XmlText]
        public string Contents { get; set; } = null!;
    }

    public class Dot : DocImage { }
    public class DotFile : DocImage { }
    public class Image : DocImage { }
}
