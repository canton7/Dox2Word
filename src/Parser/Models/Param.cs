using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class Param
    {
        [XmlElement("type")]
        public LinkedText? Type { get; set; }

        [XmlElement("declname")]
        public string? DeclName { get; set; }

        [XmlElement("defname")]
        public string? DefName { get; set; }
    }
}
