using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class EnumValue : IDoxDescribable
    {
        [XmlAttribute("id")]
        public string Id { get; set; } = null!;

        [XmlElement("name")]
        public string Name { get; set; } = null!;

        [XmlElement("initializer")]
        public LinkedText? Initializer { get; set; }

        [XmlElement("briefdescription")]
        public Description? BriefDescription { get; set; }

        [XmlElement("detaileddescription")]
        public Description? DetailedDescription { get; set; }
    }
}