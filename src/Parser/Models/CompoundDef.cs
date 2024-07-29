using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class CompoundDef : IDoxDescribable
    {
        [XmlAttribute("id")]
        public string Id { get; set; } = null!;

        [XmlAttribute("kind")]
        public CompoundKind Kind { get; set; }

        [XmlElement("compoundname")]
        public string? CompoundName { get; set; }

        [XmlElement("title")]
        public string Title { get; set; } = null!;

        [XmlElement("includes")]
        public List<Inc> Includes { get; } = [];

        [XmlElement("briefdescription")]
        public Description BriefDescription { get; set; } = null!;

        [XmlElement("detaileddescription")]
        public Description DetailedDescription { get; set; } = null!;

        [XmlElement("innergroup")]
        public List<Ref> InnerGroups { get; } = [];

        [XmlElement("innerfile")]
        public List<Ref> InnerFiles { get; } = [];

        [XmlElement("innerclass")]
        public List<Ref> InnerClasses { get; } = [];

        [XmlElement("sectiondef")]
        public List<SectionDef> Sections { get; } = [];
    }
}
