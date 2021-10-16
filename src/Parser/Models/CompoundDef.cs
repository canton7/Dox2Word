using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class CompoundDef
    {
        [XmlAttribute("id")]
        public string Id { get; set; } = null!;

        [XmlAttribute("kind")]
        public CompoundKind Kind { get; set; }

        [XmlElement("title")]
        public string Title { get; set; } = null!;

        [XmlElement("briefdescription")]
        public Description BriefDescription { get; set; } = null!;

        [XmlElement("detaileddescription")]
        public Description DetailedDescription { get; set; } = null!;

        [XmlElement("innergroup")]
        public List<Ref> InnerGroups { get; } = new();

        [XmlElement("innerfile")]
        public List<Ref> InnerFiles { get; } = new();

        [XmlElement("innerclass")]
        public List<Ref> InnerClasses { get; } = new();

        [XmlElement("sectiondef")]
        public List<SectionDef> Sections { get; } = new();
    }
}
