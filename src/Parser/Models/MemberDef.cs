using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class MemberDef : IDoxDescribable
    {
        [XmlAttribute("kind")]
        public DoxMemberKind Kind { get; set; }

        [XmlAttribute("prot")]
        public DoxProtectionKind Protection { get; set; }

        [XmlAttribute("static")]
        public DoxBool IsStatic { get; set; }

        [XmlAttribute("const")]
        public DoxBool IsConst { get; set; }

        [XmlAttribute("explicit")]
        public DoxBool IsExplicit { get; set; }

        [XmlAttribute("inline")]
        public DoxBool IsInline { get; set; }

        [XmlAttribute("virt")]
        public DoxVirtualKind Virtual { get; set; }

        [XmlElement("name")]
        public string Name { get; set; } = null!;

        [XmlElement("bitfield")]
        public string? Bitfield { get; set; }

        [XmlElement("type")]
        public LinkedText? Type { get; set; }

        [XmlElement("definition")]
        public string? Definition { get; set; }

        [XmlElement("argsstring")]
        public string? ArgsString { get; set; }

        [XmlElement("param")]
        public List<Param> Params { get; } = new();

        [XmlElement("enumvalue")]
        public List<EnumValue> EnumValues { get; } = new();

        [XmlElement("briefdescription")]
        public Description? BriefDescription { get; set; }

        [XmlElement("detaileddescription")]
        public Description? DetailedDescription { get; set; }

        [XmlElement("initializer")]
        public LinkedText? Initializer { get; set; }
    }
}
