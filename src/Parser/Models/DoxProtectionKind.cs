using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public enum DoxProtectionKind
    {
        [XmlEnum("public")]
        Public,

        [XmlEnum("protected")]
        Protected,

        [XmlEnum("private")]
        Private,

        [XmlEnum("package")]
        Package
    }
}