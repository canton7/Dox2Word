using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public enum DoxVirtualKind
    {
        [XmlEnum("non-virtual")]
        NonVirtual,

        [XmlEnum("virtual")]
        Virtual,

        [XmlEnum("pure-virtual")]
        PureVirtual,
    }
}