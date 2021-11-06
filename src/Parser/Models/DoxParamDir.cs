using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public enum DoxParamDir
    {
        [XmlEnum("")]
        None,

        [XmlEnum("in")]
        In,

        [XmlEnum("out")]
        Out,

        [XmlEnum("inout")]
        InOut,
    }
}
