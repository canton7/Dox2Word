using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public enum DoxBool
    {
        [XmlEnum("yes")]
        Yes,

        [XmlEnum("no")]
        No,
    }
}
