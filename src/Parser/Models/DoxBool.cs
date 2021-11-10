using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public enum DoxBool
    {
        [XmlEnum("no")]
        No,

        [XmlEnum("yes")]
        Yes,
    }
}
