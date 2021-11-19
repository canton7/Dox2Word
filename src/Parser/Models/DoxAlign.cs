using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public enum DoxAlign
    {
        [XmlEnum("left")]
        Left,

        [XmlEnum("right")]
        Right,

        [XmlEnum("center")]
        Center,
    }
}