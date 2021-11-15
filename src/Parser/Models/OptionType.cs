using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public enum OptionType
    {
        [XmlEnum("int")]
        Int,

        [XmlEnum("bool")]
        Bool,

        [XmlEnum("string")]
        String,

        [XmlEnum("stringlist")]
        StringList,
    }
}