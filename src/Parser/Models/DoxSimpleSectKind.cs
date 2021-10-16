using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public enum DoxSimpleSectKind
    {
        [XmlEnum("see")]
        See,

        [XmlEnum("return")]
        Return,

        [XmlEnum("author")]
        Author,

        [XmlEnum("authors")]
        Authors,

        [XmlEnum("version")]
        Version,

        [XmlEnum("since")]
        Since,

        [XmlEnum("date")]
        Date,

        [XmlEnum("note")]
        Note,

        [XmlEnum("warning")]
        Warning,

        [XmlEnum("pre")]
        Pre,

        [XmlEnum("post")]
        Post,

        [XmlEnum("copyright")]
        Copyright,

        [XmlEnum("invariant")]
        Invariant,

        [XmlEnum("remark")]
        Remark,

        [XmlEnum("attention")]
        Attention,

        [XmlEnum("par")]
        Par,

        [XmlEnum("rcs")]
        Rcs,
    }
}