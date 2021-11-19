using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public enum CompoundKind
    {
        [XmlEnum("class")]
        Class,

        [XmlEnum("struct")]
        Struct,

        [XmlEnum("union")]
        Union,

        [XmlEnum("interface")]
        Interface,

        [XmlEnum("protocol")]
        Protocol,

        [XmlEnum("category")]
        Category,

        [XmlEnum("exception")]
        Exception,

        [XmlEnum("service")]
        Service,

        [XmlEnum("singleton")]
        Singleton,

        [XmlEnum("module")]
        Module,

        [XmlEnum("type")]
        Type,

        [XmlEnum("file")]
        File,

        [XmlEnum("namespace")]
        Namespace,

        [XmlEnum("group")]
        Group,

        [XmlEnum("page")]
        Page,

        [XmlEnum("example")]
        Example,

        [XmlEnum("dir")]
        Dir,

        [XmlEnum("concept")]
        Concept,
    }
}
