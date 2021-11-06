using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public enum DoxParamListKind
    {
        [XmlEnum("param")]
        Param,

        [XmlEnum("retval")]
        RetVal,

        [XmlEnum("exception")]
        Exception,

        [XmlEnum("templateparam")]
        TemplateParam,
    }
}
