using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
