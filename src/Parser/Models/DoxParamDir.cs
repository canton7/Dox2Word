using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
