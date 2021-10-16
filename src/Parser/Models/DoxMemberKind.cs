using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public enum DoxMemberKind
    {
        [XmlEnum("define")]
        Define,

        [XmlEnum("property")]
        Property,

        [XmlEnum("event")]
        Event,

        [XmlEnum("variable")]
        Variable,

        [XmlEnum("typedef")]
        Typedef,

        [XmlEnum("enum")]
        Enum,

        [XmlEnum("function")]
        Function,

        [XmlEnum("signal")]
        Signal,

        [XmlEnum("prototype")]
        Prototype,

        [XmlEnum("friend")]
        Friend,

        [XmlEnum("dcop")]
        Dcop,

        [XmlEnum("slot")]
        Slot,

        [XmlEnum("interface")]
        Interface,

        [XmlEnum("service")]
        Service,
    }
}
