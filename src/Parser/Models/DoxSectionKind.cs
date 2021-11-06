using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public enum DoxSectionKind
    {
        [XmlEnum("user-defined")]
        UserDefined,

        [XmlEnum("public-type")]
        PublicType,
        
        [XmlEnum("public-func")]
        PublicFunc,

        [XmlEnum("public-attrib")]
        PublicAttrib,

        [XmlEnum("public-slot")]
        PublicSlot,

        [XmlEnum("signal")]
        Signal,

        [XmlEnum("dcop-func")]
        DcopFunc,

        [XmlEnum("property")]
        Property,

        [XmlEnum("event")]
        Event,

        [XmlEnum("public-static-func")]
        PublicStaticFunc,

        [XmlEnum("public-static-attrib")]
        PublicStaticAttrib,

        [XmlEnum("protected-type")]
        ProtectedType,

        [XmlEnum("protected-func")]
        ProtectedFunc,

        [XmlEnum("protected-attrib")]
        ProtectedAttrib,

        [XmlEnum("protected-slot")]
        ProtectedSlot,

        [XmlEnum("protected-static-func")]
        ProtectedStaticFunc,

        [XmlEnum("protected-static-attrib")]
        ProtectedStaticAttrib,

        [XmlEnum("package-type")]
        PackageType,

        [XmlEnum("package-func")]
        PackageFunc,

        [XmlEnum("package-attrib")]
        PackageAttib,

        [XmlEnum("package-static-func")]
        PackageStaticFunc,

        [XmlEnum("package-static-attrib")]
        PackageStaticAttib,

        [XmlEnum("private-type")]
        PrivateType,

        [XmlEnum("private-func")]
        PrivateFunc,

        [XmlEnum("private-attrib")]
        PrivateAttrib,

        [XmlEnum("private-slot")]
        PrivateSlot,

        [XmlEnum("private-static-func")]
        PrivateStaticFunc,

        [XmlEnum("private-static-attrib")]
        PrivateStaticAttrib,

        [XmlEnum("friend")]
        Friend,
        
        [XmlEnum("related")]
        Related,

        [XmlEnum("define")]
        Define,

        [XmlEnum("prototype")]
        Prototype,

        [XmlEnum("typedef")]
        Typedef,

        [XmlEnum("enum")]
        Enum,

        [XmlEnum("func")]
        Func,

        [XmlEnum("var")]
        Var,
    }
}
