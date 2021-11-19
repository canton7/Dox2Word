using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class SectionDef
    {
        [XmlElement("memberdef")]
        public List<MemberDef> Members { get; } = new();
    }
}
