using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    [XmlRoot("doxygen")]
    public class Doxygen
    {
        [XmlElement("compounddef")]
        public List<CompoundDef> CompoundDefs { get; } = new();
    }
}
