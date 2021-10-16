using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    [XmlRoot("doxygen")]
    public class DoxygenFile
    {
        [XmlElement("compounddef")]
        public List<CompoundDef> CompoundDefs { get; } = new();
    }
}
