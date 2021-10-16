using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    [XmlRoot("doxygenindex")]
    public class DoxygenIndex
    {
        [XmlElement("compound")]
        public List<Compound> Compounds { get; } = new();
    }
}
