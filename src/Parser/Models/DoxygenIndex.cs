using System.Collections.Generic;
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
