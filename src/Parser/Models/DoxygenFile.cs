using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    [XmlRoot("doxyfile")]
    public class DoxygenFile
    {
        [XmlElement("option")]
        public List<Option> Options { get; } = new();
    }
}
