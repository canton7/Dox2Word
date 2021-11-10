using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocTable
    {
        [XmlAttribute("rows")]
        public int NumRows { get; set; }

        [XmlAttribute("cols")]
        public int NumCols { get; set; }

        [XmlElement("row")]
        public List<DocRow> Rows { get; } = new();
    }
}
