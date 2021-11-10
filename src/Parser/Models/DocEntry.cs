using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocEntry
    {
        [XmlAttribute("thead")]
        public DoxBool IsTableHead { get; set; }

        [XmlAttribute("colspan")]
        public int ColSpan { get; set; }

        [XmlAttribute("rowspan")]
        public int RowSpan { get; set; }

        [XmlElement("para")]
        public List<DocPara> Paras { get; } = new();
    }
}