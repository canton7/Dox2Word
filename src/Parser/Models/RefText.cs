using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class RefText
    {
        [XmlAttribute("refid")]
        public string RefId { get; set; } = null!;

        [XmlAttribute("kindref")]
        public string KindRef { get; set; } = null!;

        [XmlText]
        public string Name { get; set; } = null!;
    }
}
