using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class Ref
    {
        [XmlAttribute("refid")]
        public string RefId { get; set; } = null!;

        [XmlText]
        public string Name { get; set; } = null!;
    }
}
