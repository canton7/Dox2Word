using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class Highlight
    {
        // There's a lot of complexity here... And we're going to ignore all of it, and just pull out strings
        [XmlText(typeof(string))]
        [XmlElement("sp", typeof(Sp))]
        [XmlAnyElement]
        public List<object> Parts { get; } = new();
    }
}
