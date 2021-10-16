using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class LinkedText
    {
        [XmlText(typeof(string))]
        [XmlElement("ref", typeof(RefText))]
        public List<object> Type { get; } = new();
    }
}
