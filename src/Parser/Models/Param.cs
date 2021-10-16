using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class Param
    {
        [XmlElement("type")]
        public LinkedText? Type { get; set; }

        [XmlElement("declname")]
        public string? DeclName { get; set; }
    }
}
