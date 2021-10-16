using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocSimpleSect
    {
        [XmlAttribute("kind")]
        public DoxSimpleSectKind Kind { get; set; }

        [XmlElement("para")]
        public DocPara? Para { get; set; }
    }
}