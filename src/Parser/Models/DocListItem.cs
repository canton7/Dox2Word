using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocListItem
    {
        [XmlElement("para")]
        public List<DocPara> Paras { get; } = new();
    }
}
