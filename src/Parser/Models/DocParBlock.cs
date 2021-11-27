using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocParBlock
    {
        [XmlElement("para")]
        public List<DocPara> Paras { get; } = new();
    }
}
