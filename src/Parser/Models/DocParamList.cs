using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocParamList
    {
        [XmlAttribute("kind")]
        public DoxParamListKind Kind { get; set; }

        [XmlElement("parameteritem")]
        public List<DocParamListItem> ParameterItems { get; } = new();
    }
}
