using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocVariableList
    {
        [XmlElement("varlistentry", typeof(DocVarListEntry))]
        [XmlElement("listitem", typeof(DocListItem))]
        public List<object> Parts { get; } = new();
    }
}
