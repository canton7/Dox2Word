using System.Collections.Generic;
using System.Xml.Serialization;
using Dox2Word.Model;

namespace Dox2Word.Parser.Models
{
    public class DocList
    {
        public ListParagraphType Type { get; set; }

        [XmlElement("listitem")]
        public List<DocListItem> Items { get; } = new();   
    }
}
