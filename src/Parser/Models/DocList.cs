using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public abstract class DocList
    {
        [XmlElement("listitem")]
        public List<DocListItem> Items { get; } = new();   
    }

    public class OrderedList : DocList { }
    public class UnorderedList : DocList { }
}
