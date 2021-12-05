using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocVarListEntry
    {
        [XmlElement("term")]
        public DocTitle Term { get; set; } = null!;
    }
}
