using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocUrlLink : ITextContainer
    {
        [XmlAttribute("url")]
        public string Url { get; set; } = null!;

        //[XmlText(typeof(string))]
        //[XmlElement(typeof(DocTitleCmdGroup))]
        public List<DocCmdGroup> Parts { get; } = new();
    }
}
