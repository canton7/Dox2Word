using System.Collections.Generic;
using System.Xml.Serialization;
using Dox2Word.Model;

namespace Dox2Word.Parser.Models
{
    public class DocMarkup : ITextContainer
    {
        public TextRunFormat Format { get; set; }

        //[XmlText(typeof(string))]
        //[XmlElement(typeof(DocCmdGroup))]
        [XmlText]
        public List<DocCmdGroup> Parts { get; } = new();
    }
}
