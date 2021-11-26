using System.Xml;

namespace Dox2Word.Parser.Models
{
    public class DocUrlLink : DocTitleCmdGroup
    {
        public string Url { get; set; } = null!;

        protected override void ReadAttributes(XmlReader reader)
        {
            this.Url = reader.GetAttribute("url");
        }
    }
}
