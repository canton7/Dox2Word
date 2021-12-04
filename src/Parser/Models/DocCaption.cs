using System.Xml;

namespace Dox2Word.Parser.Models
{
    public class DocCaption : DocTitleCmdGroup
    {
        public string? Id { get; set; }

        protected override void ReadAttributes(XmlReader reader)
        {
            this.Id = reader.GetAttribute("Id");
        }
    }
}