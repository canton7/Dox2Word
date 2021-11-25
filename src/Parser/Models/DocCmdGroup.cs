using System.Xml;

namespace Dox2Word.Parser.Models
{
    public class DocCmdGroup : DocTitleCmdGroup
    {
        protected override object? ParseElement(XmlReader reader)
        {
            return reader.Name switch
            {
                "programlisting" => Load<Listing>(reader),
                "orderedlist" => Load<OrderedList>(reader),
                "itemizedlist" => Load<UnorderedList>(reader),
                _ => base.ParseElement(reader),
            };
        }
    }
}