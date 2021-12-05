using System.Collections.Generic;
using System.Xml;
using Dox2Word.Model;

namespace Dox2Word.Parser.Models
{
    public abstract class DocCmdGroup : DocTitleCmdGroup
    {
        public List<DocParamList> ParameterLists { get; } = new();

        protected override object? ParseElement(XmlReader reader)
        {
            return reader.Name switch
            {
                "hruler" => new HorizontalRulerDocEmptyType(),
                "preformatted" => Load<PreformattedDocMarkup>(reader),
                "programlisting" => Load<Listing>(reader),
                "verbatim" => Load<PreformattedDocMarkup>(reader),
                "indexentry" => Unsupported("indexentry"),
                "orderedlist" => Load<DocList>(reader, x => x.Type = ListParagraphType.Number),
                "itemizedlist" => Load<DocList>(reader, x => x.Type = ListParagraphType.Bullet),
                "simplesect" => Load<DocSimpleSect>(reader),
                "title" => Unsupported("title"), // Used as part of pages
                "variablelist" => Load<DocVariableList>(reader),
                "table" => Load<DocTable>(reader),
                "heading" => Unsupported("heading"), // Used as part of pages
                "dotfile" => Load<DotFile>(reader),
                "mscfile" => Unsupported("mscfile"),
                "diafile" => Unsupported("diafile"),
                "toclist" => Unsupported("toclist"),
                "language" => Unsupported("language"),
                "parameterlist" => Do(() => this.ParameterLists.Add(Load<DocParamList>(reader))),
                "xrefsect" => Load<DocXRefSect>(reader),
                "copydoc" => Unsupported("copydoc"), // I don't know how to cause this. Doxygen's testing doesn't use it
                "blockquote" => Load<DocBlockQuote>(reader),
                "parblock" => Load<DocParBlock>(reader),
                _ => base.ParseElement(reader),
            };
        }
    }
}