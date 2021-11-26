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
                "programlisting" => Load<Listing>(reader),
                "orderedlist" => Load<DocList>(reader, x => x.Type = ListParagraphType.Number),
                "itemizedlist" => Load<DocList>(reader, x => x.Type = ListParagraphType.Bullet),
                "simplesect" => Load<DocSimpleSect>(reader),
                "table" => Load<DocTable>(reader),
                "dotfile" => Load<DotFile>(reader),
                "parameterlist" => Do(() => this.ParameterLists.Add(Load<DocParamList>(reader))),
                "xrefsect" => Load<DocXRefSect>(reader),
                _ => base.ParseElement(reader),
            };
        }
    }
}