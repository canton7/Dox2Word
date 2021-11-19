using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocPara
    {
        [XmlElement("parameterlist")]
        public List<DocParamList> ParameterLists { get; } = new();

        [XmlText(typeof(string))]
        [XmlElement("simplesect", typeof(DocSimpleSect))]
        [XmlElement("orderedlist", typeof(OrderedList))]
        [XmlElement("itemizedlist", typeof(UnorderedList))]
        [XmlElement("xrefsect", typeof(DocXRefSect))]
        [XmlElement("programlisting", typeof(Listing))]
        [XmlElement("dot", typeof(Dot))]
        [XmlElement("dotfile", typeof(DotFile))]
        [XmlElement("bold", typeof(BoldMarkup))]
        [XmlElement("emphasis", typeof(ItalicMarkup))]
        [XmlElement("computeroutput", typeof(MonospaceMarkup))]
        [XmlElement("ref", typeof(Ref))]
        [XmlElement("table", typeof(DocTable))]
        [XmlAnyElement]
        public List<object> Parts { get; } = new();
    }

    public class BoldMarkup : DocPara { }
    public class ItalicMarkup : DocPara { }
    public class MonospaceMarkup : DocPara { }
}
