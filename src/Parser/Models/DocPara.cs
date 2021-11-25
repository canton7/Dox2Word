using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocPara : DocCmdGroup
    {
        public List<DocParamList> ParameterLists { get; } = new();

        //[XmlText(typeof(string))]
        //[XmlElement("ulink", typeof(DocUrlLink))]
        //[XmlElement("simplesect", typeof(DocSimpleSect))]
        //[XmlElement("orderedlist", typeof(OrderedList))]
        //[XmlElement("itemizedlist", typeof(UnorderedList))]
        //[XmlElement("xrefsect", typeof(DocXRefSect))]
        //[XmlElement("programlisting", typeof(Listing))]
        //[XmlElement("dot", typeof(Dot))]
        //[XmlElement("dotfile", typeof(DotFile))]
        //[XmlElement("bold", typeof(BoldMarkup))]
        //[XmlElement("emphasis", typeof(ItalicMarkup))]
        //[XmlElement("computeroutput", typeof(MonospaceMarkup))]
        //[XmlElement("ref", typeof(Ref))]
        //[XmlElement("table", typeof(DocTable))]
        //[XmlAnyElement]
    }
}
