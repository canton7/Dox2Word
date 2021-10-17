using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
        [XmlElement("bold", typeof(BoldMarkup))]
        [XmlElement("emphasis", typeof(ItalicMarkup))]
        [XmlElement("computeroutput", typeof(MonospaceMarkup))]
        [XmlAnyElement]
        public List<object> Parts { get; } = new();
    }

    public class BoldMarkup : DocPara { }
    public class ItalicMarkup : DocPara { }
    public class MonospaceMarkup : DocPara { }
}
