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
        [XmlText(typeof(string))]
        [XmlElement("parameterlist", typeof(DocParamList))]
        [XmlElement("simplesect", typeof(DocSimpleSect))]
        [XmlAnyElement]
        public List<object> Parts { get; } = new();

        //[XmlAnyElement]
        //public List<XmlElement> Test { get; } = new();
    }
}
