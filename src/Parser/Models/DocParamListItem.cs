using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocParamListItem
    {
        [XmlElement("parameternamelist")]
        public List<DocParamNameList> ParameterNameList { get; } = new();

        [XmlElement("parameterdescription")]
        public Description ParameterDescription { get; set; } = null!;
    }
}
