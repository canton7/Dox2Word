using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocParamNameList
    {
        [XmlElement("parametername")]
        public DocParamName ParameterName { get; set; } = null!;
    }
}