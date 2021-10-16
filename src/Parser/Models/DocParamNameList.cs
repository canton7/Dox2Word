using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocParamNameList
    {
        // TODO: Not actually strictly a string I think?
        [XmlElement("parametername")]
        public string ParameterName { get; set; } = null!;
    }
}