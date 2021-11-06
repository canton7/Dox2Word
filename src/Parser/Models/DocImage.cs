using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class Dot
    {
        [XmlText]
        public string Contents { get; set; } = null!;
    }
}
