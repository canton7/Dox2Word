using System.Collections.Generic;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class Option
    {
        [XmlAttribute("id")]
        public string Id { get; set; } = null!;

        [XmlAttribute("type")]
        public OptionType Type { get; set; }

        [XmlElement("value")]
        public List<string> Values { get; } = new();
    }
}
