using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace Dox2Word.Parser.Models
{
    public class DocPara : IXmlSerializable
    {
        //[XmlElement("parameterlist")]
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
        public List<object> Parts { get; } = new();

        public XmlSchema? GetSchema() => null;

        public void ReadXml(XmlReader reader)
        {
            reader.ReadStartElement("para");

            while (true)
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:
                        reader.MoveToContent();
                        var subReader = reader.ReadSubtree();
                        subReader.Read();
                        object? result = this.ParseElement(subReader);
                        if (result == null)
                        {
                            var doc = new XmlDocument();
                            doc.Load(subReader);
                            result = doc.DocumentElement;
                        }
                        this.Parts.Add(result);
                        reader.Read();
                        break;

                    case XmlNodeType.EndElement:
                        if (reader.Name == "para")
                        {
                            reader.ReadEndElement();
                            return;
                        }
                        break;

                    case XmlNodeType.Text:
                        this.Parts.Add(reader.ReadContentAsString());
                        break;

                    case XmlNodeType.Whitespace:
                        reader.Read();
                        break;
                }
            }
        }

        protected virtual object? ParseElement(XmlReader reader)
        {
            return reader.Name switch
            {
                "para" => Load<DocPara>(reader),
                "ref" => Load<Ref>(reader),
                _ => null,
            };
        }

        protected static T Load<T>(XmlReader reader)
        {
            return (T)new XmlSerializer(typeof(T), new XmlRootAttribute(reader.Name)).Deserialize(reader);
        }

        public void WriteXml(XmlWriter writer) => throw new NotSupportedException();
    }

    public class BoldMarkup : DocPara { }
    public class ItalicMarkup : DocPara { }
    public class MonospaceMarkup : DocPara { }
}
