using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using Dox2Word.Model;

namespace Dox2Word.Parser.Models
{
    public class DocTitleCmdGroup
    {
        public List<object> Parts { get; } = new();

        public XmlSchema? GetSchema() => null;

        public void ReadXml(XmlReader reader)
        {
            reader.ReadStartElement();

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
                        this.ProcessPart(result);
                        reader.Read();
                        break;

                    case XmlNodeType.Text:
                        this.ProcessPart(reader.ReadContentAsString());
                        break;

                    case XmlNodeType.Whitespace:
                        reader.Read(); // TODO
                        break;

                    case XmlNodeType.EndElement:
                        reader.ReadEndElement();
                        return;

                    default:
                        throw new ParserException($"Unexpected XML node type {reader.NodeType}");
                }
            }
        }

        private void ProcessPart(object part)
        {
            this.Parts.Add(part);
        }

        protected virtual object? ParseElement(XmlReader reader)
        {
            if (DocEmptyParser.TryLookup(reader.Name, out string? result))
            {
                return result;
            }

            return reader.Name switch
            {
                "ulink" => Load<DocUrlLink>(reader),
                "bold" => Load<DocMarkup>(reader, x => x.Format = TextRunFormat.Bold),
                "emphasis" => Load<DocMarkup>(reader, x => x.Format = TextRunFormat.Italic),
                "computeroutput" => Load<DocMarkup>(reader, x => x.Format = TextRunFormat.Monospace),
                _ => null,
            };
        }

        protected static T Load<T>(XmlReader reader, Action<T>? configurer = null)
        {
            var instance = (T)new XmlSerializer(typeof(T), new XmlRootAttribute(reader.Name)).Deserialize(reader);
            configurer?.Invoke(instance);
            return instance;
        }

        public void WriteXml(XmlWriter writer) => throw new NotSupportedException();
    }
}
