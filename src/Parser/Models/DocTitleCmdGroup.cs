using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using Dox2Word.Logging;
using Dox2Word.Model;

namespace Dox2Word.Parser.Models
{
    public abstract class DocTitleCmdGroup : IXmlSerializable
    {
        public List<object> Parts { get; } = new();

        public XmlSchema? GetSchema() => null;

        public void ReadXml(XmlReader reader)
        {
            int startingDepth = reader.Depth;
            this.ReadAttributes(reader);

            // We only care about whitespace when it sits between two child elements
            string? whitespace = null;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    // Whitespace is significant if it comes before a child (e.g. <bold>foo</bold> <bold>bar</bold>)
                    if (whitespace != null)
                    {
                        this.ProcessPart(whitespace);
                    }
                    reader.MoveToContent();
                    var subReader = reader.ReadSubtree();
                    subReader.Read();
                    object? result = this.ParseElement(subReader);
                    if (result != null)
                    {
                        this.ProcessPart(result);
                    }
                }

                whitespace = null;

                switch (reader.NodeType)
                {
                    case XmlNodeType.Text:
                    case XmlNodeType.CDATA:
                    case XmlNodeType.SignificantWhitespace: // Whitespace within xml:space="preserve"
                        this.ProcessPart(reader.Value.TrimStart('\n'));
                        break;

                    case XmlNodeType.Whitespace:
                        whitespace = reader.Value;
                        break;

                    case XmlNodeType.EntityReference:
                        Logger.Instance.Warning($"Unknown entity reference {reader.Name}");
                        break;

                    case XmlNodeType.EndElement when reader.Depth == startingDepth:
                        reader.ReadEndElement();
                        return;

                    default:
                        break;
                }
            }
        }

        private void ProcessPart(object part)
        {
            this.Parts.Add(part);
        }

        protected virtual void ReadAttributes(XmlReader reader) { }

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
                "dot" => Load<Dot>(reader),
                "ref" => Load<Ref>(reader),
                "linebreak" => "\n",
                _ => LoadXmlElement(reader),
            };
        }

        protected static T Load<T>(XmlReader reader, Action<T>? configurer = null)
        {
            var instance = (T)new XmlSerializer(typeof(T), new XmlRootAttribute(reader.Name)).Deserialize(reader);
            configurer?.Invoke(instance);
            return instance;
        }

        protected static XmlElement LoadXmlElement(XmlReader reader)
        {
            var doc = new XmlDocument();
            doc.Load(reader);
            return doc.DocumentElement;
        }

        public static object? Do(Action a)
        {
            a();
            return null;
        }

        public void WriteXml(XmlWriter writer) => throw new NotSupportedException();
    }
}
